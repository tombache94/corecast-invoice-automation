#!/usr/bin/env node
/**
 * Daily orchestrator — runs at 5:50am via Windows Task Scheduler.
 *
 * Replaces Cowork's `corecast-morning-sync` task. Chains:
 *   1. Load state (last-run timestamp, processed message IDs)
 *   2. Fetch active jobs from Monday (board 5027051539)
 *   3. Scan Outlook inboxes (tom@ + accounts@) for emails since last run
 *   4. Classify each email → active job (rule-based, config/job-routing.json)
 *   5. Format Cowork-style notes
 *   6. Dry-run: write a JSON dump to data/orchestrator-dryrun/
 *      Live:    append notes to each job's Notes long-text column on Monday
 *   7. Call file-attachments.js (always honours its own --dry-run flag)
 *   8. Call refresh-dashboard.js (skipped in dry-run mode)
 *   9. Save state
 *
 * Usage:
 *   node scripts/daily-orchestrator.js                 # live
 *   node scripts/daily-orchestrator.js --dry-run       # no Monday writes, no dashboard refresh
 *   node scripts/daily-orchestrator.js --lookback 48   # override lookback (default 26h)
 *   node scripts/daily-orchestrator.js --no-dashboard  # skip refresh chain (testing only)
 */

require('dotenv').config();
const path = require('path');
const fs = require('fs');
const { spawnSync } = require('child_process');

const { OutlookClient } = require('../src/email/outlook');
const { fetchJobs } = require('../src/dashboard/sources/monday');
const { loadRouting } = require('../src/orchestrator/routing');
const { classifyEmail } = require('../src/orchestrator/classify');
const { formatNote, appendNotesToJob } = require('../src/orchestrator/notes');
const { pollOverdueInvoices, formatOverdueNote } = require('../src/orchestrator/xeroOverdue');
const { ingest } = require('../src/index');

const STATE_FILE = path.resolve(__dirname, '../data/orchestrator-state.json');
const DRYRUN_DIR = path.resolve(__dirname, '../data/orchestrator-dryrun');
const MAILBOXES = ['tom@corecastconcrete.com.au', 'accounts@corecastconcrete.com.au'];
const DEFAULT_LOOKBACK_HOURS = 26;

function parseArgs(argv) {
  const args = { dryRun: false, lookbackHours: null, withDashboard: true };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--dry-run') args.dryRun = true;
    else if (a === '--no-dashboard') args.withDashboard = false;
    else if (a === '--lookback') args.lookbackHours = Number(argv[++i]);
  }
  if (args.dryRun) args.withDashboard = false;
  return args;
}

function loadState() {
  const empty = { lastRunISO: null, processedMessageIds: [], notedOverdueInvoiceIds: [] };
  if (!fs.existsSync(STATE_FILE)) return empty;
  try {
    const parsed = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
    return {
      lastRunISO: parsed.lastRunISO || null,
      processedMessageIds: Array.isArray(parsed.processedMessageIds) ? parsed.processedMessageIds : [],
      notedOverdueInvoiceIds: Array.isArray(parsed.notedOverdueInvoiceIds) ? parsed.notedOverdueInvoiceIds : [],
    };
  } catch {
    return empty;
  }
}

function saveState(state) {
  fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
  fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const startedAt = new Date();
  const startedISO = startedAt.toISOString();

  console.log('🏗️  CoreCast — Daily Orchestrator');
  console.log(`   Mode   : ${args.dryRun ? '🔍 DRY RUN (no Monday writes, no dashboard refresh)' : '✏️  LIVE'}`);
  console.log(`   Started: ${startedISO}\n`);

  // 1. State + lookback window
  const state = loadState();
  const lookbackHours = args.lookbackHours ?? DEFAULT_LOOKBACK_HOURS;
  const fallbackSince = new Date(startedAt.getTime() - lookbackHours * 3600 * 1000).toISOString();
  const sinceISO = state.lastRunISO || fallbackSince;
  console.log(`   Since  : ${sinceISO}`);
  console.log(`   Seen IDs: ${state.processedMessageIds.length} (deduped)`);

  // 2a. Invoice ingestion — runs in both dry-run and live mode because it's
  // pure read-from-Outlook + write-to-local-store. Dedup by internetMessageId
  // means re-runs (and the same email cc'd across mailboxes) are idempotent.
  // Pulls last 7 days from accounts@, tom@, and josh@ — subcontractors send
  // invoices to any of those addresses. Without this, the weekly Tuesday
  // match task would never see new invoices (manual ingestion was the only path).
  const INVOICE_MAILBOXES = [
    'accounts@corecastconcrete.com.au',
    'tom@corecastconcrete.com.au',
    'josh@corecastconcrete.com.au',
  ];
  console.log(`\n📥 Ingesting new invoices from ${INVOICE_MAILBOXES.length} mailboxes…`);
  const totals = { total: 0, filtered: 0, duplicates: 0, stored: 0, review_needed: 0, pending: 0 };
  for (const mb of INVOICE_MAILBOXES) {
    try {
      const s = await ingest({ mailbox: mb, limit: 100, sinceDays: 7 });
      if (s) {
        for (const k of Object.keys(totals)) totals[k] += s[k] || 0;
        console.log(`   ${mb}: scanned=${s.total}, matched=${s.filtered}, dup=${s.duplicates}, new=${s.stored}`);
      }
    } catch (err) {
      console.log(`   ⚠️  ${mb} failed (continuing): ${err.message}`);
    }
  }
  console.log(`   TOTAL across mailboxes: scanned=${totals.total}, matched=${totals.filtered}, skipped-dup=${totals.duplicates}, stored=${totals.stored} (review_needed=${totals.review_needed}, pending=${totals.pending})`);

  // 2. Load active jobs
  console.log('\n📋 Loading active jobs from Monday…');
  let manualAllocations = {};
  try {
    manualAllocations = require('../src/dashboard/manualAllocations.json');
  } catch (_) {}
  const { jobs } = await fetchJobs(manualAllocations);
  const activeJobs = jobs.filter((j) => j.stage === 'In Progress' || j.stage === 'Not begun');
  console.log(`   ${activeJobs.length} active: ${activeJobs.map((j) => j.name).join(' · ')}`);

  // 3. Outlook scan
  console.log('\n📨 Scanning Outlook…');
  const outlook = new OutlookClient({
    tenantId: process.env.MS_TENANT_ID,
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
  });
  const seenIds = new Set(state.processedMessageIds);
  const allMessages = [];
  const mailboxHealth = [];
  for (const mb of MAILBOXES) {
    try {
      const msgs = await outlook.listRecentMessages(mb, { sinceISO, top: 100 });
      const fresh = msgs.filter((m) => !seenIds.has(m.id));
      console.log(`   ${mb}: ${msgs.length} fetched, ${fresh.length} new (rest already processed)`);
      mailboxHealth.push({ mailbox: mb, ok: true, fetched: msgs.length, fresh: fresh.length });
      allMessages.push(...fresh);
    } catch (err) {
      console.log(`   ${mb}: ⚠️  failed — ${err.message}`);
      mailboxHealth.push({ mailbox: mb, ok: false, error: err.message });
    }
  }

  // 4. Classify
  console.log('\n🏷️  Classifying…');
  const routing = loadRouting();
  const byJob = new Map(); // jobId → { job, hits[] }
  const unclassified = [];
  for (const msg of allMessages) {
    const match = classifyEmail(msg, activeJobs, routing);
    if (match) {
      const bucket = byJob.get(match.jobId) || { job: activeJobs.find((j) => j.id === match.jobId), hits: [] };
      bucket.hits.push({ msg, match });
      byJob.set(match.jobId, bucket);
    } else {
      unclassified.push(msg);
    }
  }
  for (const { job, hits } of byJob.values()) {
    console.log(`   ${job.name}: ${hits.length} email(s)`);
  }
  if (unclassified.length) console.log(`   Unclassified: ${unclassified.length}`);

  // 5. Build email-derived notes
  const updatesPerJob = [];
  for (const { job, hits } of byJob.values()) {
    const newEntries = hits.map((h) => formatNote(h.msg));
    updatesPerJob.push({
      jobId: job.id,
      jobName: job.name,
      currentNotesPreview: (job.notes || '').slice(0, 200),
      newEntries,
      sourceMessageIds: hits.map((h) => h.msg.id),
      xeroOverdueInvoiceIds: [],
    });
  }

  // 5b. Xero overdue receivables poll
  // Per Tom's rule: match invoices to jobs via the Reference field only,
  // never by client name (Macrete/Pennant are shared across jobs).
  console.log('\n💵 Polling Xero for overdue receivables…');
  let xeroOverdueSummary = { polled: 0, newlyNoted: 0, alreadyNoted: 0, unassigned: 0 };
  const unassignedXeroOverdue = [];
  const newlyNotedInvoiceIds = [];
  try {
    const overdueResults = await pollOverdueInvoices(activeJobs);
    const notedIds = new Set(state.notedOverdueInvoiceIds || []);
    xeroOverdueSummary.polled = overdueResults.length;

    for (const r of overdueResults) {
      if (r.reason !== 'matched') {
        xeroOverdueSummary.unassigned++;
        unassignedXeroOverdue.push({
          invoiceNumber: r.invoice.invoiceNumber,
          client: r.invoice.client,
          reference: r.invoice.reference,
          amountDue: r.invoice.amountDue,
          daysOverdue: r.invoice.daysOverdue,
          reason: r.reason,
        });
        continue;
      }
      if (notedIds.has(r.invoice.invoiceId)) {
        xeroOverdueSummary.alreadyNoted++;
        continue;
      }
      const note = formatOverdueNote(r.invoice);
      const existing = updatesPerJob.find((u) => u.jobId === r.match.id);
      if (existing) {
        existing.newEntries.push(note);
        existing.xeroOverdueInvoiceIds.push(r.invoice.invoiceId);
      } else {
        updatesPerJob.push({
          jobId: r.match.id,
          jobName: r.match.name,
          currentNotesPreview: (r.match.notes || '').slice(0, 200),
          newEntries: [note],
          sourceMessageIds: [],
          xeroOverdueInvoiceIds: [r.invoice.invoiceId],
        });
      }
      newlyNotedInvoiceIds.push(r.invoice.invoiceId);
      xeroOverdueSummary.newlyNoted++;
    }
    console.log(`   Polled: ${xeroOverdueSummary.polled} overdue · New: ${xeroOverdueSummary.newlyNoted} · Already noted: ${xeroOverdueSummary.alreadyNoted} · Unassigned: ${xeroOverdueSummary.unassigned}`);
  } catch (err) {
    console.log(`   ⚠️  Xero poll failed (continuing): ${err.message}`);
    xeroOverdueSummary.error = err.message;
  }

  // 6. Output
  if (args.dryRun) {
    fs.mkdirSync(DRYRUN_DIR, { recursive: true });
    const stamp = startedISO.replace(/[:.]/g, '-');
    const outFile = path.join(DRYRUN_DIR, `dryrun-${stamp}.json`);
    fs.writeFileSync(
      outFile,
      JSON.stringify(
        {
          startedAt: startedISO,
          mode: 'dry-run',
          sinceISO,
          lookbackHours,
          activeJobsCount: activeJobs.length,
          mailboxHealth,
          totalEmailsConsidered: allMessages.length,
          updatesPerJob,
          xeroOverdueSummary,
          unassignedXeroOverdue,
          unclassified: unclassified.map((m) => ({
            mailbox: m.mailbox,
            from: m.from,
            fromName: m.fromName,
            subject: m.subject,
            receivedAt: m.receivedAt,
          })),
        },
        null,
        2,
      ),
    );
    console.log(`\n🔍 Dry run output → ${outFile}`);
    console.log(`   Would have written notes to ${updatesPerJob.length} job(s).`);
  } else {
    console.log('\n✏️  Writing notes to Monday…');
    for (const upd of updatesPerJob) {
      const job = activeJobs.find((j) => j.id === upd.jobId);
      try {
        await appendNotesToJob(upd.jobId, job?.notes || '', upd.newEntries);
        console.log(`   ✅ ${upd.jobName}: appended ${upd.newEntries.length} note(s)`);
      } catch (err) {
        console.log(`   ⚠️  ${upd.jobName}: write failed — ${err.message}`);
      }
    }
  }

  // 7. Update state (always — even dry-run — so we don't re-fetch the same
  //    emails forever; but in dry-run we DON'T mark them as processed so
  //    Cowork's parallel run sees the same emails too. Just bump lastRunISO
  //    in live mode.
  if (!args.dryRun) {
    const newProcessedIds = allMessages.map((m) => m.id);
    const keepRecent = (state.processedMessageIds || []).slice(-2000); // cap memory
    state.processedMessageIds = [...keepRecent, ...newProcessedIds];
    state.notedOverdueInvoiceIds = [...(state.notedOverdueInvoiceIds || []), ...newlyNotedInvoiceIds];
    state.lastRunISO = startedISO;
    saveState(state);
    console.log(`\n💾 State saved (lastRunISO=${state.lastRunISO}, ${state.processedMessageIds.length} msg ids · ${state.notedOverdueInvoiceIds.length} xero overdue ids)`);
  }

  // 8. Xero push + attach (live mode only)
  // Push all pending invoices to Xero as draft bills, then upload the PDF
  // attachments from Outlook. Both scripts are idempotent — duplicates are
  // caught by Xero's bill lookup and attachment dedup by filename.
  if (!args.dryRun) {
    console.log('\n💳 Pushing invoices to Xero…');
    const pushResult = spawnSync('node', [path.resolve(__dirname, 'push-to-xero.js')], {
      stdio: 'inherit',
      shell: false,
    });
    if (pushResult.status !== 0) {
      console.log(`   ⚠️  push-to-xero.js exited with code ${pushResult.status}`);
    }

    console.log('\n📎 Attaching PDFs to Xero bills…');
    const attachResult = spawnSync('node', [path.resolve(__dirname, 'attach-invoices-to-xero.js')], {
      stdio: 'inherit',
      shell: false,
    });
    if (attachResult.status !== 0) {
      console.log(`   ⚠️  attach-invoices-to-xero.js exited with code ${attachResult.status}`);
    }
  } else {
    console.log('\n⏭   Skipped Xero push + attach (dry-run)');
  }

  // 9. Chain: dashboard refresh (live mode only)
  if (args.withDashboard) {
    console.log('\n🔄 Calling refresh-dashboard.js…');
    const result = spawnSync('node', [path.resolve(__dirname, 'refresh-dashboard.js')], {
      stdio: 'inherit',
      shell: false,
    });
    if (result.status !== 0) {
      console.log(`   ⚠️  refresh-dashboard.js exited with code ${result.status}`);
    }
  } else {
    console.log('\n⏭   Skipped dashboard refresh (dry-run or --no-dashboard)');
  }

  console.log('\n✅ Orchestrator finished');
}

main().catch((err) => {
  console.error('Fatal:', err.message);
  if (process.env.DEBUG) console.error(err.stack);
  process.exit(1);
});
