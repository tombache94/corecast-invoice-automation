/**
 * Combine outputs from Monday, Xero (bills + receivables), SharePoint, and the
 * manualAllocations config into a single DASHBOARD_DATA blob ready for the
 * dashboard's render functions.
 */

const fs = require('fs');
const path = require('path');
const { fetchJobs } = require('./sources/monday');
const { fetchBillsByJob } = require('./sources/xeroBills');
const { fetchSpendTransactions } = require('./sources/xeroBankTransactions');
const { fetchReceivables } = require('./sources/xeroReceivables');
const { fetchQuotes } = require('./sources/sharepoint');
const { listJobAndQuoteFolders } = require('../sharepoint/folders');
const { matchFolder } = require('../sharepoint/jobLinkSync');
const { fetchLatestActivityPerJob } = require('./sources/latestActivity');

const MANUAL_PATH = path.resolve(__dirname, 'manualAllocations.json');

// Build a dynamic job-text matcher from the live SharePoint folder list.
// Replaces the previous hardcoded JOB_KEYWORDS table — any folder that
// exists under ACTIVE JOBS or Active Quotes is automatically a keyword,
// so new jobs are picked up immediately with no code change.
//
// `folderToJobKey` maps each SharePoint folder name back to the canonical
// Monday jobKey when one matches via the existing matchFolder() logic.
// Folders without a Monday peer fall back to using the folder name itself
// as the bucket key (mainly: quote folders for jobs not yet started).
//
// Matching strategy: longest-substring-wins, using both the full folder
// name AND dash-separated parts ≥4 chars, so a reference like
// "Baldivis | Materials" matches "Focus - Parkland Heights - Baldivis"
// via the "baldivis" part.
function buildJobDetector(folders, folderToJobKey) {
  // Pre-compute candidates per folder once.
  const folderCandidates = folders.map((f) => {
    const lower = f.name.toLowerCase();
    const parts = lower.split(/\s*[-–—|]\s*/).map((s) => s.trim()).filter((s) => s.length >= 4);
    const unique = Array.from(new Set([lower, ...parts]));
    return { folderName: f.name, candidates: unique };
  });

  return function detectJobFromText(...texts) {
    const text = texts.filter(Boolean).join(' ').toLowerCase();
    if (!text) return null;
    let best = null;
    for (const fc of folderCandidates) {
      for (const cand of fc.candidates) {
        if (text.includes(cand)) {
          // Longest match wins — "parkland heights" (16) beats "focus" (5)
          // if both appear in the same reference.
          if (!best || cand.length > best.length) {
            best = { folderName: fc.folderName, length: cand.length };
          }
        }
      }
    }
    if (!best) return null;
    return folderToJobKey[best.folderName] || best.folderName;
  };
}

async function compose({ dataAsAt = null } = {}) {
  const manual = JSON.parse(fs.readFileSync(MANUAL_PATH, 'utf8'));

  const [mondayResult, billsResult, spendResult, receivablesResult, quotesResult, foldersResult] = await Promise.all([
    fetchJobs(manual).catch((err) => ({ jobs: [], error: err.message })),
    fetchBillsByJob().catch((err) => ({ byJob: {}, untagged: [], totalBills: 0, error: err.message })),
    fetchSpendTransactions().catch((err) => ({ transactions: [], totalTransactions: 0, error: err.message })),
    fetchReceivables().catch((err) => ({ invoices: [], totalOutstanding: 0, totalOverdue: 0, overdueCount: 0, error: err.message })),
    fetchQuotes().catch((err) => ({ quotes: [], available: false, reason: err.message })),
    listJobAndQuoteFolders().catch((err) => ({ folders: [], status: { jobs: { ok: false, error: err.message }, quotes: { ok: false, error: err.message } } })),
  ]);

  const jobs = mondayResult.jobs || [];
  const billsByJob = billsResult.byJob || {};
  const receivables = receivablesResult.invoices || [];
  const folders = foldersResult.folders || [];

  // Build the folder → canonical jobKey reverse map for receivables
  // attribution. Folders that don't match a Monday job (e.g. quote-only
  // folders) bucket under their folder name instead.
  const folderToJobKey = {};
  for (const j of jobs) {
    const f = matchFolder(j.jobKey, folders);
    if (f) folderToJobKey[f.name] = j.jobKey;
  }
  const detectJobFromText = buildJobDetector(folders, folderToJobKey);

  // Attribute SPEND bank transactions to jobs by matching the Reference
  // against currently-active Monday jobs. We accept any match that includes
  // the job's jobKey or full name as a substring (case-insensitive).
  // "Active" here = stage In Progress or Not begun. Spend that doesn't match
  // an active job is dropped from the per-job aggregation (untagged).
  const activeJobs = jobs.filter((j) => j.stage === 'In Progress' || j.stage === 'Not begun');
  const activePatterns = activeJobs.map((j) => ({
    jobKey: j.jobKey,
    needles: Array.from(new Set([
      (j.jobKey || '').toLowerCase().trim(),
      (j.name   || '').toLowerCase().trim(),
    ].filter(Boolean))),
  }));
  function matchSpendToActiveJob(reference) {
    if (!reference) return null;
    const ref = String(reference).toLowerCase();
    let best = null;
    for (const ap of activePatterns) {
      for (const needle of ap.needles) {
        if (ref.includes(needle)) {
          // Prefer the longest needle match — "parkland heights - baldivis"
          // beats just "baldivis" if both job patterns matched.
          if (!best || needle.length > best.matchedLength) {
            best = { jobKey: ap.jobKey, matchedLength: needle.length };
          }
        }
      }
    }
    return best ? best.jobKey : null;
  }

  // Fallback attribution for bills with no Job tracking category: try the
  // bill-level Reference field against active Monday jobs. Catches bills
  // (especially DRAFTs) where Tom set the Reference but didn't set the Job
  // tracking category. Mirrors how spend transactions are attributed below.
  for (const entry of (billsResult.untagged || [])) {
    const job = matchSpendToActiveJob(entry.billReference);
    if (!job) continue;
    const slot = billsByJob[job] ||
      (billsByJob[job] = { total: 0, labour: 0, materials: 0, lines: [] });
    slot.total = round2((slot.total || 0) + entry.amount);
    if (entry.labour) {
      slot.labour = round2((slot.labour || 0) + entry.amount);
    } else {
      slot.materials = round2((slot.materials || 0) + entry.amount);
    }
    slot.lines.push(entry);
  }

  for (const t of (spendResult.transactions || [])) {
    const job = matchSpendToActiveJob(t.reference);
    if (!job) continue;
    const slot = billsByJob[job] ||
      (billsByJob[job] = { total: 0, labour: 0, materials: 0, labourHours: 0, lines: [] });
    slot.spend = round2((slot.spend || 0) + t.amount);
    slot.total = round2((slot.total || 0) + t.amount);
    slot.lines.push(t);
  }

  // Seed each Xero-sourced job slot with a rate-derived labour-hours estimate
  // so manual add-ons can append explicit hours on top.
  const hourlyRateInit = Number(manual.labourHourlyRate || 0) || null;
  for (const slot of Object.values(billsByJob)) {
    slot.labourHours = hourlyRateInit ? slot.labour / hourlyRateInit : 0;
  }

  // Manual labour add-ons fold into the relevant job's labour bucket.
  // - When `xeroContacts` is set, any tagged Xero lines from those contacts
  //   on this job are subtracted first so the manual entry replaces (rather
  //   than supplements) the Xero contribution. Use this for cases where a
  //   single Xero bill covers multiple jobs and only one job's portion is
  //   being manually split out.
  // - When `xeroContacts` is absent, the entry is purely additive (use this
  //   for labour that isn't in Xero at all, e.g. Jack Henderson).
  // - Use explicit `hours` when provided; otherwise fall back to amount / rate.
  for (const addon of manual.labourAddOns || []) {
    const slot = billsByJob[addon.job] ||
      (billsByJob[addon.job] = { total: 0, labour: 0, materials: 0, labourHours: 0, lines: [] });

    // Subtract Xero contributions from the named contacts before adding the manual.
    const excludeContacts = (addon.xeroContacts || []).map((s) => String(s).toLowerCase().trim());
    if (excludeContacts.length) {
      const remainingLines = [];
      for (const line of slot.lines) {
        const c = String(line.contact || '').toLowerCase().trim();
        if (excludeContacts.includes(c)) {
          slot.total = round2(slot.total - line.amount);
          if (line.labour) {
            slot.labour = round2(slot.labour - line.amount);
            slot.labourHours = (slot.labourHours || 0) - (hourlyRateInit ? line.amount / hourlyRateInit : 0);
          } else {
            slot.materials = round2(slot.materials - line.amount);
          }
        } else {
          remainingLines.push(line);
        }
      }
      slot.lines = remainingLines;
    }

    const amount = Number(addon.amount || 0);
    slot.total = round2(slot.total + amount);
    slot.labour = round2(slot.labour + amount);
    const hours = addon.hours != null
      ? Number(addon.hours)
      : (hourlyRateInit ? amount / hourlyRateInit : 0);
    slot.labourHours = (slot.labourHours || 0) + hours;
    slot.lines.push({
      billId: null,
      billNumber: 'manual',
      contact: addon.person,
      date: null,
      status: 'MANUAL',
      accountCode: '',
      accountName: 'Manual labour allocation',
      description: addon.note || '',
      amount,
      hours: addon.hours != null ? Number(addon.hours) : null,
      labour: true,
      replacedXeroContacts: excludeContacts.length ? addon.xeroContacts : undefined,
    });
  }

  // Attribute receivables to jobs by the Reference field ONLY. Macrete (and
  // others) bill us across multiple jobs, so client name is ambiguous; and
  // descriptions are too freeform to trust. The Reference field is the
  // canonical job tag — if it doesn't name a job, the invoice goes to
  // unattributed and Tom updates Xero rather than us guessing.
  const receivablesByJob = {};
  const unattributedReceivables = [];
  for (const r of receivables) {
    const job = detectJobFromText(r.reference);
    if (job) (receivablesByJob[job] ??= []).push(r);
    else unattributedReceivables.push(r);
  }

  const hourlyRate = Number(manual.labourHourlyRate || 0) || null;
  // Scan Outlook for the most recent classified email per active job, so
  // the dashboard's "Latest" field always reflects current activity (not
  // whatever was last hand-written into the Monday Notes column).
  const activeJobsForLatest = jobs.filter((j) => j.stage === 'In Progress' || j.stage === 'Not begun');
  let latestActivityByJobId = {};
  try {
    latestActivityByJobId = await fetchLatestActivityPerJob(activeJobsForLatest);
  } catch (_) {
    // best-effort — fall back to notes-first-line on the render side
  }

  const enrichedJobs = jobs.map((j) => {
    const costs = billsByJob[j.jobKey] || { total: 0, labour: 0, materials: 0, labourHours: 0, lines: [] };
    // Prefer the additive labourHours we built up (rate-derived for Xero +
    // explicit-where-given for manual addons). Fall back to the rate calc.
    const labourHours = costs.labourHours != null
      ? Math.round(costs.labourHours * 10) / 10
      : (hourlyRate ? Math.round((costs.labour / hourlyRate) * 10) / 10 : null);
    const budget = (manual.jobBudgets || {})[j.jobKey] || null;
    const labourPctOfBudget = budget?.labour ? Math.round((costs.labour / budget.labour) * 100) : null;
    const totalPctOfBudget = budget?.total ? Math.round((costs.total / budget.total) * 100) : null;
    // "Projected margin" = expected margin at job completion. Prefer the
    // budgeted total cost when we have one (manualAllocations.jobBudgets);
    // otherwise fall back to current actual costs as a lower bound.
    const projectedTotalCost = budget?.total ?? costs.total;
    const projectedMargin = j.contractValue
      ? {
          amount: round2(j.contractValue - projectedTotalCost),
          pct: Math.round(((j.contractValue - projectedTotalCost) / j.contractValue) * 100),
          basis: budget?.total ? 'budget' : 'costs-to-date',
        }
      : null;
    const jobReceivables = receivablesByJob[j.jobKey] || [];
    const overdueAmount = round2(
      jobReceivables.filter((r) => (r.daysOverdue ?? 0) > 0).reduce((s, r) => s + r.amountDue, 0),
    );
    // Derive client from the job's first attributed receivable. Same
    // source the Cashflow tab uses, so client names stay consistent across
    // tabs. Falls back to whatever the Monday job source provided (none
    // currently — but future-proofed if a client column gets added there).
    const client = (jobReceivables[0] && jobReceivables[0].client) || j.client || null;
    return {
      ...j,
      client,
      costs: { total: costs.total, labour: costs.labour, materials: costs.materials },
      labourHours,
      labourPctOfBudget,
      totalPctOfBudget,
      budget,
      projectedMargin,
      receivables: jobReceivables,
      overdueAmount,
      latestActivity: latestActivityByJobId[j.id] || null,
    };
  });

  const inProgress = enrichedJobs.filter((j) => j.stage === 'In Progress');
  const notBegun = enrichedJobs.filter((j) => j.stage === 'Not begun');
  const completed = enrichedJobs.filter((j) => j.stage === 'Completed');

  const labourCostByJob = {};
  for (const j of enrichedJobs) {
    if (j.costs.labour > 0) labourCostByJob[j.jobKey] = j.costs.labour;
  }
  const totalLabourHours = Math.round(
    enrichedJobs.reduce((s, j) => s + (j.labourHours || 0), 0),
  );

  // Auto-generate flags from receivables data.
  const flags = [];
  for (const r of receivables.filter((x) => (x.daysOverdue ?? 0) > 0).slice(0, 8)) {
    const sev = r.daysOverdue >= 30 ? 'urgent' : 'warn';
    flags.push({
      severity: sev,
      icon: sev === 'urgent' ? '🚨' : '⚠️',
      text: `${r.invoiceNumber || '(no #)'} ${r.client} $${formatMoney(r.amountDue)} — ${r.daysOverdue} days overdue.`,
    });
  }

  // Quote forecast — base = sum of QuoteValue across all SharePoint quote
  // folders (skipping nulls/zero). Folders without a value contribute 0,
  // so the forecast grows as you populate the column.
  const quoteList = quotesResult.quotes || [];
  const forecastBase = quoteList.reduce((s, q) => s + (Number(q.quoteValue) || 0), 0);
  const winRates = manual.forecastWinRates || {};
  const forecast = {
    pipelineTotal: round2(forecastBase),
    conservative: round2(forecastBase * (winRates.conservative ?? 0.25)),
    base:         round2(forecastBase * (winRates.base         ?? 0.40)),
    optimistic:   round2(forecastBase * (winRates.optimistic   ?? 0.60)),
    bestCase:     round2(forecastBase * (winRates.bestCase     ?? 1.00)),
  };

  return {
    meta: {
      generatedAt: new Date().toISOString(),
      dataAsAt: dataAsAt || formatDate(new Date()),
      companyName: manual.company?.name || 'CoreCast Concrete Pty Ltd',
      abn: manual.company?.abn || '',
      sources: {
        monday:          { ok: !mondayResult.error,        message: mondayResult.error || null,        count: jobs.length },
        xeroBills:       { ok: !billsResult.error,         message: billsResult.error || null,         count: billsResult.totalBills || 0 },
        xeroSpend:       { ok: !spendResult.error,         message: spendResult.error || null,         count: spendResult.totalTransactions || 0 },
        xeroReceivables: { ok: !receivablesResult.error,   message: receivablesResult.error || null,   count: receivables.length },
        sharepoint:      { ok: !!quotesResult.available,   message: quotesResult.reason || null,       count: (quotesResult.quotes || []).length },
        spFolders:       { ok: foldersResult.status?.jobs?.ok && foldersResult.status?.quotes?.ok, message: [foldersResult.status?.jobs?.error, foldersResult.status?.quotes?.error].filter(Boolean).join(' / ') || null, count: folders.length },
      },
    },
    summary: {
      activeJobs: inProgress.length + notBegun.length,
      activeJobsBreakdown: `${inProgress.length} in progress · ${notBegun.length} not begun`,
      totalOverdue: receivablesResult.totalOverdue || 0,
      overdueCount: receivablesResult.overdueCount || 0,
      totalOutstanding: receivablesResult.totalOutstanding || 0,
      // "Invoices sent" = sent but not yet due (status='upcoming' OR 'due-soon').
      // These are receivables on the way in — Tom wants visibility on this
      // separately from the Overdue headline.
      totalUpcoming: round2((receivables || [])
        .filter((r) => r.status === 'upcoming' || r.status === 'due-soon')
        .reduce((s, r) => s + r.amountDue, 0)),
      upcomingCount: (receivables || [])
        .filter((r) => r.status === 'upcoming' || r.status === 'due-soon').length,
      totalLabourHours,
      labourCostByJob,
    },
    jobs: { inProgress, notBegun, completed },
    cashflow: {
      totalOutstanding: receivablesResult.totalOutstanding || 0,
      totalOverdue: receivablesResult.totalOverdue || 0,
      receivables: receivables,
      unattributedReceivables,
      costBreakdownByJob: Object.fromEntries(
        Object.entries(billsByJob).map(([job, val]) => [
          job,
          {
            total: val.total,
            labour: val.labour,
            materials: val.materials,
            spend: val.spend || 0,
            lines: val.lines || [],
          },
        ]),
      ),
    },
    flags,
    quotes: {
      files: quotesResult.quotes || [],
      sourceAvailable: !!quotesResult.available,
      sourceMessage: quotesResult.reason || null,
      forecast,
    },
  };
}

function round2(n) { return Math.round(Number(n) * 100) / 100; }
function formatMoney(n) {
  return Number(n).toLocaleString('en-AU', { maximumFractionDigits: 0 });
}
function formatDate(d) {
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' });
}

module.exports = { compose };
