/**
 * Post-refresh email notification.
 *
 * Compares the just-composed dashboard data against the previous run's
 * snapshot (data/dashboard-last-state.json) and emails Tom a short
 * summary of what changed + the current state.
 *
 * Designed to be best-effort: never throws (caller logs the result),
 * and tolerates a missing previous-state file on first run.
 */

const fs = require('fs');
const path = require('path');
const { config } = require('../config');
const { OutlookClient } = require('../email/outlook');

const STATE_FILE = path.resolve(__dirname, '../../data/dashboard-last-state.json');
// Comma-separated list — one email gets sent with all addresses on the To: line.
const DEFAULT_RECIPIENTS = (
  process.env.DASHBOARD_NOTIFY_RECIPIENT
  || 'tom@corecastconcrete.com.au,josh@corecastconcrete.com.au'
)
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);
const SENDER = process.env.DASHBOARD_NOTIFY_SENDER || config.alerts?.reportSender || 'accounts@corecastconcrete.com.au';
const PUBLIC_URL = process.env.DASHBOARD_PUBLIC_URL || 'https://corecast-dashboard.pages.dev';

function loadPrevState() {
  if (!fs.existsSync(STATE_FILE)) return null;
  try {
    return JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
  } catch {
    return null;
  }
}

function saveState(state) {
  fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
  fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2));
}

function snapshotFromData(data) {
  const receivableIds = (data.cashflow?.receivables || []).map((r) => r.invoiceId);
  const folderNames = Object.values(data.meta?.sources || {})
    .map((s) => s.count) // not what we want — leave folders for now
    .filter(() => false);
  return {
    refreshedAt: data.meta?.generatedAt || new Date().toISOString(),
    summary: {
      activeJobs: data.summary?.activeJobs ?? null,
      totalOverdue: data.summary?.totalOverdue ?? 0,
      overdueCount: data.summary?.overdueCount ?? 0,
      totalOutstanding: data.summary?.totalOutstanding ?? 0,
      totalLabourHours: data.summary?.totalLabourHours ?? 0,
    },
    receivableIds,
    overdueInvoiceIds: (data.cashflow?.receivables || [])
      .filter((r) => (r.daysOverdue ?? 0) > 0)
      .map((r) => r.invoiceId),
    sourceCounts: Object.fromEntries(
      Object.entries(data.meta?.sources || {}).map(([k, v]) => [k, v.count || 0]),
    ),
  };
}

function formatMoney(n) {
  return Number(n).toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDelta(curr, prev, opts = {}) {
  if (prev == null) return null;
  const delta = curr - prev;
  if (Math.abs(delta) < (opts.epsilon ?? 0.01)) return '= no change';
  const arrow = delta > 0 ? '▲' : '▼';
  const sign = delta > 0 ? '+' : '−';
  const abs = Math.abs(delta);
  if (opts.money) return `${arrow} ${sign}$${formatMoney(abs)}`;
  return `${arrow} ${sign}${Math.round(abs)}`;
}

function buildEmailBody(currSnap, prevSnap, data, filingDigest, filingStats) {
  const lines = [];
  const refreshedLocal = new Date(currSnap.refreshedAt).toLocaleString('en-AU', {
    timeZone: 'Australia/Perth',
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', hour12: false,
  });

  lines.push(`Dashboard refreshed at ${refreshedLocal} Perth time:`);
  lines.push(PUBLIC_URL);
  lines.push('');

  // === CHANGES SECTION ===
  if (prevSnap) {
    const prevLocal = new Date(prevSnap.refreshedAt).toLocaleString('en-AU', {
      timeZone: 'Australia/Perth',
      day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: false,
    });
    lines.push(`CHANGES SINCE LAST REFRESH (${prevLocal}):`);

    const prevIds = new Set(prevSnap.receivableIds || []);
    const newReceivables = (data.cashflow?.receivables || []).filter((r) => !prevIds.has(r.invoiceId));
    const removedIds = (prevSnap.receivableIds || []).filter((id) => !currSnap.receivableIds.includes(id));

    if (newReceivables.length === 0 && removedIds.length === 0) {
      lines.push('• No new receivables.');
    } else {
      if (newReceivables.length) {
        lines.push(`• ${newReceivables.length} new receivable(s):`);
        for (const r of newReceivables.slice(0, 5)) {
          const od = r.daysOverdue > 0 ? `${r.daysOverdue}d overdue` : (r.daysOverdue <= 0 && r.daysOverdue != null ? `due in ${-r.daysOverdue}d` : 'no due date');
          lines.push(`    – ${r.invoiceNumber || '(no #)'} ${r.client} $${formatMoney(r.amountDue)} (${od})`);
        }
        if (newReceivables.length > 5) lines.push(`    – ...and ${newReceivables.length - 5} more`);
      }
      if (removedIds.length) lines.push(`• ${removedIds.length} receivable(s) paid/voided since last run`);
    }

    const prevOverdue = new Set(prevSnap.overdueInvoiceIds || []);
    const newlyOverdue = (data.cashflow?.receivables || [])
      .filter((r) => (r.daysOverdue ?? 0) > 0 && !prevOverdue.has(r.invoiceId));
    if (newlyOverdue.length) {
      lines.push(`• ${newlyOverdue.length} invoice(s) newly crossed into overdue:`);
      for (const r of newlyOverdue.slice(0, 5)) {
        lines.push(`    – ${r.invoiceNumber || '(no #)'} ${r.client} $${formatMoney(r.amountDue)} (${r.daysOverdue}d overdue)`);
      }
    }

    const overdueDelta = formatDelta(currSnap.summary.totalOverdue, prevSnap.summary.totalOverdue, { money: true });
    const outstandingDelta = formatDelta(currSnap.summary.totalOutstanding, prevSnap.summary.totalOutstanding, { money: true });
    const jobsDelta = formatDelta(currSnap.summary.activeJobs, prevSnap.summary.activeJobs);
    if (overdueDelta !== '= no change') lines.push(`• Total overdue: ${overdueDelta}`);
    if (outstandingDelta !== '= no change') lines.push(`• Total outstanding: ${outstandingDelta}`);
    if (jobsDelta !== '= no change') lines.push(`• Active jobs: ${jobsDelta}`);

    const prevFolderCount = prevSnap.sourceCounts?.spFolders ?? null;
    const currFolderCount = currSnap.sourceCounts?.spFolders ?? null;
    if (prevFolderCount != null && currFolderCount != null && prevFolderCount !== currFolderCount) {
      const arrow = currFolderCount > prevFolderCount ? '▲' : '▼';
      lines.push(`• SharePoint job/quote folders: ${arrow} ${prevFolderCount} → ${currFolderCount}`);
    }
    lines.push('');
  } else {
    lines.push('CHANGES: (first refresh — no previous state to compare)');
    lines.push('');
  }

  // === CURRENT STATE ===
  lines.push('CURRENT STATE:');
  lines.push(`• ${currSnap.summary.activeJobs ?? '?'} active jobs`);
  lines.push(`• $${formatMoney(currSnap.summary.totalOverdue)} overdue across ${currSnap.summary.overdueCount} invoice(s)`);
  lines.push(`• $${formatMoney(currSnap.summary.totalOutstanding)} total outstanding`);

  // Labour hours — show per-job breakdown so it reconciles with the
  // dashboard's per-card numbers, plus the total.
  const jobsForHours = [
    ...(data.jobs?.inProgress || []),
    ...(data.jobs?.notBegun || []),
  ].filter((j) => (j.labourHours || 0) > 0);
  if (jobsForHours.length) {
    const perJob = jobsForHours
      .sort((a, b) => (b.labourHours || 0) - (a.labourHours || 0))
      .map((j) => `${j.jobKey || j.name}: ${Math.round(j.labourHours)} hrs`)
      .join(' · ');
    lines.push(`• ~${currSnap.summary.totalLabourHours} labour hours logged (${perJob})`);
  } else {
    lines.push(`• ~${currSnap.summary.totalLabourHours} labour hours logged`);
  }
  lines.push('');

  // === FILED ATTACHMENTS ===
  if (filingDigest && filingDigest.length > 0) {
    lines.push('FILED ATTACHMENTS THIS RUN:');
    const byTier = { keyword: [], domain: [], to_sort: [] };
    for (const e of filingDigest) {
      const tier = byTier[e.tier] ? e.tier : 'to_sort';
      byTier[tier].push(e);
    }
    if (byTier.keyword.length) {
      lines.push(`• ${byTier.keyword.length} file(s) → matched to a job folder:`);
      for (const e of byTier.keyword.slice(0, 10)) {
        lines.push(`    – ${e.attachmentName}`);
        lines.push(`        from ${e.from || 'unknown'}  —  "${e.subject || '(no subject)'}"`);
        lines.push(`        → ${e.label}`);
        if (e.webUrl) lines.push(`        ${e.webUrl}`);
      }
      if (byTier.keyword.length > 10) lines.push(`    – ...and ${byTier.keyword.length - 10} more`);
    }
    if (byTier.domain.length) {
      lines.push(`• ${byTier.domain.length} file(s) → routed to Accounts Payable by supplier domain:`);
      for (const e of byTier.domain.slice(0, 10)) {
        lines.push(`    – ${e.attachmentName}  (from ${e.from})  → ${e.label}`);
      }
      if (byTier.domain.length > 10) lines.push(`    – ...and ${byTier.domain.length - 10} more`);
    }
    if (byTier.to_sort.length) {
      lines.push(`• ⚠️  ${byTier.to_sort.length} file(s) → sent to "To Sort" (needs manual review):`);
      for (const e of byTier.to_sort.slice(0, 10)) {
        lines.push(`    – ${e.attachmentName}  (from ${e.from})  —  "${e.subject || '(no subject)'}"`);
      }
      if (byTier.to_sort.length > 10) lines.push(`    – ...and ${byTier.to_sort.length - 10} more`);
    }
    lines.push('');
  } else if (filingStats) {
    lines.push(`FILED ATTACHMENTS THIS RUN: 0 new files (scanned ${filingStats.messagesScanned} message(s), ${filingStats.alreadyFiled} already filed)`);
    lines.push('');
  }

  // === SOURCE HEALTH ===
  const sources = data.meta?.sources || {};
  const healthBits = Object.entries(sources).map(([name, s]) => `${s.ok ? '✅' : '⚠️ '}${name}`);
  lines.push('SOURCE HEALTH: ' + healthBits.join(' · '));
  const unhealthy = Object.entries(sources).filter(([, s]) => !s.ok);
  for (const [name, s] of unhealthy) {
    lines.push(`  ⚠️  ${name}: ${s.message || '(no message)'}`);
  }

  lines.push('');
  lines.push('— CoreCast Dashboard Refresh');

  return lines.join('\n');
}

function buildSubject(currSnap, prevSnap, filingDigest) {
  const local = new Date(currSnap.refreshedAt).toLocaleString('en-AU', {
    timeZone: 'Australia/Perth',
    day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: false,
  });
  const filedCount = Array.isArray(filingDigest) ? filingDigest.length : 0;
  const filedBit = filedCount > 0 ? ` · ${filedCount} file${filedCount > 1 ? 's' : ''} filed` : '';
  if (!prevSnap) return `[CoreCast] Dashboard refreshed — ${local}${filedBit}`;
  const newRec = currSnap.receivableIds.filter((id) => !(prevSnap.receivableIds || []).includes(id)).length;
  const overdueDelta = currSnap.summary.totalOverdue - prevSnap.summary.totalOverdue;
  const flagBit = newRec > 0 ? ` · ${newRec} new invoice${newRec > 1 ? 's' : ''}` : '';
  const overdueBit = Math.abs(overdueDelta) >= 1 ? ` · overdue ${overdueDelta > 0 ? '▲' : '▼'}$${formatMoney(Math.abs(overdueDelta))}` : '';
  return `[CoreCast] Dashboard refreshed — ${local}${flagBit}${overdueBit}${filedBit}`;
}

async function notifyOnRefresh(data, { recipient, filingDigest, filingStats } = {}) {
  try {
    const prevSnap = loadPrevState();
    const currSnap = snapshotFromData(data);

    const tenantId = process.env.MS_TENANT_ID;
    const clientId = process.env.MS_CLIENT_ID;
    const clientSecret = process.env.MS_CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) {
      return { ok: false, message: 'Missing MS_* creds — cannot send notification' };
    }

    const outlook = new OutlookClient({ tenantId, clientId, clientSecret });
    const subject = buildSubject(currSnap, prevSnap, filingDigest);
    const body = buildEmailBody(currSnap, prevSnap, data, filingDigest, filingStats);
    const to = recipient
      ? (Array.isArray(recipient) ? recipient : [recipient])
      : DEFAULT_RECIPIENTS;

    await outlook.sendMail(SENDER, { to, subject, text: body });

    // Only save state on a successful send — if email fails, next refresh
    // should still see the previous state and try to send the (cumulative) diff.
    saveState(currSnap);
    return { ok: true, message: `Sent to ${to.join(', ')}`, subject };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

module.exports = { notifyOnRefresh, buildEmailBody, buildSubject, snapshotFromData };
