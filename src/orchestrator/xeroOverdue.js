/**
 * Direct Xero overdue-receivable poll for the daily orchestrator.
 *
 * Reuses the dashboard's fetchReceivables() to pull all unpaid ACCREC
 * invoices, filters to status='overdue', and matches each invoice's
 * Reference field (substring, case-insensitive) against the active jobs
 * list — both the full Monday job name AND the canonical jobKey/alias
 * from manualAllocations.jobAliases.
 *
 * Per Tom's rule: NEVER guess from client name. Macrete bills both
 * Eneabba and Regans Ford; Pennant appears on multiple jobs. The
 * Reference field is the only safe source of job assignment.
 *
 * Caller is responsible for deduplication against previously-noted
 * invoice IDs (state.notedOverdueInvoiceIds).
 */

const { fetchReceivables } = require('../dashboard/sources/xeroReceivables');

const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

async function pollOverdueInvoices(activeJobs) {
  const { invoices } = await fetchReceivables();
  const overdue = invoices.filter((inv) => inv.status === 'overdue' && inv.daysOverdue > 0);

  return overdue.map((inv) => classifyInvoice(inv, activeJobs));
}

function classifyInvoice(invoice, activeJobs) {
  const refRaw = (invoice.reference || '').trim();
  if (!refRaw) {
    return { invoice, match: null, reason: 'reference-empty' };
  }
  const refLower = refRaw.toLowerCase();

  // Build (full name + jobKey alias) candidates per active job. Match if
  // either appears as a substring of the Reference. We do NOT match in the
  // reverse direction — the rule is "Reference contains job name", not
  // "job name contains Reference" (which would over-match: a Reference of
  // "Misc" would match a job named "Misc Repairs").
  for (const job of activeJobs) {
    const candidates = uniqLower([job.name, job.jobKey].filter(Boolean));
    for (const cand of candidates) {
      if (refLower.includes(cand)) {
        return { invoice, match: job, reason: 'matched', matchedOn: cand };
      }
    }
  }

  return { invoice, match: null, reason: 'no-active-job-matched' };
}

function uniqLower(arr) {
  const seen = new Set();
  const out = [];
  for (const s of arr) {
    const k = s.toLowerCase().trim();
    if (k && !seen.has(k)) {
      seen.add(k);
      out.push(k);
    }
  }
  return out;
}

function formatOverdueNote(invoice, today = new Date()) {
  const day = String(today.getDate()).padStart(2, '0');
  const mon = MONTHS[today.getMonth()];
  const year = today.getFullYear();
  const date = `${day}-${mon}-${year}`;
  const amount = Number(invoice.amountDue).toLocaleString('en-AU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  const invRef = invoice.invoiceNumber || invoice.invoiceId.slice(0, 8);
  const due = invoice.dueDate || 'unknown';
  return `[${date}] XERO: ${invRef} to ${invoice.client} ($${amount}) is overdue (due ${due}, now ${invoice.daysOverdue} days overdue). Ref: "${invoice.reference}".`;
}

module.exports = { pollOverdueInvoices, formatOverdueNote };
