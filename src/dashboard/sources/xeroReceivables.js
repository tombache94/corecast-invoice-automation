/**
 * Fetch all unpaid ACCREC (sales) invoices from Xero. Computes days-overdue
 * relative to today's date and classifies each invoice as overdue / due-soon /
 * upcoming. Used for the Cashflow tab and for invoice tags on each job card.
 */

const { xeroRequest } = require('../../xero/api');

async function fetchReceivables() {
  // Xero won't return PAID invoices in this filter — we want what's outstanding.
  const all = [];
  for (const status of ['AUTHORISED', 'SUBMITTED', 'DRAFT']) {
    const res = await xeroRequest({
      path: `/api.xro/2.0/Invoices?where=Type%3D%3D%22ACCREC%22%26%26Status%3D%3D%22${status}%22&page=1`,
    });
    if (res.status === 200 && Array.isArray(res.body.Invoices)) {
      all.push(...res.body.Invoices);
    }
  }

  const today = new Date();
  today.setUTCHours(0, 0, 0, 0);

  const enriched = all
    .filter((inv) => Number(inv.AmountDue || 0) > 0) // skip zeroes / fully paid
    .map((inv) => {
      const due = inv.DueDate ? parseXeroDate(inv.DueDate) : null;
      const daysOverdue = due ? Math.floor((today - due) / 86_400_000) : null;
      const status = classifyDueness(daysOverdue);
      return {
        invoiceId: inv.InvoiceID,
        invoiceNumber: inv.InvoiceNumber || '',
        reference: inv.Reference || '',
        client: inv.Contact?.Name || '?',
        description: (inv.LineItems?.[0]?.Description || '').slice(0, 120),
        date: parseXeroDate(inv.Date)?.toISOString().slice(0, 10) || null,
        dueDate: due ? due.toISOString().slice(0, 10) : null,
        amountDue: Number(inv.AmountDue),
        amountTotal: Number(inv.Total),
        status,            // overdue | due-soon | upcoming
        daysOverdue,
        xeroStatus: inv.Status, // DRAFT / AUTHORISED / SUBMITTED
      };
    });

  // Sort: most-overdue first, then by due date
  enriched.sort((a, b) => {
    const ao = a.daysOverdue == null ? -Infinity : a.daysOverdue;
    const bo = b.daysOverdue == null ? -Infinity : b.daysOverdue;
    if (ao !== bo) return bo - ao;
    return (a.dueDate || '').localeCompare(b.dueDate || '');
  });

  const totalOutstanding = round2(enriched.reduce((s, x) => s + x.amountDue, 0));
  const totalOverdue = round2(
    enriched.filter((x) => x.daysOverdue != null && x.daysOverdue > 0)
            .reduce((s, x) => s + x.amountDue, 0),
  );
  const overdueCount = enriched.filter((x) => x.daysOverdue != null && x.daysOverdue > 0).length;

  return {
    invoices: enriched,
    totalOutstanding,
    totalOverdue,
    overdueCount,
    fetchedAt: new Date().toISOString(),
  };
}

function parseXeroDate(s) {
  if (!s) return null;
  // Xero dates come as either ISO ("2026-05-01") or "/Date(1234567890)+0000/"
  const m = String(s).match(/\/Date\((\d+)/);
  if (m) return new Date(Number(m[1]));
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

function classifyDueness(daysOverdue) {
  if (daysOverdue == null) return 'no-due-date';
  if (daysOverdue > 0) return 'overdue';
  if (daysOverdue >= -7) return 'due-soon';
  return 'upcoming';
}

function round2(n) {
  return Math.round(Number(n) * 100) / 100;
}

module.exports = { fetchReceivables };
