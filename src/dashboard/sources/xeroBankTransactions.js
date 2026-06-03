/**
 * Fetch SPEND bank transactions from Xero. Returns the raw normalised list —
 * compose.js handles attribution to a specific job by matching the Reference
 * against currently-active Monday jobs.
 *
 * Includes both reconciled and unreconciled (the default API behaviour).
 * Limited to the last 18 months at the API level so we don't drag the full
 * org history every refresh.
 */

const { xeroRequest } = require('../../xero/api');

const PAGE_SIZE_HINT = 100; // Xero default
const MAX_PAGES = 5;        // 500 transactions cap — generous for current volume

function parseXeroDate(s) {
  if (!s) return null;
  const m = String(s).match(/\/Date\((\d+)/);
  if (m) return new Date(Number(m[1]));
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

async function fetchSpendTransactions() {
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 18);
  const dateStr = `DateTime(${cutoff.getFullYear()},${cutoff.getMonth() + 1},${cutoff.getDate()})`;
  const where = encodeURIComponent(`Type=="SPEND"&&Date>=${dateStr}`);

  const all = [];
  for (let page = 1; page <= MAX_PAGES; page++) {
    const res = await xeroRequest({
      path: `/api.xro/2.0/BankTransactions?where=${where}&page=${page}&order=Date%20DESC`,
    });
    if (res.status !== 200 || !Array.isArray(res.body.BankTransactions)) break;
    const txns = res.body.BankTransactions;
    all.push(...txns);
    if (txns.length < PAGE_SIZE_HINT) break; // last page
  }

  const transactions = [];
  for (const t of all) {
    const amount = Number(t.SubTotal || 0); // ex-GST, consistent with bills
    if (!amount) continue;
    const date = parseXeroDate(t.Date);
    transactions.push({
      transactionId: t.BankTransactionID,
      reference: t.Reference || '',
      contact: t.Contact?.Name || '?',
      date: date ? date.toISOString().slice(0, 10) : null,
      isReconciled: !!t.IsReconciled,
      bankAccount: t.BankAccount?.Name || '',
      description: (t.LineItems?.[0]?.Description || '').slice(0, 120),
      amount,
      type: 'spend', // distinguishes from ACCPAY bill lines downstream
    });
  }

  return { transactions, totalTransactions: all.length };
}

module.exports = { fetchSpendTransactions };
