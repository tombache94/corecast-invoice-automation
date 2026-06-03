/**
 * Fetch all ACCPAY (purchase) bills from Xero, grouped by Job tracking option.
 * Returns per-job cost totals split by category code so the dashboard can show
 * labour vs materials per job.
 */

const { xeroRequest } = require('../../xero/api');

const ACCOUNT_LABEL = {
  '310': 'Cost of Goods Sold',
  '477': 'Wages and Salaries',
  '412': 'Consulting & Accounting',
  '473': 'Repairs and Maintenance',
  '425': 'Freight & Courier',
  '429': 'General Expenses',
};

// Description carries the strongest labour signal in this org's data — most
// bills are coded to 310 (Cost of Goods Sold) regardless of whether they're
// for labour or materials, so we lean on the line description instead.
function isLabourLine(description, accountCode) {
  const d = String(description || '').toLowerCase();
  // Strong materials/equipment signals — short-circuit first.
  if (/\bmaterial(s)?\b|\bfreight\b|\bdelivery\b|\bcrane\b|\btransport\b|\btilt tray\b|\bformwork\b|\bbracket\b|\bsilicone\b|\bbolt\b|\bferrule\b|\btape\b|\bfillet\b/.test(d)) return false;
  if (accountCode === '425') return false; // Freight & Courier
  // Labour signals.
  if (/\blabour\b|\blabor\b|\bwages?\b|\bsub-?contract|\bsteel fixing\b|\bworks at\b|\b\d+\s*hrs?\b|\b\d+\s*hours?\b|\bx\s*\d+\s*days?\b|\bconcrete\b/.test(d)) return true;
  // Account-code fallback (rarely populated as labour-specific in this org's data).
  return accountCode === '477' || accountCode === '412';
}

function parseXeroDate(s) {
  if (!s) return null;
  const m = String(s).match(/\/Date\((\d+)/);
  if (m) return new Date(Number(m[1]));
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

async function fetchBillsByJob() {
  const allBills = [];
  for (const status of ['DRAFT', 'AUTHORISED', 'PAID']) {
    const res = await xeroRequest({
      path: `/api.xro/2.0/Invoices?where=Type%3D%3D%22ACCPAY%22%26%26Status%3D%3D%22${status}%22&page=1`,
    });
    if (res.status === 200 && Array.isArray(res.body.Invoices)) {
      allBills.push(...res.body.Invoices);
    }
  }

  // Group by Job tracking option
  const byJob = {};
  let untagged = [];

  for (const bill of allBills) {
    const lineItems = bill.LineItems || [];
    const billDate = parseXeroDate(bill.Date);
    for (const li of lineItems) {
      const trackings = li.Tracking || [];
      const job = trackings.find((t) => /^job/i.test(t.Name))?.Option || null;
      const amount = Number(li.LineAmount || 0);
      const accountCode = li.AccountCode || '';
      const labour = isLabourLine(li.Description, accountCode);

      const entry = {
        billId: bill.InvoiceID,
        billNumber: bill.InvoiceNumber || bill.Reference || '',
        billReference: bill.Reference || '',
        contact: bill.Contact?.Name || '?',
        date: billDate ? billDate.toISOString().slice(0, 10) : null,
        status: bill.Status,
        accountCode,
        accountName: ACCOUNT_LABEL[accountCode] || accountCode,
        description: li.Description || '',
        amount,
        labour,
      };

      if (!job) {
        untagged.push(entry);
        continue;
      }

      if (!byJob[job]) {
        byJob[job] = { total: 0, labour: 0, materials: 0, lines: [] };
      }
      byJob[job].total += amount;
      if (labour) byJob[job].labour += amount;
      else byJob[job].materials += amount;
      byJob[job].lines.push(entry);
    }
  }

  // Round totals
  for (const job of Object.keys(byJob)) {
    byJob[job].total = round2(byJob[job].total);
    byJob[job].labour = round2(byJob[job].labour);
    byJob[job].materials = round2(byJob[job].materials);
  }

  return { byJob, untagged, totalBills: allBills.length };
}

function round2(n) {
  return Math.round(Number(n) * 100) / 100;
}

module.exports = { fetchBillsByJob };
