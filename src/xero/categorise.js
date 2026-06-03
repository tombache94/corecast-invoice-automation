/**
 * Map a parsed/persisted invoice to either the LABOUR flow (existing matchWeek)
 * or a BILL flow that creates a draft Accounts Payable invoice in Xero.
 *
 * Account codes come from the Corecast Pty Ltd Xero org.
 */

const ACCOUNT_CODES = {
  COST_OF_GOODS_SOLD: '310',
  WAGES_AND_SALARIES: '477',
  CONSULTING_AND_ACCOUNTING: '412',
  REPAIRS_AND_MAINTENANCE: '473',
  FREIGHT_AND_COURIER: '425',
  GENERAL_EXPENSES: '429',
};

const ACCOUNT_NAMES = {
  '310': 'Cost of Goods Sold',
  '477': 'Wages and Salaries',
  '412': 'Consulting & Accounting',
  '473': 'Repairs and Maintenance',
  '425': 'Freight & Courier',
  '429': 'General Expenses',
};

// Subcontractors who always send labour invoices. Matched case-insensitively
// against invoice.sender. Even when hours weren't extracted from the PDF,
// invoices from these senders should route to Wages & Salaries (477), not COGS.
const KNOWN_LABOUR_SENDERS = [
  'aaron norris',
  'alani fonua',
  'james elliott',
  'jack henderson',
  'demberel ganbayar',
  'nick mourkakos',
  'veronty',
  'kalpana munsi',
];

/**
 * Decide what account code an invoice belongs to.
 * All invoices — including labour — now return a BILL result so the push
 * pipeline can create draft bills in Xero. Labour invoices use account 477
 * (Wages and Salaries); supplier invoices use 310/412/473/425 as before.
 *
 * @param {object} invoice - persisted or parsed invoice
 * @returns {{type:'BILL', accountCode:string, accountName:string, reason:string}}
 */
function categoriseInvoice(invoice) {
  const senderLower = (invoice.sender || '').toLowerCase();

  // Known labour subcontractors — always route to 477 regardless of whether
  // hours were extracted from the PDF.
  if (KNOWN_LABOUR_SENDERS.some((name) => senderLower.includes(name))) {
    return bill('477', 'known-labour-sender');
  }

  const text = [
    invoice.sender,
    invoice.senderEmail,
    invoice.subject,
    invoice.preview,
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  if (/\b(pkf|accounting|accountant|audit|bookkeep)\b/.test(text)) {
    return bill('412', 'consulting / accounting keyword');
  }

  if (/\b(hire|equipment|plant|repair|maintenance|servic(e|ing))\b/.test(text)) {
    return bill('473', 'equipment / repairs keyword');
  }

  if (/\b(freight|delivery|transport|courier|cartage|haulage)\b/.test(text)) {
    return bill('425', 'freight / delivery keyword');
  }

  // Legacy: hoursClaimed > 0 also signals labour for any sender not in the
  // known list above (e.g. new subcontractors not yet added).
  if (invoice.hoursClaimed != null && invoice.hoursClaimed > 0) {
    return bill('477', 'hours-claimed > 0');
  }

  return bill('310', 'default supplier / materials');
}

function bill(code, reason) {
  return {
    type: 'BILL',
    accountCode: code,
    accountName: ACCOUNT_NAMES[code] || code,
    reason,
  };
}

module.exports = { categoriseInvoice, ACCOUNT_CODES, ACCOUNT_NAMES };
