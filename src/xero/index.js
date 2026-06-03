const { categoriseInvoice, ACCOUNT_NAMES } = require('./categorise');

const SUSPICIOUS_AMOUNT_THRESHOLD = 50000;

/**
 * Push a single invoice through the Xero bill pipeline.
 *
 * @param {object} invoice - persisted invoice (sender, amount, date, weekEnding, ...)
 * @param {object} opts
 * @param {object} opts.client - XeroBillsClient instance (or any object exposing the same methods)
 * @param {boolean} [opts.dryRun=false] - run categorisation + duplicate check, but don't create
 * @returns {Promise<object>} - result envelope (see below)
 */
async function pushInvoiceToXero(invoice, { client, dryRun = false } = {}) {
  const category = categoriseInvoice(invoice);

  if (invoice.amount == null) {
    return {
      skipped: true,
      reason: 'amount_missing',
      category,
      flag: 'manual_review',
    };
  }

  if (Number(invoice.amount) > SUSPICIOUS_AMOUNT_THRESHOLD) {
    return {
      skipped: true,
      reason: 'amount_suspicious',
      category,
      flag: 'manual_review',
      detail: `amount $${invoice.amount} exceeds $${SUSPICIOUS_AMOUNT_THRESHOLD} threshold`,
    };
  }

  if (!invoice.sender) {
    return { skipped: true, reason: 'sender_missing', category, flag: 'manual_review' };
  }

  if (!client) throw new Error('pushInvoiceToXero: client is required');

  const contact = await client.findOrCreateContact({
    name: invoice.sender,
    email: invoice.senderEmail || null,
  });

  const billDate = invoice.date || invoice.weekEnding || new Date().toISOString().slice(0, 10);
  const reference = buildReference(invoice);

  const window = weekWindow(invoice.weekEnding || invoice.date || billDate);
  const existing = await client.findExistingBill({
    contactId: contact.contactID,
    dateFrom: window.from,
    dateTo: window.to,
    amount: invoice.amount,
  });

  if (existing) {
    return {
      skipped: true,
      reason: 'duplicate',
      category,
      xeroId: existing.invoiceID,
      xeroNumber: existing.invoiceNumber,
      contact: { contactID: contact.contactID, name: contact.name },
    };
  }

  if (dryRun) {
    return {
      dryRun: true,
      category,
      contact: { contactID: contact.contactID, name: contact.name },
      wouldCreate: {
        date: billDate,
        amount: invoice.amount,
        accountCode: category.accountCode,
        accountName: ACCOUNT_NAMES[category.accountCode] || category.accountCode,
        reference,
      },
    };
  }

  const bill = await client.createDraftBill({
    contact,
    date: billDate,
    description: buildDescription(invoice),
    amount: invoice.amount,
    accountCode: category.accountCode,
    reference,
  });

  return {
    created: true,
    category,
    contact: { contactID: contact.contactID, name: contact.name },
    xeroId: bill.invoiceID,
    xeroNumber: bill.invoiceNumber,
    accountCode: category.accountCode,
    accountName: ACCOUNT_NAMES[category.accountCode] || category.accountCode,
    amount: invoice.amount,
    reference,
  };
}

function buildReference(invoice) {
  if (invoice.invoiceNumber) return String(invoice.invoiceNumber).slice(0, 255);
  if (invoice.subject) return String(invoice.subject).slice(0, 255);
  if (invoice.weekEnding) return `Week ending ${invoice.weekEnding}`;
  return invoice.id || '';
}

function buildDescription(invoice) {
  const ref = invoice.invoiceNumber || invoice.subject || invoice.id || '';
  if (invoice.weekEnding) return `Invoice ${ref} — w/e ${invoice.weekEnding}`.trim();
  return `Invoice ${ref}`.trim();
}

/**
 * Compute a [Mon..Sun] window around the supplied date so duplicate checks
 * catch bills posted on slightly different dates within the same week.
 */
function weekWindow(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  if (Number.isNaN(d.getTime())) {
    return { from: iso, to: iso };
  }
  const dow = d.getUTCDay(); // 0 = Sun, 1 = Mon, ..., 6 = Sat
  const monOffset = dow === 0 ? -6 : 1 - dow;
  const monday = new Date(d);
  monday.setUTCDate(d.getUTCDate() + monOffset);
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  return {
    from: monday.toISOString().slice(0, 10),
    to: sunday.toISOString().slice(0, 10),
  };
}

/**
 * Run the bill pipeline against a list of invoices.
 *
 * @param {Array<object>} invoices
 * @param {object} opts - same as pushInvoiceToXero, plus optional `onProgress(invoice, result)`
 * @returns {Promise<{results: Array, counts: object}>}
 */
async function pushManyToXero(invoices, opts = {}) {
  const results = [];
  for (const invoice of invoices) {
    let result;
    try {
      result = await pushInvoiceToXero(invoice, opts);
    } catch (err) {
      result = { error: err.message, category: categoriseInvoice(invoice) };
    }
    if (typeof opts.onProgress === 'function') {
      try { opts.onProgress(invoice, result); } catch (_) { /* swallow */ }
    }
    results.push({ invoiceId: invoice.id, sender: invoice.sender, amount: invoice.amount, result });
  }

  const counts = { total: results.length, created: 0, duplicates: 0, labour: 0, flagged: 0, errors: 0 };
  for (const r of results) {
    if (r.result.error) counts.errors++;
    else if (r.result.created) counts.created++;
    else if (r.result.reason === 'duplicate') counts.duplicates++;
    else if (r.result.reason === 'labour') counts.labour++;
    else if (r.result.flag === 'manual_review') counts.flagged++;
  }
  return { results, counts };
}

module.exports = {
  pushInvoiceToXero,
  pushManyToXero,
  weekWindow,
  buildReference,
  SUSPICIOUS_AMOUNT_THRESHOLD,
};
