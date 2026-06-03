const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');
const { parseTimesheetExport, lookupHours } = require('../connecteam/timesheetParser');
const { categoriseInvoice } = require('../xero/categorise');
const { pushManyToXero, pushInvoiceToXero } = require('../xero');

const TIMESHEET_PATTERN = /^timeclock-timesheet_overview_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.xlsx$/i;
const SENDER_SKIP_FILE = path.resolve(__dirname, '../../config/sender-skip.json');

/**
 * Load the list of sender names (substrings) to exclude from the matcher.
 * Missing/unreadable file = empty list (skip nothing).
 */
function loadSenderSkipList() {
  try {
    if (!fsSync.existsSync(SENDER_SKIP_FILE)) return [];
    const parsed = JSON.parse(fsSync.readFileSync(SENDER_SKIP_FILE, 'utf8'));
    return Array.isArray(parsed.senders)
      ? parsed.senders.map((s) => String(s).toLowerCase().trim()).filter(Boolean)
      : [];
  } catch (_) {
    return [];
  }
}

function shouldSkipInvoice(invoice, skipList) {
  if (!skipList.length) return false;
  const sender = String(invoice.sender || '').toLowerCase();
  return skipList.some((needle) => sender.includes(needle));
}

/**
 * Find the most recent Connecteam timesheet export in the watch folder.
 * Picks the file whose week-end date (encoded in the filename) is the latest.
 *
 * @param {string} folder - Path to the watch folder
 * @returns {Promise<string|null>} - Absolute path to the latest XLSX, or null
 */
async function findLatestTimesheet(folder) {
  let entries;
  try {
    entries = await fs.readdir(folder);
  } catch (err) {
    if (err.code === 'ENOENT') return null;
    throw err;
  }

  const matches = entries
    .map((name) => {
      const m = name.match(TIMESHEET_PATTERN);
      return m ? { name, weekStart: m[1], weekEnd: m[2] } : null;
    })
    .filter(Boolean);

  if (!matches.length) return null;

  matches.sort((a, b) => (a.weekEnd < b.weekEnd ? 1 : a.weekEnd > b.weekEnd ? -1 : 0));
  return path.resolve(folder, matches[0].name);
}

/**
 * Classify a single invoice against the timesheet hours map.
 * Returns one of: match | mismatch | not_found | no_hours_in_timesheet | invoice_hours_missing
 */
function classifyInvoice(invoice, byName, tolerance) {
  if (invoice.hoursClaimed == null) {
    return { status: 'invoice_hours_missing', timesheetHours: null, diff: null };
  }

  const timesheetHours = lookupHours(invoice.sender || '', byName);

  if (timesheetHours == null) {
    return { status: 'not_found', timesheetHours: null, diff: null };
  }

  if (timesheetHours === 0) {
    return { status: 'no_hours_in_timesheet', timesheetHours: 0, diff: null };
  }

  const diff = Math.round((invoice.hoursClaimed - timesheetHours) * 100) / 100;
  if (Math.abs(diff) <= tolerance) {
    return { status: 'match', timesheetHours, diff };
  }
  return { status: 'mismatch', timesheetHours, diff };
}

/**
 * Map a classification result to the persisted invoice store status.
 */
function statusForStore(classification) {
  switch (classification) {
    case 'match':
      return 'validated';
    case 'mismatch':
      return 'mismatch';
    case 'not_found':
    case 'no_hours_in_timesheet':
    case 'invoice_hours_missing':
      return 'review_needed';
    default:
      return 'review_needed';
  }
}

/**
 * Run the weekly match.
 *
 * @param {object} opts
 * @param {string} [opts.timesheetPath] - Explicit XLSX path (overrides watch folder)
 * @param {string} [opts.watchFolder]   - Folder to scan when timesheetPath is omitted
 * @param {number} [opts.tolerance]     - Hours tolerance for mismatch (default 0.5)
 * @param {object} opts.store           - InvoiceStore instance (required)
 * @param {boolean} [opts.updateStatuses=true] - Persist status changes back to the store
 * @param {object} [opts.xeroClient]    - Optional XeroBillsClient. If provided, BILL invoices
 *                                        in the same week are pushed to Xero as draft bills.
 * @param {boolean} [opts.xeroDryRun=false] - Categorise + duplicate-check only, no creation
 * @returns {Promise<object>} - { weekStart, weekEnd, timesheetPath, results[], counts, bills }
 */
async function matchWeek(opts) {
  const {
    timesheetPath: explicitPath,
    watchFolder,
    tolerance = 0.5,
    store,
    updateStatuses = true,
    xeroClient = null,
    xeroDryRun = false,
  } = opts;

  if (!store) throw new Error('matchWeek: store is required');

  const timesheetPath = explicitPath
    ? path.resolve(explicitPath)
    : await findLatestTimesheet(watchFolder);

  if (!timesheetPath) {
    throw new Error(
      `No timesheet found. Looked in: ${watchFolder}. ` +
        `Expected filename pattern: timeclock-timesheet_overview_YYYY-MM-DD_YYYY-MM-DD.xlsx`,
    );
  }

  const timesheet = parseTimesheetExport(timesheetPath);
  const { weekStart, weekEnd, byName, employees } = timesheet;

  if (!weekStart || !weekEnd) {
    throw new Error(
      `Could not extract week range from filename: ${path.basename(timesheetPath)}`,
    );
  }

  const rawInvoices = await store.getAllInRange(weekStart, weekEnd);

  // Filter out senders on the skip list (config/sender-skip.json) — these are
  // emails that look invoice-shaped but aren't real subcontractor invoices.
  const skipList = loadSenderSkipList();
  const skippedInvoices = [];
  const invoices = [];
  for (const inv of rawInvoices) {
    if (shouldSkipInvoice(inv, skipList)) skippedInvoices.push(inv);
    else invoices.push(inv);
  }

  // Split: hoursClaimed > 0 → labour matching; everything else → bill pipeline.
  const labourInvoices = [];
  const billInvoices = [];
  for (const invoice of invoices) {
    const cat = categoriseInvoice(invoice);
    if (cat.type === 'LABOUR') labourInvoices.push(invoice);
    else billInvoices.push(invoice);
  }

  const results = [];
  for (const invoice of labourInvoices) {
    const classification = classifyInvoice(invoice, byName, tolerance);
    const newStoreStatus = statusForStore(classification.status);

    let storeUpdate = null;
    if (updateStatuses && invoice.status !== newStoreStatus) {
      const note =
        classification.status === 'match'
          ? `Validated against Connecteam (${classification.timesheetHours}h)`
          : classification.status === 'mismatch'
            ? `Mismatch: invoice ${invoice.hoursClaimed}h vs Connecteam ${classification.timesheetHours}h (diff ${classification.diff}h)`
            : classification.status === 'not_found'
              ? `Sender "${invoice.sender}" not found in Connecteam timesheet`
              : classification.status === 'no_hours_in_timesheet'
                ? `Found in timesheet but 0 hours clocked`
                : `Invoice hours not extracted`;
      try {
        storeUpdate = await store.updateStatus(invoice.id, newStoreStatus, { note });
      } catch (err) {
        storeUpdate = { error: err.message };
      }
    }

    results.push({
      invoiceId: invoice.id,
      sender: invoice.sender,
      amount: invoice.amount,
      hoursClaimed: invoice.hoursClaimed,
      weekEnding: invoice.weekEnding,
      timesheetHours: classification.timesheetHours,
      diff: classification.diff,
      classification: classification.status,
      previousStatus: invoice.status,
      newStatus: newStoreStatus,
      updated: storeUpdate && !storeUpdate.error ? true : false,
      updateError: storeUpdate?.error || null,
    });
  }

  const counts = results.reduce(
    (acc, r) => {
      acc[r.classification] = (acc[r.classification] || 0) + 1;
      return acc;
    },
    { match: 0, mismatch: 0, not_found: 0, no_hours_in_timesheet: 0, invoice_hours_missing: 0 },
  );
  counts.total = results.length;

  let bills = null;
  if (billInvoices.length) {
    bills = xeroClient
      ? await pushManyToXero(billInvoices, { client: xeroClient, dryRun: xeroDryRun })
      : await collectBillsWithoutXero(billInvoices);

    if (updateStatuses && xeroClient) {
      for (const entry of bills.results) {
        const r = entry.result;
        if (r.error) {
          await safeUpdate(store, entry.invoiceId, 'review_needed', `Xero error: ${r.error}`);
        } else if (r.created) {
          await safeUpdate(store, entry.invoiceId, 'billed', `Draft bill created in Xero: ${r.xeroNumber || r.xeroId} (${r.accountName})`, {
            xeroBillId: r.xeroId,
            xeroBillNumber: r.xeroNumber || null,
            xeroAccountCode: r.accountCode,
          });
        } else if (r.reason === 'duplicate') {
          await safeUpdate(store, entry.invoiceId, 'bill_duplicate', `Bill already exists in Xero: ${r.xeroNumber || r.xeroId}`, {
            xeroBillId: r.xeroId,
            xeroBillNumber: r.xeroNumber || null,
          });
        } else if (r.flag === 'manual_review') {
          const detail = r.detail || r.reason;
          await safeUpdate(store, entry.invoiceId, 'review_needed', `Bill flagged: ${detail}`);
        }
      }
    }
  }

  return {
    weekStart,
    weekEnd,
    timesheetPath,
    employees,
    byName,
    tolerance,
    results,
    counts,
    bills,
    skippedInvoices,
  };
}

async function safeUpdate(store, id, status, note, fields) {
  try {
    await store.updateStatus(id, status, { note, fields });
  } catch (err) {
    // swallow — the in-memory `bills.results` already records the outcome
  }
}

/**
 * When no Xero client is configured, we still surface BILL invoices in the report
 * so Tom knows what would have been created. Uses a stub client so suspicious-amount
 * and missing-field rules still surface in the report.
 */
async function collectBillsWithoutXero(invoices) {
  const stubClient = {
    async findOrCreateContact() {
      throw new Error('no_xero_client');
    },
    async findExistingBill() { return null; },
    async createDraftBill() { throw new Error('no_xero_client'); },
  };
  const results = [];
  for (const invoice of invoices) {
    let result;
    try {
      result = await pushInvoiceToXero(invoice, { client: stubClient, dryRun: true });
    } catch (err) {
      if (err.message !== 'no_xero_client') throw err;
      // Stub threw at contact lookup — invoice passed all static rules.
      result = { skipped: true, reason: 'no_xero_client', category: categoriseInvoice(invoice) };
    }
    results.push({
      invoiceId: invoice.id,
      sender: invoice.sender,
      amount: invoice.amount,
      result,
    });
  }
  const counts = { total: results.length, created: 0, duplicates: 0, labour: 0, flagged: 0, errors: 0 };
  for (const r of results) {
    if (r.result.flag === 'manual_review') counts.flagged++;
  }
  return { results, counts };
}

module.exports = {
  matchWeek,
  findLatestTimesheet,
  classifyInvoice,
  statusForStore,
  TIMESHEET_PATTERN,
};
