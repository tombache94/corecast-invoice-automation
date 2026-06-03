const fs = require('fs/promises');
const path = require('path');
const { XeroClient } = require('xero-node');

const DEFAULT_SCOPES = [
  'offline_access',
  'accounting.transactions',
  'accounting.contacts',
  'accounting.attachments',
];

/**
 * Wrapper around the xero-node SDK that:
 *   - reads credentials from config / env
 *   - keeps the refresh token cached on disk (Xero rotates them on every refresh)
 *   - exposes the high-level operations the bill pipeline needs
 *
 * This wrapper is intentionally thin — see src/xero/index.js for the orchestration
 * (categorise → find/create contact → check duplicate → create draft bill).
 */
class XeroBillsClient {
  constructor({
    clientId,
    clientSecret,
    tenantId,
    refreshToken,
    accessToken,
    tokenCacheFile,
    scopes = DEFAULT_SCOPES,
  }) {
    if (!clientId || !clientSecret) {
      throw new Error(
        'XeroBillsClient: XERO_CLIENT_ID and XERO_CLIENT_SECRET are required. ' +
          'Register an app at https://developer.xero.com and set the credentials in .env.',
      );
    }
    if (!tenantId) {
      throw new Error(
        'XeroBillsClient: XERO_TENANT_ID is required (Corecast Pty Ltd = 024c8088-8095-482b-906e-f72d9b8acaee).',
      );
    }

    this.tenantId = tenantId;
    this.tokenCacheFile = tokenCacheFile ? path.resolve(tokenCacheFile) : null;
    this._initialAccessToken = accessToken || null;
    this._initialRefreshToken = refreshToken || null;

    this.sdk = new XeroClient({
      clientId,
      clientSecret,
      redirectUris: ['http://localhost/xero-callback'],
      scopes,
    });

    this._ready = null;
  }

  /**
   * Ensure we have a valid access token. Loads cached refresh token from disk if
   * available, otherwise falls back to the value supplied in the constructor.
   */
  async init() {
    if (this._ready) return this._ready;
    this._ready = (async () => {
      let refreshToken = this._initialRefreshToken;
      let accessToken = this._initialAccessToken;

      if (this.tokenCacheFile) {
        const cached = await readTokenCache(this.tokenCacheFile);
        if (cached?.refresh_token) refreshToken = cached.refresh_token;
        if (cached?.access_token && !accessToken) accessToken = cached.access_token;
      }

      if (!refreshToken && !accessToken) {
        throw new Error(
          'XeroBillsClient: no access_token or refresh_token available. ' +
            'Run scripts/xero-auth.js to perform the initial OAuth handshake, ' +
            'or set XERO_REFRESH_TOKEN in .env.',
        );
      }

      // xero-node v15 requires initialize() before setTokenSet/refreshToken —
      // it builds the internal OpenID Connect client, without which
      // refreshToken() throws "Cannot read properties of undefined (reading 'refresh')".
      if (typeof this.sdk.initialize === 'function') {
        await this.sdk.initialize();
      }

      this.sdk.setTokenSet({
        access_token: accessToken || '',
        refresh_token: refreshToken || '',
      });

      if (refreshToken) {
        const newToken = await this.sdk.refreshToken();
        if (this.tokenCacheFile) await writeTokenCache(this.tokenCacheFile, newToken);
      }
    })();
    return this._ready;
  }

  /**
   * Search Xero contacts by name, returning the first match (case-insensitive).
   * Returns null if not found.
   */
  async findContactByName(name) {
    if (!name) return null;
    await this.init();
    const trimmed = String(name).trim();
    if (!trimmed) return null;
    // searchTerm spans Name, FirstName, LastName, Email — broader than where=Name=="X".
    const res = await this.sdk.accountingApi.getContacts(
      this.tenantId,
      undefined, // ifModifiedSince
      undefined, // where
      'Name ASC', // order
      undefined, // iDs
      undefined, // page
      false, // includeArchived
      false, // summaryOnly
      trimmed, // searchTerm
    );
    const contacts = res.body?.contacts || [];
    if (!contacts.length) return null;
    const lower = trimmed.toLowerCase();
    const exact = contacts.find((c) => (c.name || '').toLowerCase() === lower);
    return exact || contacts[0];
  }

  /**
   * Find a contact by name, or create one if none exists.
   */
  async findOrCreateContact({ name, email }) {
    if (!name) throw new Error('findOrCreateContact: name is required');
    const existing = await this.findContactByName(name);
    if (existing) return existing;

    const payload = { name: String(name).trim() };
    if (email) payload.emailAddress = email;
    const res = await this.sdk.accountingApi.createContacts(this.tenantId, {
      contacts: [payload],
    });
    const created = res.body?.contacts?.[0];
    if (!created) throw new Error(`Xero: failed to create contact "${name}"`);
    return created;
  }

  /**
   * Look for an existing ACCPAY (bill) for the given contact within a date window.
   * Used to avoid creating duplicate drafts on re-runs of the weekly job.
   *
   * @param {object} args
   * @param {string} args.contactId
   * @param {string} args.dateFrom - ISO yyyy-mm-dd, inclusive lower bound
   * @param {string} [args.dateTo] - ISO yyyy-mm-dd, inclusive upper bound
   * @param {string} [args.reference] - if provided, also requires a Reference match
   */
  async findExistingBill({ contactId, dateFrom, dateTo, reference }) {
    if (!contactId) throw new Error('findExistingBill: contactId is required');
    if (!dateFrom) throw new Error('findExistingBill: dateFrom is required');
    await this.init();

    const clauses = [`Type=="ACCPAY"`, `Date>=DateTime(${dateParts(dateFrom)})`];
    if (dateTo) clauses.push(`Date<=DateTime(${dateParts(dateTo)})`);
    if (reference) clauses.push(`Reference=="${reference.replace(/"/g, '\\"')}"`);
    const where = clauses.join('&&');

    const res = await this.sdk.accountingApi.getInvoices(
      this.tenantId,
      undefined,                              // ifModifiedSince
      where,
      'Date DESC',
      undefined,                              // iDs
      undefined,                              // invoiceNumbers
      [contactId],                            // contactIDs
      ['DRAFT', 'SUBMITTED', 'AUTHORISED', 'PAID'], // statuses — excludes DELETED and VOIDED
    );
    const invoices = res.body?.invoices || [];
    return invoices[0] || null;
  }

  /**
   * Create a single DRAFT Accounts Payable invoice (bill).
   *
   * @param {object} bill
   * @param {object} bill.contact - { contactID }
   * @param {string} bill.date - yyyy-mm-dd
   * @param {string} bill.dueDate - yyyy-mm-dd
   * @param {string} bill.description
   * @param {number} bill.amount - tax-inclusive total (matches Tax Inclusive setting on org)
   * @param {string} bill.accountCode
   * @param {string} bill.reference
   * @param {string} [bill.taxType='INPUT'] - GST on purchases
   */
  async createDraftBill(bill) {
    await this.init();
    const required = ['contact', 'date', 'amount', 'accountCode'];
    for (const k of required) {
      if (bill[k] == null) throw new Error(`createDraftBill: ${k} is required`);
    }

    const payload = {
      type: 'ACCPAY',
      status: 'DRAFT',
      contact: { contactID: bill.contact.contactID },
      date: bill.date,
      dueDate: bill.dueDate || addDays(bill.date, 7),
      lineAmountTypes: 'Inclusive',
      lineItems: [
        {
          description: bill.description || `Bill for ${bill.reference || bill.date}`,
          quantity: 1,
          unitAmount: round2(bill.amount),
          accountCode: bill.accountCode,
          taxType: bill.taxType || 'INPUT',
        },
      ],
      reference: bill.reference || '',
      currencyCode: 'AUD',
    };

    const res = await this.sdk.accountingApi.createInvoices(this.tenantId, {
      invoices: [payload],
    });
    const created = res.body?.invoices?.[0];
    if (!created) throw new Error('Xero: createInvoices returned no invoice');
    return created;
  }
}

function dateParts(iso) {
  const [y, m, d] = iso.split('-').map((s) => parseInt(s, 10));
  return `${y}, ${m}, ${d}`;
}

function addDays(iso, days) {
  const d = new Date(iso + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function round2(n) {
  return Math.round(Number(n) * 100) / 100;
}

async function readTokenCache(file) {
  try {
    const raw = await fs.readFile(file, 'utf8');
    return JSON.parse(raw);
  } catch (err) {
    if (err.code === 'ENOENT') return null;
    throw err;
  }
}

async function writeTokenCache(file, tokenSet) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  const data = {
    access_token: tokenSet.access_token,
    refresh_token: tokenSet.refresh_token,
    expires_at: tokenSet.expires_at,
    updated_at: new Date().toISOString(),
  };
  await fs.writeFile(file, JSON.stringify(data, null, 2) + '\n');
}

module.exports = { XeroBillsClient, DEFAULT_SCOPES };
