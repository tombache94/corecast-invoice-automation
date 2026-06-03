const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const lockfile = require('proper-lockfile');

const VALID_STATUSES = [
  'pending',
  'validated',
  'mismatch',
  'review_needed',
  'paid',
  'billed', // draft bill created in Xero
  'bill_duplicate', // bill already existed in Xero on re-run
];

function makeId() {
  return `inv_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
}

function normalizeDate(value, field) {
  if (value == null || value === '') return null;
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) {
    throw new Error(`Invalid ${field}: ${value}`);
  }
  return d.toISOString().slice(0, 10);
}

function assertStatus(status) {
  if (!VALID_STATUSES.includes(status)) {
    throw new Error(`Invalid status "${status}". Must be one of: ${VALID_STATUSES.join(', ')}`);
  }
}

async function ensureFile(file) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  try {
    await fs.access(file);
  } catch {
    await fs.writeFile(file, JSON.stringify({ invoices: [] }, null, 2) + '\n');
  }
}

async function readData(file) {
  const raw = await fs.readFile(file, 'utf8');
  if (!raw.trim()) return { invoices: [] };
  const parsed = JSON.parse(raw);
  if (!Array.isArray(parsed.invoices)) return { invoices: [] };
  return parsed;
}

async function writeDataAtomic(file, data) {
  const tmp = `${file}.tmp.${process.pid}.${crypto.randomBytes(3).toString('hex')}`;
  await fs.writeFile(tmp, JSON.stringify(data, null, 2) + '\n');
  await fs.rename(tmp, file);
}

async function withWriteLock(file, fn) {
  await ensureFile(file);
  const release = await lockfile.lock(file, {
    retries: { retries: 50, factor: 1.3, minTimeout: 25, maxTimeout: 200 },
    stale: 10_000,
    realpath: false,
  });
  try {
    return await fn();
  } finally {
    await release();
  }
}

class InvoiceStore {
  constructor(file) {
    this.file = path.resolve(file);
  }

  async addInvoice(input = {}) {
    const status = input.status || 'pending';
    assertStatus(status);

    return withWriteLock(this.file, async () => {
      const data = await readData(this.file);
      const now = new Date().toISOString();
      const record = {
        id: input.id || makeId(),
        sender: input.sender ?? null,
        amount: input.amount ?? null,
        date: normalizeDate(input.date, 'date'),
        hoursClaimed: input.hoursClaimed ?? null,
        weekEnding: normalizeDate(input.weekEnding, 'weekEnding'),
        status,
        notes: input.notes || '',
        // Dedup key — Outlook's internetMessageId. Globally unique per email
        // and stable across reads. Set by the ingest pipeline so the matcher's
        // re-runs don't create duplicate records.
        internetMessageId: input.internetMessageId || null,
        outlookMessageId: input.outlookMessageId || null,
        mailbox: input.mailbox || null,
        createdAt: input.createdAt || now,
        updatedAt: now,
      };
      data.invoices.push(record);
      await writeDataAtomic(this.file, data);
      return record;
    });
  }

  /**
   * Has an invoice with this Outlook internetMessageId already been stored?
   * Returns true on match, false otherwise (including the no-id case).
   */
  async hasInternetMessageId(id) {
    if (!id) return false;
    await ensureFile(this.file);
    const data = await readData(this.file);
    return data.invoices.some((i) => i.internetMessageId === id);
  }

  async getInvoice(id) {
    await ensureFile(this.file);
    const data = await readData(this.file);
    return data.invoices.find((i) => i.id === id) || null;
  }

  async updateStatus(id, status, { note, fields } = {}) {
    assertStatus(status);
    return withWriteLock(this.file, async () => {
      const data = await readData(this.file);
      const inv = data.invoices.find((i) => i.id === id);
      if (!inv) throw new Error(`Invoice not found: ${id}`);
      inv.status = status;
      inv.updatedAt = new Date().toISOString();
      if (note) inv.notes = inv.notes ? `${inv.notes}\n${note}` : note;
      if (fields && typeof fields === 'object') Object.assign(inv, fields);
      await writeDataAtomic(this.file, data);
      return inv;
    });
  }

  async getByStatus(status) {
    assertStatus(status);
    await ensureFile(this.file);
    const data = await readData(this.file);
    return data.invoices.filter((i) => i.status === status);
  }

  async getAllForWeek(weekEnding) {
    const target = normalizeDate(weekEnding, 'weekEnding');
    await ensureFile(this.file);
    const data = await readData(this.file);
    return data.invoices.filter((i) => i.weekEnding === target);
  }

  async getAllInRange(startDate, endDate) {
    const start = normalizeDate(startDate, 'startDate');
    const end = normalizeDate(endDate, 'endDate');
    await ensureFile(this.file);
    const data = await readData(this.file);
    return data.invoices.filter((i) => {
      if (!i.weekEnding) return false;
      return i.weekEnding >= start && i.weekEnding <= end;
    });
  }
}

module.exports = { InvoiceStore, VALID_STATUSES };
