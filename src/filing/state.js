const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const lockfile = require('proper-lockfile');

async function ensureFile(file) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  try {
    await fs.access(file);
  } catch {
    await fs.writeFile(file, JSON.stringify({ filed: [] }, null, 2) + '\n');
  }
}

async function readData(file) {
  const raw = await fs.readFile(file, 'utf8');
  if (!raw.trim()) return { filed: [] };
  const parsed = JSON.parse(raw);
  if (!Array.isArray(parsed.filed)) return { filed: [] };
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

// Dedup key uses fields that are stable across mailboxes so the same email
// arriving at multiple inboxes (cc, distribution list, etc.) only files once.
//
// - internetMessageId is the RFC 5322 Message-ID header — identical in every
//   recipient's copy.
// - attachmentName + attachmentSize identify the attachment within that message.
//
// Falls back to a per-mailbox key when internetMessageId is missing (rare —
// drafts, malformed forwards). That degrades safely: still dedupes within
// one mailbox, just not across them.
function makeKey({ internetMessageId, attachmentName, attachmentSize, mailbox, messageId, attachmentId }) {
  if (internetMessageId) {
    return `imid:${internetMessageId}::${attachmentName}::${attachmentSize ?? '?'}`;
  }
  return `mbox:${mailbox}::${messageId}::${attachmentId}`;
}

class FilingStore {
  constructor(file) {
    this.file = path.resolve(file);
  }

  async hasBeenFiled(identity) {
    await ensureFile(this.file);
    const data = await readData(this.file);
    const key = makeKey(identity);
    return data.filed.some((r) => r.key === key);
  }

  async record(entry) {
    return withWriteLock(this.file, async () => {
      const data = await readData(this.file);
      const key = makeKey(entry);
      data.filed.push({
        key,
        // Identity fields used to build the key:
        internetMessageId: entry.internetMessageId || null,
        attachmentSize: entry.attachmentSize ?? null,
        // Per-mailbox forensic trace (handy for debugging — not part of the key):
        mailbox: entry.mailbox,
        messageId: entry.messageId,
        attachmentId: entry.attachmentId,
        attachmentName: entry.attachmentName,
        subject: entry.subject || '',
        from: entry.from || '',
        keyword: entry.keyword || null,
        destination: entry.destination,
        outcome: entry.outcome,
        webUrl: entry.webUrl || null,
        filedAt: new Date().toISOString(),
        notes: entry.notes || '',
      });
      await writeDataAtomic(this.file, data);
    });
  }
}

module.exports = { FilingStore };
