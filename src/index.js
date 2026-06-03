const { config, assertMicrosoftCreds } = require('./config');
const { OutlookClient } = require('./email/outlook');
const { isInvoiceEmail } = require('./email/filter');
const { parseInvoiceMessage } = require('./email/parser');
const { InvoiceStore } = require('./storage/invoices');

const REVIEW_THRESHOLD = 0.7;

function decideStatus(parsed) {
  const reasons = [];
  if (parsed.amount == null) reasons.push('amount not extracted');
  else if (parsed.confidence.amount < REVIEW_THRESHOLD)
    reasons.push(`low amount confidence (${parsed.confidence.amount})`);

  if (parsed.hoursClaimed == null) reasons.push('hours not extracted');
  else if (parsed.confidence.hoursClaimed < REVIEW_THRESHOLD)
    reasons.push(`low hours confidence (${parsed.confidence.hoursClaimed})`);

  if (parsed.date == null) reasons.push('date not extracted');

  return reasons.length
    ? { status: 'review_needed', notes: reasons.join('; ') }
    : { status: 'pending', notes: 'awaiting Connecteam validation' };
}

async function ingest(args = {}, deps = {}) {
  assertMicrosoftCreds();

  const mailbox = args.mailbox || config.mailboxes[0];
  const limit = args.limit || config.mailboxFetchLimit;
  // Default: only look back 14 days. Stops one-off ingestion runs from
  // re-scanning the entire inbox each time, AND limits the dedup-check
  // surface area to recent messages.
  const sinceDays = args.sinceDays != null ? args.sinceDays : 14;
  const sinceISO = sinceDays > 0
    ? new Date(Date.now() - sinceDays * 24 * 3600 * 1000).toISOString()
    : null;

  let outlook = deps.outlook;
  if (!outlook) {
    outlook = new OutlookClient(config.microsoft);
  }
  const store = deps.store || new InvoiceStore(config.storage.invoicesFile);

  console.log(`\nIngesting from ${mailbox} (top ${limit}${sinceISO ? `, since ${sinceISO.slice(0,10)}` : ''}, read-only)...`);

  const summaries = await outlook.listRecentMessages(mailbox, { top: limit, sinceISO });
  console.log(`Found ${summaries.length} message(s) in inbox.\n`);

  const stats = { total: summaries.length, filtered: 0, stored: 0, pending: 0, review_needed: 0, duplicates: 0, errors: 0 };
  const reviews = [];

  for (let i = 0; i < summaries.length; i++) {
    const s = summaries[i];
    console.log(`[${i + 1}/${summaries.length}] ${s.receivedAt}  ${s.from || '(no sender)'}`);
    console.log(`        subj: ${s.subject || '(no subject)'}`);

    // Pre-filter: skip obvious non-invoices
    if (!isInvoiceEmail(s)) {
      console.log(`        SKIPPED: does not match invoice heuristic\n`);
      continue;
    }

    stats.filtered++;

    // Dedup by Outlook's internetMessageId — globally unique per email and
    // stable across reads. Skip if already stored.
    if (s.internetMessageId && await store.hasInternetMessageId(s.internetMessageId)) {
      console.log(`        SKIPPED: already in store (internetMessageId match)\n`);
      stats.duplicates++;
      continue;
    }

    try {
      const full = await outlook.getMessage(mailbox, s.id, { includeAttachments: true });
      const parsed = await parseInvoiceMessage(full);
      const decision = decideStatus(parsed);

      const stored = await store.addInvoice({
        sender: parsed.sender,
        amount: parsed.amount,
        date: parsed.date,
        hoursClaimed: parsed.hoursClaimed,
        weekEnding: parsed.weekEnding,
        status: decision.status,
        notes: decision.notes,
        internetMessageId: s.internetMessageId,
        outlookMessageId: s.id,
        mailbox,
      });

      const c = parsed.confidence;
      console.log(
        `        parsed:  amt=${(parsed.amount ?? '?').toString().padEnd(8)}(c${c.amount})  ` +
        `hrs=${(parsed.hoursClaimed ?? '?').toString().padEnd(5)}(c${c.hoursClaimed})  ` +
        `date=${(parsed.date ?? '?').toString().padEnd(11)}(c${c.date})  ` +
        `week=${parsed.weekEnding ?? '?'}`
      );
      console.log(`        stored:  ${stored.status}  [${stored.id}]`);

      if (decision.status === 'review_needed') {
        console.log(`        review:  ${decision.notes}`);
        reviews.push({ id: stored.id, from: s.from, subject: s.subject, reason: decision.notes });
      }

      stats.stored++;
      stats[decision.status]++;
    } catch (err) {
      console.log(`        ERROR: ${err.message}`);
      stats.errors++;
    }
    console.log('');
  }

  console.log('=== Summary ===');
  console.log(`  Messages scanned:   ${stats.total}`);
  console.log(`  Matched filter:     ${stats.filtered}`);
  console.log(`  Skipped (dup):      ${stats.duplicates}`);
  console.log(`  Stored as pending:  ${stats.pending}`);
  console.log(`  Stored for review:  ${stats.review_needed}`);
  console.log(`  Errors:             ${stats.errors}`);
  console.log(`  Storage file:       ${config.storage.invoicesFile}`);

  if (reviews.length) {
    console.log(`\n  Items needing review:`);
    for (const r of reviews) {
      console.log(`    - ${r.from} | ${r.subject}`);
      console.log(`      reason: ${r.reason}`);
    }
  }

  return stats;
}

if (require.main === module) {
  const args = {};
  if (process.argv.includes('--mailbox') || process.argv.includes('-m')) {
    const idx = process.argv.indexOf('--mailbox') !== -1 ? process.argv.indexOf('--mailbox') : process.argv.indexOf('-m');
    args.mailbox = process.argv[idx + 1];
  }
  if (process.argv.includes('--limit') || process.argv.includes('-l')) {
    const idx = process.argv.indexOf('--limit') !== -1 ? process.argv.indexOf('--limit') : process.argv.indexOf('-l');
    args.limit = parseInt(process.argv[idx + 1], 10);
  }
  if (process.argv.includes('--since-days')) {
    const idx = process.argv.indexOf('--since-days');
    args.sinceDays = parseInt(process.argv[idx + 1], 10);
  }

  (async () => {
    try {
      await ingest(args);
    } catch (err) {
      console.error('Fatal:', err.message);
      if (process.env.DEBUG) console.error(err.stack);
      process.exit(1);
    }
  })();
}

module.exports = { ingest, decideStatus };