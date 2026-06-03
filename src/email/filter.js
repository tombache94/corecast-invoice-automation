/**
 * Invoice email filter — heuristic to separate invoices from noise.
 *
 * Three tiers, evaluated in order:
 *   1. Known sender allowlist (config/invoice-senders.json) — always passes
 *   2. Subject contains an invoice keyword
 *   3. Has a PDF/image attachment, OR body mentions $ + hours
 */

const fs = require('fs');
const path = require('path');

const INVOICE_KEYWORDS = ['invoice', 'tax invoice', 'inv ', 'labour', 'labor', 'hours worked'];
const SENDERS_FILE = path.resolve(__dirname, '../../config/invoice-senders.json');

// Emails from these domains are never invoices — block before allowlist check.
// post.xero.com: Xero sends payment notification emails that contain dollar amounts
//   and pass the keyword/attachment heuristics but are already in Xero.
// corecastconcrete.com.au: own-domain forwards/CCs get picked up as invoices.
const BLOCKED_SENDER_DOMAINS = ['post.xero.com', 'corecastconcrete.com.au'];

let cachedSenders = null;
function loadKnownSenders() {
  if (cachedSenders !== null) return cachedSenders;
  try {
    if (!fs.existsSync(SENDERS_FILE)) {
      cachedSenders = [];
      return cachedSenders;
    }
    const parsed = JSON.parse(fs.readFileSync(SENDERS_FILE, 'utf8'));
    cachedSenders = Array.isArray(parsed.senders)
      ? parsed.senders.map((s) => String(s).toLowerCase().trim()).filter(Boolean)
      : [];
  } catch (_) {
    cachedSenders = [];
  }
  return cachedSenders;
}

function isInvoiceEmail(message) {
  // Tier 0 — blocked sender domains (Xero notifications, own-domain forwards).
  const fromAddress = String(message.from || '').toLowerCase();
  const atIdx = fromAddress.lastIndexOf('@');
  if (atIdx !== -1) {
    const domain = fromAddress.slice(atIdx + 1);
    if (BLOCKED_SENDER_DOMAINS.some((d) => domain === d || domain.endsWith('.' + d))) {
      return false;
    }
  }

  // Tier 1 — known sender allowlist. Subs sending from personal Gmail with
  // vague subjects ("James Elliott") need to bypass the keyword/body heuristic.
  if (fromAddress && loadKnownSenders().includes(fromAddress)) return true;

  const subject = (message.subject || '').toLowerCase();
  const hasInvoiceKeyword = INVOICE_KEYWORDS.some((kw) => subject.includes(kw));

  const hasPdfAttachment = message.hasAttachments &&
    message.attachments?.some((a) =>
      a.contentType?.includes('pdf') ||
      a.name?.endsWith('.pdf')
    );

  const body = (message.bodyPreview || '').toLowerCase();
  const hasDollarSign = body.includes('$') || body.includes('aud');
  const hasHours = /\d+\s*(hours?|hrs?|h\b)/.test(body);

  return hasInvoiceKeyword || hasPdfAttachment || (hasDollarSign && hasHours);
}

module.exports = { isInvoiceEmail, loadKnownSenders };
