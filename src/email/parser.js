const MONTHS = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
};

const RE_LABELLED_TOTAL = /\b(?:grand\s+total|amount\s+due|total)\b\s*:?\s*(?:AUD?\$|A\$|\$)?\s*([\d,]+(?:\.\d{1,2})?)/gi;
const RE_DOLLAR = /(?:AUD?\$|A\$|\$)\s?([\d,]+(?:\.\d{1,2})?)/g;
const RE_HOURS = /(\d+(?:\.\d+)?)\s*(hours?|hrs?|h)\b/gi;
const RE_DATE_ISO = /\b(\d{4}-\d{2}-\d{2})\b/;
const RE_DATE_DMY_TEXT = /\b(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{4})\b/i;
const RE_DATE_MDY_TEXT = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})\b/i;
const RE_DATE_NUMERIC = /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/;

function htmlToText(html) {
  if (!html) return '';
  let t = html;
  t = t.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ');
  t = t.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ');
  t = t.replace(/<br\s*\/?>/gi, '\n');
  t = t.replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n');
  t = t.replace(/<\/td>\s*<td[^>]*>/gi, '\t');
  t = t.replace(/<[^>]+>/g, ' ');
  t = t
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/gi, '"');
  t = t.replace(/[\t ]+/g, ' ').replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n');
  return t.trim();
}

function parseAmount(text) {
  const labelled = [];
  let m;
  RE_LABELLED_TOTAL.lastIndex = 0;
  while ((m = RE_LABELLED_TOTAL.exec(text)) !== null) {
    const v = parseFloat(m[1].replace(/,/g, ''));
    if (Number.isFinite(v) && v > 0) labelled.push(v);
  }
  if (labelled.length) {
    return { value: Math.max(...labelled), confidence: 0.95, source: 'labelled-total' };
  }

  const all = [];
  RE_DOLLAR.lastIndex = 0;
  while ((m = RE_DOLLAR.exec(text)) !== null) {
    const v = parseFloat(m[1].replace(/,/g, ''));
    if (Number.isFinite(v) && v > 0) all.push(v);
  }
  if (!all.length) return { value: null, confidence: 0, source: 'none' };
  if (all.length === 1) return { value: all[0], confidence: 0.8, source: 'single-dollar' };
  return { value: Math.max(...all), confidence: 0.5, source: 'max-of-many' };
}

function parseHours(text) {
  const matches = [];
  let m;
  RE_HOURS.lastIndex = 0;
  while ((m = RE_HOURS.exec(text)) !== null) {
    const v = parseFloat(m[1]);
    if (Number.isFinite(v) && v > 0 && v < 200) {
      matches.push({ value: v, unit: m[2].toLowerCase() });
    }
  }
  if (!matches.length) return { value: null, confidence: 0, source: 'none' };
  const labelled = matches.filter((x) => x.unit !== 'h');
  if (labelled.length === 1) {
    return { value: labelled[0].value, confidence: 0.95, source: 'single-labelled' };
  }
  if (labelled.length > 1) {
    return {
      value: Math.max(...labelled.map((x) => x.value)),
      confidence: 0.7,
      source: 'multiple-labelled-took-max',
    };
  }
  return { value: matches[0].value, confidence: 0.5, source: 'bare-h-suffix' };
}

function tryISODate(text) {
  const m = text.match(RE_DATE_ISO);
  if (!m) return null;
  const d = new Date(m[1] + 'T00:00:00Z');
  return Number.isNaN(d.getTime()) ? null : { iso: m[1], confidence: 0.95, source: 'iso' };
}

function tryTextDate(text) {
  let m = text.match(RE_DATE_DMY_TEXT);
  if (m) {
    const day = parseInt(m[1], 10);
    const mo = MONTHS[m[2].toLowerCase().slice(0, 3)];
    const y = parseInt(m[3], 10);
    const d = new Date(Date.UTC(y, mo, day));
    if (!Number.isNaN(d.getTime())) {
      return { iso: d.toISOString().slice(0, 10), confidence: 0.9, source: 'text-dmy' };
    }
  }
  m = text.match(RE_DATE_MDY_TEXT);
  if (m) {
    const mo = MONTHS[m[1].toLowerCase().slice(0, 3)];
    const day = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);
    const d = new Date(Date.UTC(y, mo, day));
    if (!Number.isNaN(d.getTime())) {
      return { iso: d.toISOString().slice(0, 10), confidence: 0.9, source: 'text-mdy' };
    }
  }
  return null;
}

function tryNumericDate(text) {
  const m = text.match(RE_DATE_NUMERIC);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10) - 1;
  let y = parseInt(m[3], 10);
  if (y < 100) y += 2000;
  if (day < 1 || day > 31 || mo < 0 || mo > 11) return null;
  const d = new Date(Date.UTC(y, mo, day));
  if (Number.isNaN(d.getTime())) return null;
  return { iso: d.toISOString().slice(0, 10), confidence: 0.7, source: 'numeric-dmy' };
}

function parseDate(text, fallbackISO) {
  const candidate = tryISODate(text) || tryTextDate(text) || tryNumericDate(text);

  if (candidate) {
    // Reject dates that are implausibly far in the future relative to when the email
    // arrived. MYOB / Xero reminder emails show the payment due date (often 30–50 days
    // out) prominently in the body, and the regex picks that up instead of the invoice
    // date. Standard Australian payment terms are 7–30 days; anything more than 40 days
    // in the future is almost certainly a payment deadline — fall back to email received date.
    if (fallbackISO) {
      const receivedMs = new Date(fallbackISO).getTime();
      const extractedMs = new Date(candidate.iso + 'T00:00:00Z').getTime();
      if (!Number.isNaN(receivedMs) && extractedMs - receivedMs > 40 * 24 * 3600 * 1000) {
        return { iso: fallbackISO.slice(0, 10), confidence: 0.3, source: 'email-received-date' };
      }
    }
    return candidate;
  }

  return fallbackISO
    ? { iso: fallbackISO.slice(0, 10), confidence: 0.3, source: 'email-received-date' }
    : { iso: null, confidence: 0, source: 'none' };
}

function nearestFriday(yyyymmdd) {
  if (!yyyymmdd) return null;
  const d = new Date(yyyymmdd + 'T00:00:00Z');
  if (Number.isNaN(d.getTime())) return null;
  const offsets = { 0: -2, 1: -3, 2: 3, 3: 2, 4: 1, 5: 0, 6: -1 };
  d.setUTCDate(d.getUTCDate() + offsets[d.getUTCDay()]);
  return d.toISOString().slice(0, 10);
}

async function defaultExtractPdf(buffer) {
  const pdf = require('pdf-parse');
  const result = await pdf(buffer);
  return result.text || '';
}

async function gatherText(message, extractPdf) {
  const parts = [];
  if (message.subject) parts.push(`SUBJECT: ${message.subject}`);

  const body = message.body || {};
  const ct = (body.contentType || '').toLowerCase();
  if (ct === 'html') parts.push(htmlToText(body.content || ''));
  else if (body.content) parts.push(body.content);

  for (const att of message.attachments || []) {
    if (!att.contentBytes) continue;
    const attCt = (att.contentType || '').toLowerCase();
    const name = (att.name || '').toLowerCase();
    if (attCt.includes('pdf') || name.endsWith('.pdf')) {
      try {
        const buf = Buffer.from(att.contentBytes, 'base64');
        const text = await extractPdf(buf);
        parts.push(`ATTACHMENT (${att.name}):\n${text}`);
      } catch (err) {
        parts.push(`ATTACHMENT (${att.name}): [pdf extract failed: ${err.message}]`);
      }
    } else if (attCt.startsWith('text/')) {
      parts.push(`ATTACHMENT (${att.name}):\n${Buffer.from(att.contentBytes, 'base64').toString('utf8')}`);
    }
  }
  return parts.join('\n\n').trim();
}

async function parseInvoiceMessage(message, { extractPdf = defaultExtractPdf, useLLMFallback = true } = {}) {
  const rawContent = await gatherText(message, extractPdf);

  const fromEmail = message.from?.emailAddress?.address || null;
  const fromName = message.from?.emailAddress?.name || null;
  const sender = fromName || (fromEmail ? fromEmail.split('@')[0] : null);
  const senderConfidence = fromName ? 1.0 : fromEmail ? 0.6 : 0;
  const emailConfidence = fromEmail ? 1.0 : 0;

  let amount = parseAmount(rawContent);
  const date = parseDate(rawContent, message.receivedDateTime);
  let hours = parseHours(rawContent);
  let weekEnding = nearestFriday(date.iso);
  let llmInfo = null;

  // LLM fallback — when regex couldn't extract an amount or hours, send the
  // already-gathered raw text (subject + body + PDF attachments) to Claude
  // Haiku and merge any fields it finds. Only fields the regex missed get
  // overwritten — confident regex matches are preserved. Skipped silently
  // if ANTHROPIC_API_KEY is not configured.
  if (useLLMFallback && (amount.value == null || hours.value == null)) {
    const { extractInvoiceFields } = require('./llmParser');
    const llm = await extractInvoiceFields({
      rawContent,
      sender,
      receivedDateTime: message.receivedDateTime,
    });
    if (llm && !llm.error && !llm.skipped) {
      llmInfo = { reasoning: llm.reasoning, usage: llm.usage };
      if (amount.value == null && llm.amount != null) {
        amount = { value: llm.amount, confidence: 0.85, source: 'llm' };
      }
      if (hours.value == null && llm.hoursClaimed != null) {
        hours = { value: llm.hoursClaimed, confidence: 0.85, source: 'llm' };
      }
      if (weekEnding == null && llm.weekEnding != null) {
        weekEnding = llm.weekEnding;
      }
    } else if (llm?.error) {
      llmInfo = { error: llm.error };
    } else if (llm?.skipped) {
      llmInfo = { skipped: llm.skipped };
    }
  }

  return {
    sender,
    senderEmail: fromEmail,
    amount: amount.value,
    date: date.iso,
    hoursClaimed: hours.value,
    weekEnding,
    rawContent,
    llmInfo,
    confidence: {
      sender: senderConfidence,
      senderEmail: emailConfidence,
      amount: amount.confidence,
      date: date.confidence,
      hoursClaimed: hours.confidence,
      weekEnding: date.confidence,
    },
    sources: {
      amount: amount.source,
      date: date.source,
      hoursClaimed: hours.source,
    },
  };
}

module.exports = {
  parseInvoiceMessage,
  htmlToText,
  nearestFriday,
  parseAmount,
  parseDate,
  parseHours,
};
