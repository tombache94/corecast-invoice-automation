/**
 * LLM-based fallback extraction for invoice fields.
 *
 * Used by src/email/parser.js when regex extraction fails to find an
 * amount or hours value. Sends the already-gathered raw invoice text
 * (subject + body + PDF attachment text) to Claude Haiku 4.5 with a
 * structured-output schema, returns parsed { amount, hoursClaimed,
 * weekEnding } back.
 *
 * Cost: roughly $0.001-0.003 per invoice at typical sizes.
 *
 * Graceful: if ANTHROPIC_API_KEY is not set, returns null so the parser
 * keeps working with regex-only behaviour. No errors thrown — invoices
 * just stay in review_needed if extraction fails.
 */

const Anthropic = require('@anthropic-ai/sdk');
const { z } = require('zod');
const { zodOutputFormat } = require('@anthropic-ai/sdk/helpers/zod');

const SYSTEM_PROMPT = `You extract three structured fields from CoreCast Concrete invoice emails.

CoreCast (a Perth precast concrete manufacturer) receives invoices at accounts@corecastconcrete.com.au from two kinds of senders:
  1. LABOUR invoices from subcontractors (Aaron Norris, James Elliott, Alani Fonua, Veronty Pty Ltd, etc.) — claim hours worked
  2. SUPPLIER bills from materials/equipment vendors (Nexus Construction, Mogul Reinforcement, AB Precast, etc.) — invoice for goods/services with no hours

Extract:

- amount: Total amount due in AUD as a plain number (no $, no commas). Prefer values labelled "Amount Due", "Total", "Grand Total", "Sub-Total". If both ex-GST and inc-GST are shown, return the EX-GST amount (this is what reconciles to Xero). Return null if no amount is extractable.

- hoursClaimed: Total labour hours claimed as a decimal number. Look for "X hours", "X hrs", "Total Hours", "Hours Worked". If the invoice is a SUPPLIER bill with no hours mentioned, return null. Return null if hours not extractable.

- weekEnding: The end date of the WORK PERIOD the invoice covers, formatted YYYY-MM-DD. Look for "week ending", "period ending", "for week of", or infer from a date range (e.g. "5 May - 11 May 2026" → weekEnding 2026-05-11). CoreCast's pay cycle runs Tuesday to Monday, so a Monday end-date is normal. CRITICAL: NEVER use a "due date", "payment due date", "pay by", or "due by" date — those are payment deadlines, not work periods. For SUPPLIER bills (concrete, materials, equipment, freight, professional services) that invoice for goods rather than time worked, return null. If only an invoice/issue date is present with no work period stated, return null.

Be conservative. Return null for fields that are genuinely missing rather than guessing. Provide brief reasoning so a human can audit.`;

const InvoiceFieldsSchema = z.object({
  amount: z.number().nullable(),
  hoursClaimed: z.number().nullable(),
  weekEnding: z.string().nullable(),
  reasoning: z.string(),
});

let cachedClient = null;

function getClient() {
  if (cachedClient) return cachedClient;
  if (!process.env.ANTHROPIC_API_KEY) return null;
  cachedClient = new Anthropic();
  return cachedClient;
}

async function extractInvoiceFields({ rawContent, sender, receivedDateTime }) {
  const client = getClient();
  if (!client) return { skipped: 'no-api-key' };
  if (!rawContent || rawContent.trim().length < 20) return { skipped: 'empty-content' };

  const userMessage = [
    `Sender: ${sender || 'unknown'}`,
    `Email received: ${receivedDateTime || 'unknown'}`,
    '',
    'Invoice content (subject + body + PDF attachment text):',
    '---',
    rawContent.slice(0, 30000),
    '---',
    '',
    'Extract amount, hoursClaimed, weekEnding per the system instructions.',
  ].join('\n');

  try {
    const response = await client.messages.parse({
      model: 'claude-haiku-4-5',
      max_tokens: 1024,
      // cache_control marker present but won't actually fire — Haiku 4.5's
      // minimum cacheable prefix is ~4096 tokens and this system prompt is
      // shorter. Kept for free upgrade if the prompt grows (e.g. few-shot
      // examples added later) or if model is swapped to one with a smaller
      // minimum.
      system: [
        {
          type: 'text',
          text: SYSTEM_PROMPT,
          cache_control: { type: 'ephemeral' },
        },
      ],
      output_config: { format: zodOutputFormat(InvoiceFieldsSchema) },
      messages: [{ role: 'user', content: userMessage }],
    });

    if (!response.parsed_output) {
      return { error: `parse failed; stop_reason=${response.stop_reason}` };
    }

    const out = response.parsed_output;
    return {
      amount: out.amount,
      hoursClaimed: out.hoursClaimed,
      weekEnding: out.weekEnding,
      reasoning: out.reasoning,
      usage: response.usage,
    };
  } catch (err) {
    if (err instanceof Anthropic.RateLimitError) return { error: 'rate-limited' };
    if (err instanceof Anthropic.APIError) return { error: `api ${err.status}: ${err.message}` };
    return { error: err.message };
  }
}

module.exports = { extractInvoiceFields, SYSTEM_PROMPT, InvoiceFieldsSchema };
