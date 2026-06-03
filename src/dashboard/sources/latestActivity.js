/**
 * Compute the most recent email activity per active job by scanning
 * Outlook inboxes and classifying messages against the orchestrator's
 * routing rules. Used to populate the "Latest" field on dashboard job
 * cards — independent of Monday Notes (which may lag behind reality).
 *
 * Returns: { [jobId]: { receivedAt, from, fromName, subject, preview, summary } }
 * `summary` is a one-liner formatted in Cowork's note style for direct
 * display on the dashboard.
 */

const { OutlookClient } = require('../../email/outlook');
const { classifyEmail } = require('../../orchestrator/classify');
const { loadRouting } = require('../../orchestrator/routing');

const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const DEFAULT_MAILBOXES = ['tom@corecastconcrete.com.au', 'accounts@corecastconcrete.com.au'];
const LOOKBACK_DAYS = 14;
const TOP_PER_MAILBOX = 100;

function formatDate(iso) {
  const d = new Date(iso);
  const day = String(d.getDate()).padStart(2, '0');
  return `${day}-${MONTHS[d.getMonth()]}-${d.getFullYear()}`;
}

function buildSummary(msg) {
  const date = formatDate(msg.receivedAt);
  const senderName = (msg.fromName || msg.from || 'unknown').trim();
  const subj = (msg.subject || '(no subject)').trim();
  // Microsoft Graph's bodyPreview is the first ~255 chars of the email body
  // as plain text — good enough as an at-a-glance summary without needing
  // to fetch the full message or call an LLM. Collapse whitespace and trim.
  const previewRaw = (msg.preview || '').replace(/\s+/g, ' ').trim();
  // Strip a few common forwarded-message preambles that add noise.
  const preview = previewRaw
    .replace(/^Get Outlook for (iOS|Android|Mac|Windows)\s*_+\s*/i, '')
    .replace(/^From:.*?Subject:.*?(?=\s+\w)/, '')
    .trim();
  const previewShort = preview.length > 180 ? preview.slice(0, 180).trim() + '…' : preview;
  const tail = previewShort ? ` · ${previewShort}` : '';
  return `[${date}] EMAIL: ${senderName} — "${subj}"${tail}`;
}

async function fetchLatestActivityPerJob(activeJobs, deps = {}) {
  if (!activeJobs || activeJobs.length === 0) return {};
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) return {};

  const outlook = deps.outlook || new OutlookClient({ tenantId, clientId, clientSecret });
  const routing = deps.routing || loadRouting();
  const sinceISO = new Date(Date.now() - LOOKBACK_DAYS * 24 * 3600 * 1000).toISOString();
  const mailboxes = deps.mailboxes || DEFAULT_MAILBOXES;

  const all = [];
  for (const mb of mailboxes) {
    try {
      const msgs = await outlook.listRecentMessages(mb, { sinceISO, top: TOP_PER_MAILBOX });
      for (const m of msgs) all.push({ ...m, mailbox: mb });
    } catch (_) {
      // best-effort — skip the mailbox if it fails
    }
  }

  // Sort newest-first so the first classified hit per job is the latest.
  all.sort((a, b) => new Date(b.receivedAt) - new Date(a.receivedAt));

  const latestByJobId = {};
  for (const msg of all) {
    const match = classifyEmail(msg, activeJobs, routing);
    if (!match) continue;
    if (latestByJobId[match.jobId]) continue;
    latestByJobId[match.jobId] = {
      receivedAt: msg.receivedAt,
      from: msg.from,
      fromName: msg.fromName,
      subject: msg.subject,
      preview: (msg.preview || '').slice(0, 140),
      summary: buildSummary(msg),
    };
    if (Object.keys(latestByJobId).length === activeJobs.length) break;
  }

  return latestByJobId;
}

module.exports = { fetchLatestActivityPerJob };
