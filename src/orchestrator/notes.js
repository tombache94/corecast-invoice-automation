/**
 * Note formatting and Monday Notes-column writing.
 *
 * formatNote(message)       — produce a single Cowork-style note line from
 *                              an Outlook message.
 * appendNotesToJob(itemId,
 *   currentNotes, entries)  — append new entries to the existing notes
 *                              long-text column on Monday item. Existing
 *                              content is preserved (we prepend new at top
 *                              so it reads newest-first like Cowork does).
 */

const https = require('https');

const BOARD_ID = 5027051539;
const NOTES_COLUMN_ID = 'long_text_mm17a895';

const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function formatDate(d) {
  const day = String(d.getDate()).padStart(2, '0');
  const mon = MONTHS[d.getMonth()];
  const year = d.getFullYear();
  return `${day}-${mon}-${year}`;
}

function formatNote(message) {
  const received = message.receivedAt ? new Date(message.receivedAt) : new Date();
  const date = formatDate(received);
  const senderLabel = message.fromName && message.from
    ? `${message.fromName} <${message.from}>`
    : message.from || 'unknown sender';
  const subj = (message.subject || '(no subject)').trim();
  const preview = (message.preview || '').replace(/\s+/g, ' ').trim().slice(0, 140);
  return `[${date}] EMAIL: From ${senderLabel} — "${subj}"${preview ? ` (preview: ${preview}…)` : ''}`;
}

async function appendNotesToJob(itemId, currentNotes, newEntries) {
  if (!newEntries || newEntries.length === 0) return;
  const combined = newEntries.join('\n\n');
  const updated = currentNotes && currentNotes.trim()
    ? `${combined}\n\n${currentNotes}`
    : combined;
  await writeLongTextColumn(itemId, NOTES_COLUMN_ID, updated);
}

function writeLongTextColumn(itemId, columnId, text) {
  const apiKey = process.env.MONDAY_API_KEY;
  if (!apiKey) throw new Error('Missing MONDAY_API_KEY in .env');

  const valueJson = JSON.stringify({ text });
  const mutation = `mutation { change_column_value(board_id: ${BOARD_ID}, item_id: ${itemId}, column_id: "${columnId}", value: ${JSON.stringify(valueJson)}) { id } }`;

  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ query: mutation });
    const req = https.request(
      {
        hostname: 'api.monday.com',
        path: '/v2',
        method: 'POST',
        headers: {
          Authorization: apiKey,
          'Content-Type': 'application/json',
          'API-Version': '2024-10',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => {
          try {
            const parsed = JSON.parse(data);
            if (parsed.errors) return reject(new Error('Monday API error: ' + JSON.stringify(parsed.errors)));
            resolve(parsed.data);
          } catch (e) {
            reject(new Error(`Monday parse error: ${e.message} body=${data.slice(0, 200)}`));
          }
        });
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

module.exports = { formatNote, appendNotesToJob, BOARD_ID, NOTES_COLUMN_ID };
