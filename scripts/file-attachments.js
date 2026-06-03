#!/usr/bin/env node
/**
 * Scan recent emails in the configured mailboxes and file each attachment
 * into SharePoint based on subject-line keyword matches. Designed to run
 * unattended from Windows Task Scheduler — see README/CLAUDE.md.
 *
 * Usage:
 *   node scripts/file-attachments.js
 *   node scripts/file-attachments.js --dry-run
 *   node scripts/file-attachments.js --mailbox tom@corecastconcrete.com.au
 *   node scripts/file-attachments.js --since 2026-05-01T00:00:00Z
 *   node scripts/file-attachments.js --lookback 12      # last 12 hours
 *   node scripts/file-attachments.js --lookback 0       # disable rolling window
 *   node scripts/file-attachments.js --limit 100
 *
 * Default: scans the last FILING_LOOKBACK_HOURS hours (24 unless overridden in .env).
 */

const { fileAttachments } = require('../src/filing/fileAttachments');

function getFlag(name, short) {
  const idx = process.argv.findIndex((a) => a === `--${name}` || (short && a === `-${short}`));
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function hasFlag(name) {
  return process.argv.includes(`--${name}`);
}

(async () => {
  const lookbackRaw = getFlag('lookback');
  const args = {
    dryRun: hasFlag('dry-run'),
    mailbox: getFlag('mailbox', 'm'),
    sinceISO: getFlag('since'),
    limit: getFlag('limit', 'l') ? parseInt(getFlag('limit', 'l'), 10) : undefined,
    lookbackHours: lookbackRaw !== undefined ? Number(lookbackRaw) : undefined,
  };
  try {
    await fileAttachments(args);
  } catch (err) {
    console.error('Fatal:', err.message);
    if (process.env.DEBUG) console.error(err.stack);
    process.exit(1);
  }
})();
