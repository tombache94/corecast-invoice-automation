#!/usr/bin/env node
/**
 * For each active Monday job whose SharePoint Link column is empty, find a
 * matching folder in CORECAST/📁 ACTIVE JOBS and write its webUrl to
 * link_mm36d0ea. Idempotent — never overwrites an existing link.
 *
 * Usage:
 *   node scripts/sync-sharepoint-links.js --dry-run     # preview only
 *   node scripts/sync-sharepoint-links.js               # write to Monday
 */

require('dotenv').config();
const { syncJobLinks } = require('../src/sharepoint/jobLinkSync');

const dryRun = process.argv.includes('--dry-run');

(async () => {
  try {
    await syncJobLinks({ dryRun });
  } catch (err) {
    console.error('Fatal:', err.message);
    if (process.env.DEBUG) console.error(err.stack);
    process.exit(1);
  }
})();
