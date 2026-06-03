/**
 * Daily dashboard refresh — pulls Monday + Xero + SharePoint + manual config,
 * composes a single DASHBOARD_DATA blob, and splices it into the artifact's
 * index.html. Designed to run via Windows Task Scheduler each morning.
 *
 * Usage:
 *   node scripts/refresh-dashboard.js                      # live refresh
 *   node scripts/refresh-dashboard.js --dry-run            # don't touch HTML
 *   node scripts/refresh-dashboard.js --html "C:\\path\\to\\index.html"
 */

require('dotenv').config();
const path = require('path');
const fs = require('fs');
const { compose } = require('../src/dashboard/compose');
const { renderToHtml } = require('../src/dashboard/render');
const { syncJobLinks } = require('../src/sharepoint/jobLinkSync');
const { deployToCloudflare } = require('../src/dashboard/cloudflareDeploy');
const { notifyOnRefresh } = require('../src/dashboard/notifyOnRefresh');
const { fileAttachments } = require('../src/filing/fileAttachments');

const DEFAULT_HTML =
  process.env.DASHBOARD_HTML_PATH ||
  'C:\\Users\\tomba\\Documents\\Claude\\Artifacts\\corecast-financial-dashboard\\index.html';

function parseArgs(argv) {
  const args = { dryRun: false, htmlPath: DEFAULT_HTML };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--dry-run') args.dryRun = true;
    else if (a === '--html') args.htmlPath = argv[++i];
    else if (a.startsWith('--html=')) args.htmlPath = a.slice(7);
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));

  console.log('🏗️  CoreCast — Dashboard Refresh');
  console.log(`   Mode : ${args.dryRun ? '🔍 DRY RUN (no HTML write)' : '✏️  LIVE'}`);
  console.log(`   HTML : ${args.htmlPath}\n`);

  // Sync SharePoint folder links to Monday BEFORE composing, so the
  // dashboard immediately reflects any newly-written links. Best effort —
  // if this fails the dashboard refresh still proceeds.
  console.log('Syncing SharePoint → Monday links…');
  try {
    await syncJobLinks({ dryRun: args.dryRun });
  } catch (err) {
    console.log(`   ⚠️  link sync failed (continuing): ${err.message}`);
  }
  console.log();

  // File any new email attachments to SharePoint job folders. Returns a
  // digest of {tier, label, attachmentName, from, subject, webUrl} entries
  // which the post-refresh notification email will include. Best-effort —
  // a filing failure does not block the dashboard refresh. Pass
  // skipDigestEmail so fileAttachments doesn't send its own separate
  // digest email (we want a single consolidated refresh email).
  console.log('Filing email attachments…');
  let filingDigest = [];
  let filingStats = null;
  try {
    const result = await fileAttachments({ dryRun: args.dryRun, skipDigestEmail: true });
    filingDigest = result.digest || [];
    filingStats = result.stats || null;
  } catch (err) {
    console.log(`   ⚠️  filing failed (continuing): ${err.message}`);
  }
  console.log();

  console.log('Fetching sources in parallel…');
  const t0 = Date.now();
  const data = await compose();
  console.log(`Done in ${(Date.now() - t0)} ms\n`);

  console.log('Source health:');
  for (const [name, s] of Object.entries(data.meta.sources)) {
    const tick = s.ok ? '✅' : '⚠️ ';
    const detail = s.message ? ` — ${s.message}` : '';
    console.log(`   ${tick} ${name.padEnd(18)} count=${s.count}${detail}`);
  }
  console.log();
  console.log('Summary:');
  console.log(`   Active jobs    : ${data.summary.activeJobs} (${data.summary.activeJobsBreakdown})`);
  console.log(`   Total overdue  : $${data.summary.totalOverdue.toLocaleString('en-AU')} across ${data.summary.overdueCount} invoice(s)`);
  console.log(`   Outstanding    : $${data.summary.totalOutstanding.toLocaleString('en-AU')}`);
  console.log(`   Labour hours   : ~${data.summary.totalLabourHours} hrs`);
  console.log(`   Auto flags     : ${data.flags.length}`);
  console.log();

  if (args.dryRun) {
    const out = path.resolve(__dirname, '../data/dashboard-data.dryrun.json');
    fs.mkdirSync(path.dirname(out), { recursive: true });
    fs.writeFileSync(out, JSON.stringify(data, null, 2));
    console.log(`🔍 Dry run — composed data written to ${out}`);
    console.log('   Inspect it, then re-run without --dry-run to update the dashboard.');
    return;
  }

  renderToHtml(args.htmlPath, data);
  console.log(`✅ Dashboard updated: ${args.htmlPath}`);

  // Best-effort: push the updated HTML to Cloudflare Pages so the public
  // share URL always reflects today's data. A failure here is logged but
  // does not fail the refresh — the local dashboard is the source of truth.
  console.log('\nPublishing to Cloudflare Pages…');
  const cf = await deployToCloudflare({ htmlPath: args.htmlPath });
  if (cf.ok) {
    console.log(`✅ Cloudflare Pages: ${cf.url}`);
  } else {
    console.log(`⚠️  Cloudflare deploy skipped/failed: ${cf.message}`);
  }

  // Best-effort: email a change summary. Never blocks the refresh.
  console.log('\nSending refresh notification email…');
  const notify = await notifyOnRefresh(data, { filingDigest, filingStats });
  if (notify.ok) {
    console.log(`✅ Notification sent: ${notify.message}`);
    console.log(`   Subject: ${notify.subject}`);
  } else {
    console.log(`⚠️  Notification skipped: ${notify.message}`);
  }
}

main().catch((err) => {
  console.error('❌ Refresh failed:', err.message);
  if (process.env.DEBUG) console.error(err.stack);
  process.exit(1);
});
