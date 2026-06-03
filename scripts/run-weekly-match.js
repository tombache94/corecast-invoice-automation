#!/usr/bin/env node
/**
 * Weekly job: match invoices for the latest Connecteam timesheet export
 * and email a validation report to Tom.
 *
 * Usage:
 *   node scripts/run-weekly-match.js
 *   node scripts/run-weekly-match.js --timesheet ./timesheets/timeclock-timesheet_overview_2026-04-27_2026-05-03.xlsx
 *   node scripts/run-weekly-match.js --dry-run        # parse + match, do NOT send email or update statuses
 *   node scripts/run-weekly-match.js --no-email       # match + update statuses, but skip the email
 */

const path = require('path');
const { config, assertMicrosoftCreds, assertXeroCreds } = require('../src/config');
const { OutlookClient } = require('../src/email/outlook');
const { InvoiceStore } = require('../src/storage/invoices');
const { matchWeek } = require('../src/matcher/matchWeek');
const { sendReport, buildSubject, buildPlainText } = require('../src/matcher/sendReport');
const { XeroBillsClient } = require('../src/xero/xeroClient');

function parseArgs(argv) {
  const args = { dryRun: false, sendEmail: true, xero: 'auto', xeroDryRun: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--timesheet' || a === '-t') {
      args.timesheetPath = argv[++i];
    } else if (a === '--dry-run') {
      args.dryRun = true;
      args.sendEmail = false;
      args.xeroDryRun = true;
    } else if (a === '--no-email') {
      args.sendEmail = false;
    } else if (a === '--no-xero') {
      args.xero = 'off';
    } else if (a === '--xero-dry-run') {
      args.xeroDryRun = true;
    } else if (a === '--help' || a === '-h') {
      args.help = true;
    }
  }
  return args;
}

function printHelp() {
  console.log(`Usage: node scripts/run-weekly-match.js [options]

Options:
  --timesheet, -t <path>   Path to Connecteam XLSX export (defaults to latest in watch folder)
  --dry-run                Parse and match only — don't update statuses, send email, or create bills
  --no-email               Update invoice statuses but skip the report email
  --no-xero                Skip the Xero bill pipeline entirely
  --xero-dry-run           Categorise + duplicate-check bills, but don't create them in Xero
  --help, -h               Show this help
`);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return;
  }

  const store = new InvoiceStore(config.storage.invoicesFile);

  console.log('CoreCast — Weekly Timesheet Match');
  console.log('==================================');
  console.log(`Watch folder: ${path.resolve(config.timesheets.watchFolder)}`);
  if (args.timesheetPath) console.log(`Override:     ${args.timesheetPath}`);
  if (args.dryRun) console.log(`Mode:         DRY RUN (no writes, no email, no bills)`);
  else if (!args.sendEmail) console.log(`Mode:         No email (statuses will be updated)`);
  console.log('');

  let xeroClient = null;
  if (args.xero !== 'off' && !args.dryRun) {
    try {
      assertXeroCreds();
      xeroClient = new XeroBillsClient({
        clientId: config.xero.clientId,
        clientSecret: config.xero.clientSecret,
        tenantId: config.xero.tenantId,
        refreshToken: config.xero.refreshToken,
        accessToken: config.xero.accessToken,
        tokenCacheFile: config.xero.tokenCacheFile,
      });
      await xeroClient.init();
      console.log(`Xero:         connected (tenant ${config.xero.tenantId})${args.xeroDryRun ? ' [DRY-RUN]' : ''}`);
    } catch (err) {
      console.log(`Xero:         skipped — ${err.message}`);
      xeroClient = null;
    }
  } else if (args.xero === 'off') {
    console.log('Xero:         disabled by --no-xero');
  }
  console.log('');

  const result = await matchWeek({
    timesheetPath: args.timesheetPath,
    watchFolder: config.timesheets.watchFolder,
    tolerance: config.matcher.hoursTolerance,
    store,
    updateStatuses: !args.dryRun,
    xeroClient,
    xeroDryRun: args.xeroDryRun,
  });

  console.log(`Week:         ${result.weekStart} to ${result.weekEnd}`);
  console.log(`Timesheet:    ${result.timesheetPath}`);
  console.log(`Tolerance:    ±${result.tolerance}h`);
  console.log(`Employees:    ${result.employees.length}`);
  console.log(`Invoices:     ${result.counts.total}`);
  console.log('');
  console.log(buildPlainText(result));
  console.log('');

  if (args.sendEmail) {
    assertMicrosoftCreds();
    const outlook = new OutlookClient(config.microsoft);
    const subject = buildSubject(result);
    console.log(`Sending report...`);
    console.log(`  From:    ${config.alerts.reportSender}`);
    console.log(`  To:      ${config.alerts.reportRecipient}`);
    console.log(`  Subject: ${subject}`);
    await sendReport(result, {
      outlook,
      sender: config.alerts.reportSender,
      recipient: config.alerts.reportRecipient,
    });
    console.log('Email sent.');
  } else {
    console.log('Email skipped.');
  }

  const failures = result.results.filter((r) => r.updateError);
  if (failures.length) {
    console.log('');
    console.log(`WARNING: ${failures.length} invoice status update(s) failed:`);
    for (const f of failures) {
      console.log(`  - ${f.invoiceId} (${f.sender}): ${f.updateError}`);
    }
    process.exitCode = 2;
  }
}

if (require.main === module) {
  main().catch((err) => {
    console.error('Fatal:', err.message);
    if (process.env.DEBUG) console.error(err.stack);
    process.exit(1);
  });
}
