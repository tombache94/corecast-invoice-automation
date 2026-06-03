#!/usr/bin/env node
/**
 * Smoke-test the Xero bill pipeline.
 *
 * Two modes:
 *   --dry-run   (default) — categorise + offline simulation, no API calls
 *   --live                — connect to Xero and create a real DRAFT bill
 *
 * Usage:
 *   node scripts/test-xero-pipeline.js
 *   node scripts/test-xero-pipeline.js --live
 *   node scripts/test-xero-pipeline.js --live --supplier "Boral Materials" --amount 1234.56
 */

const path = require('path');
const { config, assertXeroCreds } = require('../src/config');
const { categoriseInvoice, ACCOUNT_NAMES } = require('../src/xero/categorise');
const { pushInvoiceToXero } = require('../src/xero');

function parseArgs(argv) {
  const args = { live: false, dryRun: true };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--live') { args.live = true; args.dryRun = false; }
    else if (a === '--dry-run') { args.live = false; args.dryRun = true; }
    else if (a === '--supplier') args.supplier = argv[++i];
    else if (a === '--amount') args.amount = parseFloat(argv[++i]);
    else if (a === '--subject') args.subject = argv[++i];
    else if (a === '--help' || a === '-h') args.help = true;
  }
  return args;
}

function fixtures(args) {
  const today = new Date().toISOString().slice(0, 10);
  return [
    {
      id: 'test_001_materials',
      sender: args.supplier || 'Test Supplier Pty Ltd',
      senderEmail: 'accounts@testsupplier.example',
      subject: args.subject || 'Invoice #TEST-001 — Materials',
      amount: args.amount ?? 1000,
      date: today,
      weekEnding: today,
      hoursClaimed: null,
    },
    {
      id: 'test_002_pkf',
      sender: 'PKF Perth',
      subject: 'Invoice #PKF-2026-04 — accounting services',
      amount: 1320,
      date: today,
      weekEnding: today,
      hoursClaimed: null,
    },
    {
      id: 'test_003_hire',
      sender: 'Total Plant Hire',
      subject: 'Equipment hire — week ending',
      amount: 880,
      date: today,
      weekEnding: today,
      hoursClaimed: null,
    },
    {
      id: 'test_004_freight',
      sender: 'Centurion Transport',
      subject: 'Freight delivery — Perth metro',
      amount: 450,
      date: today,
      weekEnding: today,
      hoursClaimed: null,
    },
    {
      id: 'test_005_labour',
      sender: 'Aaron Norris',
      subject: 'Invoice — week ending',
      amount: 1923,
      date: today,
      weekEnding: today,
      hoursClaimed: 19.23, // labour — should be skipped
    },
    {
      id: 'test_006_suspicious',
      sender: 'Suspicious Big Bill Co',
      subject: 'Invoice — large amount',
      amount: 99999,
      date: today,
      weekEnding: today,
      hoursClaimed: null,
    },
  ];
}

function printCategorisation(invoices) {
  console.log('--- Categorisation ---');
  for (const inv of invoices) {
    const cat = categoriseInvoice(inv);
    if (cat.type === 'LABOUR') {
      console.log(`  LABOUR  ${inv.sender}`);
    } else {
      console.log(
        `  BILL    ${inv.sender.padEnd(28)} → ${cat.accountCode} ${ACCOUNT_NAMES[cat.accountCode]}  (${cat.reason})`,
      );
    }
  }
  console.log('');
}

async function runDryRun(invoices) {
  console.log('Mode: DRY RUN (no API calls)\n');
  printCategorisation(invoices);

  // Stub client that simulates contacts/invoices look-ups without calling Xero.
  const stubClient = {
    async findOrCreateContact({ name }) {
      return { contactID: `stub-${name.replace(/\s+/g, '-').toLowerCase()}`, name };
    },
    async findExistingBill() { return null; },
    async createDraftBill() {
      throw new Error('createDraftBill should not be called in dry-run mode');
    },
  };

  for (const inv of invoices) {
    const result = await pushInvoiceToXero(inv, { client: stubClient, dryRun: true });
    console.log(`${inv.sender}:`);
    console.log(`  ${JSON.stringify(result, null, 2).split('\n').join('\n  ')}\n`);
  }
}

async function runLive(invoices) {
  console.log('Mode: LIVE — bills will be created as DRAFT in Xero\n');
  assertXeroCreds();

  const { XeroBillsClient } = require('../src/xero/xeroClient');
  const client = new XeroBillsClient({
    clientId: config.xero.clientId,
    clientSecret: config.xero.clientSecret,
    tenantId: config.xero.tenantId,
    refreshToken: config.xero.refreshToken,
    accessToken: config.xero.accessToken,
    tokenCacheFile: config.xero.tokenCacheFile,
  });

  console.log(`Connecting to Xero tenant ${config.xero.tenantId}...`);
  await client.init();
  console.log('Connected.\n');

  printCategorisation(invoices);

  for (const inv of invoices) {
    console.log(`→ ${inv.sender} ($${inv.amount})`);
    try {
      const result = await pushInvoiceToXero(inv, { client });
      if (result.created) {
        console.log(`  ✅ Created DRAFT bill ${result.xeroNumber || result.xeroId} under ${result.accountName}`);
      } else if (result.skipped) {
        console.log(`  ⏸️  Skipped: ${result.reason}${result.detail ? ' — ' + result.detail : ''}`);
      } else {
        console.log(`  ?  ${JSON.stringify(result)}`);
      }
    } catch (err) {
      console.log(`  ❌ ERROR: ${err.message}`);
      if (process.env.DEBUG) console.log(err.stack);
    }
    console.log('');
  }
}

function printHelp() {
  console.log(`Usage: node scripts/test-xero-pipeline.js [options]

Options:
  --dry-run                 Categorise and simulate only (default)
  --live                    Actually call Xero and create DRAFT bills
  --supplier <name>         Override first fixture supplier name
  --amount <number>         Override first fixture amount
  --subject <string>        Override first fixture subject
  --help, -h                Show this help
`);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) { printHelp(); return; }

  const invoices = fixtures(args);

  if (args.live) await runLive(invoices);
  else await runDryRun(invoices);
}

if (require.main === module) {
  main().catch((err) => {
    console.error('Fatal:', err.message);
    if (process.env.DEBUG) console.error(err.stack);
    process.exit(1);
  });
}
