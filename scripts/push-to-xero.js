#!/usr/bin/env node
require('dotenv').config();
const { config, assertXeroCreds } = require('../src/config');
const { XeroBillsClient } = require('../src/xero/xeroClient');
const { pushManyToXero } = require('../src/xero');
const { InvoiceStore } = require('../src/storage/invoices');

assertXeroCreds();

const client = new XeroBillsClient({
  clientId: config.xero.clientId,
  clientSecret: config.xero.clientSecret,
  tenantId: config.xero.tenantId,
  refreshToken: config.xero.refreshToken,
  accessToken: config.xero.accessToken,
  tokenCacheFile: config.xero.tokenCacheFile,
});

async function run() {
  await client.init();
  console.log('Connected to Xero.\n');

  const { invoices: all } = JSON.parse(require('fs').readFileSync(config.storage.invoicesFile, 'utf8'));
  const candidates = all.filter(function(inv) {
    return inv.amount != null && inv.amount <= 50000;
  });

  console.log('Pushing ' + candidates.length + ' invoices to Xero...\n');

  var { results, counts } = await pushManyToXero(candidates, {
    client: client,
    dryRun: false,
    onProgress: function(invoice, result) {
      var line = '  ' + invoice.sender + ' | $' + invoice.amount + ' | ' + invoice.date + ' → ';
      if (result.created) {
        line += 'CREATED ' + (result.xeroNumber || result.xeroId) + ' (' + result.accountName + ')';
      } else if (result.reason === 'duplicate') {
        line += 'DUPLICATE (already in Xero: ' + result.xeroNumber + ')';
      } else if (result.error) {
        line += 'ERROR: ' + result.error;
      } else {
        line += 'SKIPPED: ' + result.reason;
      }
      console.log(line);
    },
  });

  console.log('\n=== SUMMARY ===');
  console.log('  Created:    ' + counts.created);
  console.log('  Duplicates: ' + counts.duplicates);
  console.log('  Flagged:    ' + counts.flagged);
  console.log('  Errors:     ' + counts.errors);
  console.log('  Total:      ' + counts.total);
}

run().catch(function(e) { console.error('Fatal:', e.message); if (process.env.DEBUG) console.error(e.stack); process.exit(1); });
