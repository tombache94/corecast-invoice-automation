#!/usr/bin/env node
require('dotenv').config();
const { config, assertXeroCreds } = require('../src/config');
const { XeroBillsClient } = require('../src/xero/xeroClient');

assertXeroCreds();
const client = new XeroBillsClient({
  clientId: config.xero.clientId,
  clientSecret: config.xero.clientSecret,
  tenantId: config.xero.tenantId,
  refreshToken: config.xero.refreshToken,
  accessToken: config.xero.accessToken,
  tokenCacheFile: config.xero.tokenCacheFile,
});

// Bills created in run 2 that are duplicates of run 1
const TO_VOID = [
  '9718b566-ac8a-46c9-b46f-d1d82fe1db59', // Alani $3960 May 11 (run 2 duplicate)
  '58c70c62-9c0c-481d-8ef4-df368bb2e18d', // Alani $3960 May 4 (run 2 duplicate)
  '91c35c8c-1253-4839-9830-16e396e48710', // Sydney CBD Sonic $128.70 (run 2 duplicate)
  'b8839ad9-cbe7-42ef-afd7-9d55b90da314', // Alani $7469 May 25 (run 2 duplicate)
  '878116b1-7b16-4f85-8e71-6cbc048c6988', // Alani $7304 May 25 (run 2 duplicate)
  'a0372e73-bfc1-4369-a6a1-e4a1370c43cd', // Alani $3960 May 18 (run 2 duplicate)
];

async function run() {
  await client.init();
  console.log('Voiding ' + TO_VOID.length + ' duplicate bills...\n');
  for (var i = 0; i < TO_VOID.length; i++) {
    var id = TO_VOID[i];
    try {
      var res = await client.sdk.accountingApi.updateInvoice(
        client.tenantId,
        id,
        { invoices: [{ invoiceID: id, status: 'DELETED' }] }
      );
      var inv = res.body && res.body.invoices && res.body.invoices[0];
      console.log('  VOIDED ' + id + ' (' + (inv && inv.status) + ')');
    } catch (err) {
      console.log('  ERROR voiding ' + id + ': ' + err.message);
    }
  }
  console.log('\nDone.');
}

run().catch(function(e) { console.error('Fatal:', e.message); process.exit(1); });
