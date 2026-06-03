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

async function run() {
  await client.init();
  console.log('Connected to Xero.\n');
  const res = await client.sdk.accountingApi.getInvoices(
    client.tenantId,
    undefined,
    'Type=="ACCPAY"&&Date>=DateTime(2026,4,1)',
    'Date DESC',
  );
  const bills = res.body && res.body.invoices ? res.body.invoices : [];
  console.log('ACCPAY bills in Xero since April 2026 (' + bills.length + ' total):\n');
  bills.forEach(function(b) {
    var contact = (b.contact && b.contact.name) ? b.contact.name : '(unknown)';
    var status = b.status || '';
    var amt = b.total != null ? ('$' + b.total.toFixed(2)) : '$?';
    var ref = b.reference || '';
    var date = b.date ? (typeof b.date === 'string' ? b.date.slice(0, 10) : new Date(b.date).toISOString().slice(0, 10)) : '?';
    console.log('  ' + date + '  ' + pad(contact, 35) + '  ' + padLeft(amt, 12) + '  ' + pad(status, 12) + '  ' + ref);
  });
}

function pad(s, n) { s = String(s); while (s.length < n) s += ' '; return s; }
function padLeft(s, n) { s = String(s); while (s.length < n) s = ' ' + s; return s; }

run().catch(function(e) { console.error('Error:', e.message); process.exit(1); });
