#!/usr/bin/env node
/**
 * Finds DRAFT bills created by our pipeline (reference starts with "Week ending")
 * that are duplicates of pre-existing bills (same contact, same amount, within 30 days).
 * Deletes the pipeline-created drafts.
 */
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

function isoDate(d) {
  return (typeof d === 'string' ? d : new Date(d).toISOString()).slice(0, 10);
}

function daysDiff(a, b) {
  return Math.abs((new Date(a) - new Date(b)) / 86400000);
}

async function run() {
  await client.init();
  console.log('Fetching all ACCPAY bills since April 2026...\n');

  var res = await client.sdk.accountingApi.getInvoices(
    client.tenantId,
    undefined,
    'Type=="ACCPAY"&&Date>=DateTime(2026,4,1)',
    'Date DESC',
    undefined, undefined, undefined,
    ['DRAFT', 'SUBMITTED', 'AUTHORISED', 'PAID']
  );
  var bills = (res.body && res.body.invoices) ? res.body.invoices : [];
  console.log('Found ' + bills.length + ' active bills.\n');

  // Our bills have reference starting with "Week ending"
  var ourDrafts = bills.filter(function(b) {
    return b.status === 'DRAFT' && b.reference && b.reference.startsWith('Week ending');
  });

  // Other bills (manually entered — no "Week ending" reference)
  var existing = bills.filter(function(b) {
    return !b.reference || !b.reference.startsWith('Week ending');
  });

  console.log('Our pipeline drafts: ' + ourDrafts.length);
  console.log('Pre-existing bills:  ' + existing.length + '\n');

  var toDelete = [];
  ourDrafts.forEach(function(draft) {
    var contactId = draft.contact && draft.contact.contactID;
    var amount = draft.total;
    var date = isoDate(draft.date);

    var match = existing.find(function(e) {
      var eContactId = e.contact && e.contact.contactID;
      var eAmount = e.total;
      var eDate = isoDate(e.date);
      return eContactId === contactId &&
             Math.abs(eAmount - amount) < 0.02 &&
             daysDiff(date, eDate) <= 30;
    });

    if (match) {
      console.log('DUPLICATE FOUND:');
      console.log('  Ours:     ' + date + ' ' + (draft.contact && draft.contact.name) + ' $' + amount + ' DRAFT ref="' + draft.reference + '" id=' + draft.invoiceID);
      console.log('  Existing: ' + isoDate(match.date) + ' ' + (match.contact && match.contact.name) + ' $' + match.total + ' ' + match.status + ' ref="' + (match.reference || '') + '"');
      toDelete.push(draft.invoiceID);
    }
  });

  if (!toDelete.length) {
    console.log('\nNo duplicates found.');
    return;
  }

  console.log('\nDeleting ' + toDelete.length + ' duplicate draft(s)...');
  for (var i = 0; i < toDelete.length; i++) {
    var id = toDelete[i];
    try {
      await client.sdk.accountingApi.updateInvoice(
        client.tenantId, id,
        { invoices: [{ invoiceID: id, status: 'DELETED' }] }
      );
      console.log('  DELETED ' + id);
    } catch (err) {
      console.log('  ERROR deleting ' + id + ': ' + (err.message || JSON.stringify(err).slice(0, 100)));
    }
  }
  console.log('\nDone.');
}

run().catch(function(e) { console.error('Fatal:', e.message); process.exit(1); });
