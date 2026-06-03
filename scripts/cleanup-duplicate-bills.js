#!/usr/bin/env node
/**
 * Two-pass cleanup:
 *
 * Pass 1: Delete our pipeline DRAFT bills ("Week ending" reference) that duplicate
 * a pre-existing bill (different reference) for the same contact + amount within 7 days.
 *
 * Pass 2: For any contact+amount+weekEnding group that has multiple pipeline DRAFT
 * bills, keep one and delete the rest (caused by multiple push runs with the
 * date/weekEnding mismatch).
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

function mondayOf(iso) {
  var d = new Date(iso + 'T00:00:00Z');
  var dow = d.getUTCDay(); // 0=Sun
  var offset = dow === 0 ? -6 : 1 - dow;
  d.setUTCDate(d.getUTCDate() + offset);
  return d.toISOString().slice(0, 10);
}

function sameWeek(a, b) {
  return mondayOf(a) === mondayOf(b);
}

async function deleteBill(id, label) {
  try {
    await client.sdk.accountingApi.updateInvoice(
      client.tenantId, id,
      { invoices: [{ invoiceID: id, status: 'DELETED' }] }
    );
    console.log('  DELETED ' + id + ' (' + label + ')');
    return true;
  } catch (err) {
    console.log('  ERROR deleting ' + id + ': ' + (err.message || '?'));
    return false;
  }
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

  var ourDrafts = bills.filter(function(b) {
    return b.status === 'DRAFT' && b.reference && b.reference.startsWith('Week ending');
  });
  var preExisting = bills.filter(function(b) {
    return !b.reference || !b.reference.startsWith('Week ending');
  });

  var toDelete = new Set();

  // --- Pass 1: our bill duplicates a pre-existing one (same contact+amount, within 7 days) ---
  console.log('Pass 1: checking against pre-existing bills (7-day window)...');
  ourDrafts.forEach(function(draft) {
    if (toDelete.has(draft.invoiceID)) return;
    var contactId = draft.contact && draft.contact.contactID;
    var amount = draft.total;
    var date = isoDate(draft.date);

    var match = preExisting.find(function(e) {
      return (e.contact && e.contact.contactID) === contactId &&
             Math.abs((e.total || 0) - amount) < 0.02 &&
             sameWeek(date, isoDate(e.date));
    });

    if (match) {
      console.log('  DUP (pre-existing): ' + date + ' ' + (draft.contact && draft.contact.name) + ' $' + amount +
        ' | matches ' + isoDate(match.date) + ' ' + match.status + ' "' + (match.reference || '') + '"');
      toDelete.add(draft.invoiceID);
    }
  });

  // --- Pass 2: multiple pipeline drafts for same contact+amount+weekEnding ---
  console.log('\nPass 2: checking for multiple pipeline drafts per group...');
  var groups = {};
  ourDrafts.forEach(function(draft) {
    if (toDelete.has(draft.invoiceID)) return;
    var contactId = draft.contact && draft.contact.contactID;
    var key = contactId + '|' + draft.total + '|' + draft.reference;
    if (!groups[key]) groups[key] = [];
    groups[key].push(draft);
  });

  Object.keys(groups).forEach(function(key) {
    var group = groups[key];
    if (group.length <= 1) return;
    // Keep the first (oldest), delete the rest
    console.log('  GROUP: ' + group[0].contact.name + ' $' + group[0].total + ' "' + group[0].reference + '" — ' + group.length + ' bills, keeping 1');
    for (var i = 1; i < group.length; i++) {
      toDelete.add(group[i].invoiceID);
    }
  });

  if (!toDelete.size) {
    console.log('\nNo duplicates found. Xero is clean.');
    return;
  }

  console.log('\nDeleting ' + toDelete.size + ' duplicate draft(s)...');
  var ids = Array.from(toDelete);
  for (var i = 0; i < ids.length; i++) {
    var bill = ourDrafts.find(function(b) { return b.invoiceID === ids[i]; });
    var label = bill ? (isoDate(bill.date) + ' ' + (bill.contact && bill.contact.name) + ' $' + bill.total) : ids[i];
    await deleteBill(ids[i], label);
  }

  console.log('\nDone.');
}

run().catch(function(e) { console.error('Fatal:', e.message); process.exit(1); });
