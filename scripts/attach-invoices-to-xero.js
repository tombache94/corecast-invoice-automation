#!/usr/bin/env node
/* eslint-disable */
/**
 * For each invoice in the store that has an outlookMessageId and amount,
 * find the matching Xero bill, fetch the PDF attachment from Outlook,
 * and upload it to the Xero bill.
 *
 * Safe to re-run — Xero deduplicates attachments by filename.
 */
require('dotenv').config();
const fs = require('fs');
const os = require('os');
const path = require('path');
const { config, assertXeroCreds, assertMicrosoftCreds } = require('../src/config');
const { XeroBillsClient } = require('../src/xero/xeroClient');
const { OutlookClient } = require('../src/email/outlook');
const { categoriseInvoice } = require('../src/xero/categorise');
const { weekWindow } = require('../src/xero');
const { buildReference } = require('../src/xero');

assertXeroCreds();
assertMicrosoftCreds();

const xero = new XeroBillsClient({
  clientId: config.xero.clientId,
  clientSecret: config.xero.clientSecret,
  tenantId: config.xero.tenantId,
  refreshToken: config.xero.refreshToken,
  accessToken: config.xero.accessToken,
  tokenCacheFile: config.xero.tokenCacheFile,
});

const outlook = new OutlookClient(config.microsoft);

async function run() {
  await xero.init();
  console.log('Connected to Xero.\n');

  const { invoices } = JSON.parse(fs.readFileSync(config.storage.invoicesFile, 'utf8'));
  const candidates = invoices.filter(function(inv) {
    return inv.amount != null && inv.amount <= 50000 && inv.outlookMessageId && inv.mailbox;
  });

  console.log('Processing ' + candidates.length + ' invoices with Outlook message IDs...\n');

  var attached = 0, skipped = 0, errors = 0;

  for (var i = 0; i < candidates.length; i++) {
    if (i > 0) await new Promise(function(r) { setTimeout(r, 2000); }); // stay under Xero's 60 calls/min
    var inv = candidates[i];
    var label = inv.sender + ' $' + inv.amount + ' ' + inv.date;

    // Find the existing Xero bill
    var contact, bill;
    try {
      contact = await xero.findContactByName(inv.sender);
      if (!contact) { console.log('  SKIP (no Xero contact): ' + label); skipped++; continue; }

      var billDate = inv.date || inv.weekEnding;
      var window = weekWindow(inv.weekEnding || inv.date || billDate);

      // Try week window first, fall back to broader ±30 day search
      bill = await xero.findExistingBill({
        contactId: contact.contactID,
        dateFrom: window.from,
        dateTo: window.to,
        reference: buildReference(inv),
      });

      if (!bill) {
        // Fallback: search ±30 days around invoice date
        var d = new Date((inv.date || inv.weekEnding) + 'T00:00:00Z');
        var from = new Date(d); from.setUTCDate(d.getUTCDate() - 30);
        var to = new Date(d); to.setUTCDate(d.getUTCDate() + 30);
        bill = await xero.findExistingBill({
          contactId: contact.contactID,
          dateFrom: from.toISOString().slice(0, 10),
          dateTo: to.toISOString().slice(0, 10),
          reference: buildReference(inv),
        });
      }

      if (!bill) { console.log('  SKIP (bill not found in Xero): ' + label); skipped++; continue; }
    } catch (err) {
      console.log('  ERROR finding bill for ' + label + ': ' + err.message);
      errors++;
      continue;
    }

    // Fetch the Outlook message with attachments
    var pdfAttachment;
    try {
      var msg = await outlook.getMessage(inv.mailbox, inv.outlookMessageId, { includeAttachments: true });
      var atts = msg.attachments || [];
      pdfAttachment = atts.find(function(a) {
        return (a.contentType || '').toLowerCase().includes('pdf') || (a.name || '').toLowerCase().endsWith('.pdf');
      });
      if (!pdfAttachment) {
        // Try first image attachment (e.g. scanned invoice as JPG/PNG)
        pdfAttachment = atts.find(function(a) {
          var ct = (a.contentType || '').toLowerCase();
          return ct.includes('image') || ct.includes('jpeg') || ct.includes('png');
        });
      }
      if (!pdfAttachment) { console.log('  SKIP (no attachment in email): ' + label); skipped++; continue; }
    } catch (err) {
      console.log('  ERROR fetching Outlook message for ' + label + ': ' + err.message);
      errors++;
      continue;
    }

    // Upload to Xero
    try {
      var bytes = Buffer.from(pdfAttachment.contentBytes, 'base64');
      var fileName = pdfAttachment.name || (inv.sender.replace(/[^a-zA-Z0-9]/g, '_') + '_invoice.pdf');
      var mimeType = pdfAttachment.contentType || 'application/pdf';

      // Write to temp file then use createReadStream — xero-node requires a real fs stream
      var tmpFile = path.join(os.tmpdir(), 'xero_attach_' + Date.now() + '_' + fileName.replace(/[^a-zA-Z0-9._]/g, '_'));
      fs.writeFileSync(tmpFile, bytes);
      var stream = fs.createReadStream(tmpFile);

      try {
        await xero.sdk.accountingApi.createInvoiceAttachmentByFileName(
          xero.tenantId,   // xeroTenantId
          bill.invoiceID,  // invoiceID
          fileName,        // fileName
          stream,          // body
          false,           // includeOnline
          undefined,       // idempotencyKey
          { headers: { 'Content-Type': mimeType } }
        );
      } finally {
        try { fs.unlinkSync(tmpFile); } catch (_) {}
      }
      console.log('  ATTACHED ' + fileName + ' → bill ' + (bill.invoiceNumber || bill.invoiceID) + ' (' + label + ')');
      attached++;
    } catch (err) {
      var detail = err.message || (err.body && JSON.stringify(err.body)) || JSON.stringify(err);
      console.log('  ERROR uploading attachment for ' + label + ': ' + detail);
      if (err.response) console.log('    HTTP ' + err.response.statusCode + ': ' + JSON.stringify(err.response.body || '').slice(0, 200));
      errors++;
    }
  }

  console.log('\n=== SUMMARY ===');
  console.log('  Attached: ' + attached);
  console.log('  Skipped:  ' + skipped);
  console.log('  Errors:   ' + errors);
}

run().catch(function(e) { console.error('Fatal:', e.message); if (process.env.DEBUG) console.error(e.stack); process.exit(1); });
