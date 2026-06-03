/**
 * Probe whether the SharePoint Graph API exposes folder colours for items
 * in the Active Quotes folder. Runs three queries:
 *   1. List children with no $select (returns ALL driveItem properties).
 *   2. Get the first folder's full driveItem.
 *   3. Get the first folder's listItem with $expand=fields (where SP custom
 *      fields like _ColorTag typically live).
 *
 * Read-only. Just dumps field names and any colour-shaped values it finds.
 */

require('isomorphic-fetch');
require('dotenv').config();

const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

const DRIVE_ID =
  process.env.SHAREPOINT_DRIVE_ID ||
  'b!4B_dBjoPKEeuMh4XWxgxCQcyBxLlmwFLgtOWOYvkPTQGCZjiE9CmRoLS7JMod59t';
const QUOTES_PATH = 'CORECAST/📋 QUOTES/Active Quotes';

function looksColourish(key, value) {
  const k = String(key).toLowerCase();
  if (k.includes('color') || k.includes('colour') || k.includes('tag') || k.includes('hue')) return true;
  if (typeof value === 'string') {
    const v = value.toLowerCase();
    if (/^(green|red|blue|amber|yellow|orange|pink|purple|teal|gray|grey)$/.test(v)) return true;
    if (/^#[0-9a-f]{3,8}$/.test(v)) return true;
  }
  return false;
}

function dumpInteresting(obj, prefix) {
  if (!obj || typeof obj !== 'object') return;
  for (const [k, v] of Object.entries(obj)) {
    if (looksColourish(k, v)) {
      console.log(`   🎨 ${prefix}${k} = ${JSON.stringify(v)}`);
    }
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      dumpInteresting(v, `${prefix}${k}.`);
    }
  }
}

async function main() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    console.error('Missing MS_* credentials in .env'); process.exit(1);
  }

  const msal = new ConfidentialClientApplication({
    auth: { clientId, clientSecret, authority: `https://login.microsoftonline.com/${tenantId}` },
  });
  const graph = Client.init({
    authProvider: async (done) => {
      try {
        const r = await msal.acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        done(null, r.accessToken);
      } catch (e) { done(e, null); }
    },
  });

  console.log('Drive   :', DRIVE_ID);
  console.log('Folder  :', QUOTES_PATH);
  console.log('');

  // ── Query 1: list children with no $select (all default properties) ──
  console.log('── 1) /children (no $select) ──');
  let children;
  try {
    const res = await graph.api(`/drives/${DRIVE_ID}/root:/${QUOTES_PATH}:/children`).top(20).get();
    children = res.value || [];
    console.log(`   ${children.length} children returned`);
    if (children[0]) {
      console.log(`\n   Top-level keys on first item ("${children[0].name}"):`);
      console.log('   ' + Object.keys(children[0]).join(', '));
      console.log('');
      console.log('   Anything colour-shaped:');
      dumpInteresting(children[0], '');
    }
  } catch (e) {
    console.log(`   ❌ failed: ${e.statusCode} ${e.code} ${e.message}`); process.exit(2);
  }

  if (!children.length) { console.log('No children — nothing more to probe.'); return; }
  const first = children[0];

  // ── Query 2: the first folder's driveItem standalone ──
  console.log('\n── 2) /items/{id} (full driveItem) ──');
  try {
    const item = await graph.api(`/drives/${DRIVE_ID}/items/${first.id}`).get();
    console.log(`   keys: ${Object.keys(item).join(', ')}`);
    console.log('   colour-shaped:');
    dumpInteresting(item, '');
  } catch (e) {
    console.log(`   ❌ failed: ${e.statusCode} ${e.code} ${e.message}`);
  }

  // ── Query 3: listItem with all fields ──
  console.log('\n── 3) /items/{id}/listItem?$expand=fields ──');
  try {
    const li = await graph.api(`/drives/${DRIVE_ID}/items/${first.id}/listItem`).expand('fields').get();
    const fields = li.fields || {};
    console.log(`   listItem.fields keys (${Object.keys(fields).length}):`);
    console.log('   ' + Object.keys(fields).join(', '));
    console.log('');
    console.log('   colour-shaped fields:');
    dumpInteresting(fields, 'fields.');
  } catch (e) {
    console.log(`   ❌ failed: ${e.statusCode} ${e.code} ${e.message}`);
  }

  // ── Query 4: try the documented bundle.color path (some OneDrive shape) ──
  console.log('\n── 4) /items/{id}?$select=bundle,specialFolder ──');
  try {
    const item = await graph.api(`/drives/${DRIVE_ID}/items/${first.id}`).select('bundle,specialFolder,sharepointIds').get();
    console.log('   ' + JSON.stringify(item, null, 2));
  } catch (e) {
    console.log(`   ❌ failed: ${e.statusCode} ${e.code} ${e.message}`);
  }
}

main().catch((err) => { console.error('Fatal:', err.message); process.exit(99); });
