/**
 * One-time setup: create a "Quote Value" Currency column on the SharePoint
 * document library that contains the Active Quotes folder. Idempotent — if
 * the column already exists, the script reports that and exits cleanly.
 *
 * Requires Sites.ReadWrite.All (Application) on the Azure AD app, with
 * admin consent. If only Sites.Read.All is granted, this will 403; in that
 * case follow the manual instructions printed at the bottom of the output.
 */

require('isomorphic-fetch');
require('dotenv').config();

const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

// IDs harvested from scripts/probe-folder-colour.js — the SharePoint list
// that backs the CORECAST drive on the CoreCast site.
const SITE_ID = '06dd1fe0-0f3a-4728-ae32-1e175b183109';
const LIST_ID = 'e2980906-d013-46a6-82d2-ec9328779f6d';

const COLUMN = {
  name: 'QuoteValue',           // internal — used as listItem.fields.QuoteValue
  displayName: 'Quote Value',   // shown in the SharePoint UI
  currency: { locale: 'en-AU' },
};

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

  console.log('Site :', SITE_ID);
  console.log('List :', LIST_ID);
  console.log('');

  // Idempotency: bail early if a column with this internal name already exists.
  console.log('Checking for existing column…');
  let existing;
  try {
    const cols = await graph.api(`/sites/${SITE_ID}/lists/${LIST_ID}/columns`).top(200).get();
    existing = (cols.value || []).find((c) => c.name === COLUMN.name || c.displayName === COLUMN.displayName);
  } catch (err) {
    console.log(`  ❌ list columns failed: ${err.statusCode} ${err.code} ${err.message}`);
    if (err.statusCode === 403) printManualSteps();
    process.exit(2);
  }

  if (existing) {
    console.log(`  ✅ Column already exists — name="${existing.name}", displayName="${existing.displayName}", id=${existing.id}`);
    console.log('Nothing to do.');
    return;
  }

  console.log('Creating column…');
  try {
    const created = await graph.api(`/sites/${SITE_ID}/lists/${LIST_ID}/columns`).post(COLUMN);
    console.log(`  ✅ Created. id=${created.id}, name=${created.name}, displayName=${created.displayName}`);
    console.log('');
    console.log('Verify by setting a value on a folder in SharePoint:');
    console.log('  Active Quotes → right-click a folder → Properties → Quote Value');
    console.log('Then run: node scripts/refresh-dashboard.js');
  } catch (err) {
    console.log(`  ❌ create failed: ${err.statusCode} ${err.code} ${err.message}`);
    if (err.statusCode === 403) printManualSteps();
    process.exit(3);
  }
}

function printManualSteps() {
  console.log('');
  console.log('────────────────────────────────────────────────────────');
  console.log('Manual fallback (do this in the SharePoint UI):');
  console.log('  1. Open https://netorgft19809673.sharepoint.com/sites/CoreCast');
  console.log('  2. Navigate into the document library that contains "📋 QUOTES/Active Quotes"');
  console.log('  3. Click "+ Add column" in the column header row');
  console.log('  4. Choose "Currency"');
  console.log('  5. Name: "Quote Value"   Format: AU$   Decimals: 0');
  console.log('  6. Save. The column will appear library-wide; only Active Quotes will use it.');
  console.log('');
  console.log('Then re-run sharepoint.js consumers — no code change needed.');
}

main().catch((err) => { console.error('Fatal:', err.message); process.exit(99); });
