/**
 * tag-tracking-categories.js
 *
 * Finds all Xero bills/transactions with job names in the reference or description
 * and tags them to the matching Xero tracking category option.
 *
 * Usage:
 *   node scripts/tag-tracking-categories.js           -- live run
 *   node scripts/tag-tracking-categories.js --dry-run -- preview only
 *
 * Requires: data/xero-token.json (run node scripts/xero-auth.js first)
 */

require('dotenv').config();
const https = require('https');
const fs = require('fs');
const path = require('path');

const TOKEN_FILE = path.resolve(__dirname, '../data/xero-token.json');
const TENANT_ID = process.env.XERO_TENANT_ID || '024c8088-8095-482b-906e-f72d9b8acaee';
const CLIENT_ID = process.env.XERO_CLIENT_ID;
const CLIENT_SECRET = process.env.XERO_CLIENT_SECRET;
const DRY_RUN = process.argv.includes('--dry-run');

// Job keyword → tracking category option name mapping
// Keys are lowercase substrings to search in reference/description/invoice number
const JOB_KEYWORDS = [
  { keywords: ['baldivis', 'parkland heights'], job: 'Baldivis' },
  { keywords: ['eneabba'],                      job: 'Eneabba' },
  { keywords: ['osbourne park', 'osborne park', 'colray'], job: 'Osbourne Park' },
  { keywords: ['regans ford', 'regan'],         job: 'Regans Ford' },
  { keywords: ['bunnings lake', 'bunning lake'], job: 'Bunning Lake' },
  { keywords: ['pennant', 'neerabup'],          job: 'Pennant' },
  { keywords: ['montreal'],                     job: 'Montreal' },
  { keywords: ['power on'],                     job: 'Power On' },
];

// ── Helpers ────────────────────────────────────────────────────────────────────

function loadTokens() {
  if (!fs.existsSync(TOKEN_FILE)) {
    console.error(`\n❌ No token file found at ${TOKEN_FILE}`);
    console.error('Run: node scripts/xero-auth.js\n');
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
}

function saveTokens(tokens) {
  fs.writeFileSync(TOKEN_FILE, JSON.stringify({ ...tokens, saved_at: new Date().toISOString() }, null, 2));
}

function xeroRequest({ method = 'GET', path: apiPath, body, accessToken }) {
  return new Promise((resolve, reject) => {
    const bodyStr = body ? JSON.stringify(body) : null;
    const req = https.request({
      hostname: 'api.xero.com',
      path: apiPath,
      method,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'xero-tenant-id': TENANT_ID,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        ...(bodyStr ? { 'Content-Length': Buffer.byteLength(bodyStr) } : {}),
      },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

async function refreshAccessToken(tokens) {
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: tokens.refresh_token,
  }).toString();

  const credentials = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'identity.xero.com',
      path: '/connect/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': `Basic ${credentials}`,
        'Content-Length': Buffer.byteLength(body),
      },
    }, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve(JSON.parse(data)));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function detectJob(invoice) {
  const text = [
    invoice.Reference || '',
    invoice.InvoiceNumber || '',
    (invoice.LineItems || []).map(l => l.Description || '').join(' '),
  ].join(' ').toLowerCase();

  for (const { keywords, job } of JOB_KEYWORDS) {
    if (keywords.some(k => text.includes(k))) return job;
  }
  return null;
}

// ── Main ───────────────────────────────────────────────────────────────────────

async function main() {
  console.log(`\n🏗️  CoreCast — Xero Tracking Category Tagger`);
  console.log(`   Mode: ${DRY_RUN ? '🔍 DRY RUN (no changes)' : '✏️  LIVE'}\n`);

  // 1. Load + refresh token
  let tokens = loadTokens();
  console.log('🔑 Refreshing Xero access token...');
  const newTokens = await refreshAccessToken(tokens);
  if (newTokens.error) {
    console.error('❌ Token refresh failed:', newTokens.error_description || newTokens.error);
    console.error('Run: node scripts/xero-auth.js');
    process.exit(1);
  }
  tokens = { ...tokens, ...newTokens };
  saveTokens(tokens);
  console.log('✅ Token refreshed\n');

  const AT = tokens.access_token;

  // 2. Get tracking categories
  console.log('📋 Fetching tracking categories...');
  const tcRes = await xeroRequest({ path: '/api.xro/2.0/TrackingCategories', accessToken: AT });
  if (tcRes.status !== 200) {
    console.error('❌ Failed to fetch tracking categories:', tcRes.body);
    process.exit(1);
  }

  const categories = tcRes.body.TrackingCategories || [];
  console.log(`   Found ${categories.length} tracking categories:`);
  categories.forEach(c => {
    console.log(`   - ${c.Name} (${c.Options?.length || 0} options)`);
    (c.Options || []).forEach(o => console.log(`     • ${o.Name} [${o.TrackingOptionID}]`));
  });

  // Find the "Job" tracking category
  const jobCategory = categories.find(c =>
    c.Name.toLowerCase() === 'job' || c.Name.toLowerCase() === 'jobs'
  );
  if (!jobCategory) {
    console.error('\n❌ No "Job" tracking category found in Xero.');
    console.error('Available categories:', categories.map(c => c.Name).join(', '));
    console.error('Create a tracking category called "Job" in Xero first.\n');
    process.exit(1);
  }

  console.log(`\n✅ Using tracking category: "${jobCategory.Name}" [${jobCategory.TrackingCategoryID}]`);

  // Build option name → ID map
  const optionMap = {};
  for (const opt of jobCategory.Options || []) {
    optionMap[opt.Name.toLowerCase()] = { id: opt.TrackingOptionID, name: opt.Name };
  }
  console.log(`   Options: ${Object.values(optionMap).map(o => o.name).join(', ')}\n`);

  // 3. Fetch all ACCPAY bills (draft + authorised + paid)
  console.log('📥 Fetching all purchase bills from Xero...');
  let allBills = [];
  for (const status of ['DRAFT', 'AUTHORISED', 'PAID']) {
    const res = await xeroRequest({
      path: `/api.xro/2.0/Invoices?where=Type%3D%3D%22ACCPAY%22%26%26Status%3D%3D%22${status}%22&page=1`,
      accessToken: AT,
    });
    if (res.status === 200) {
      const bills = res.body.Invoices || [];
      allBills = allBills.concat(bills);
    }
  }
  console.log(`   Fetched ${allBills.length} bills total\n`);

  // 4. Match bills to jobs and tag
  let tagged = 0, skipped = 0, alreadyTagged = 0, notMatched = 0;

  for (const bill of allBills) {
    const jobName = detectJob(bill);
    if (!jobName) { notMatched++; continue; }

    const option = optionMap[jobName.toLowerCase()];
    if (!option) {
      console.log(`⚠️  Job "${jobName}" not found as a tracking option — skipping`);
      skipped++;
      continue;
    }

    // Check if already tagged. Xero's GET response on Invoices returns
    // Tracking entries with name-based fields (`Name` for the category,
    // `Option` for the option) rather than the ID-based fields it accepts
    // on POST. Match on either form so we don't re-tag bills every run.
    const alreadyHasTag = (bill.LineItems || []).some(li =>
      (li.Tracking || []).some(t => {
        const idMatch =
          t.TrackingCategoryID === jobCategory.TrackingCategoryID &&
          t.TrackingOptionID === option.id;
        const nameMatch =
          (t.Name || '').toLowerCase() === (jobCategory.Name || '').toLowerCase() &&
          (t.Option || '').toLowerCase() === (option.name || '').toLowerCase();
        return idMatch || nameMatch;
      })
    );
    if (alreadyHasTag) {
      console.log(`✓  Already tagged: ${bill.InvoiceNumber || bill.InvoiceID} → ${jobName}`);
      alreadyTagged++;
      continue;
    }

    console.log(`${DRY_RUN ? '🔍' : '🏷️ '} ${DRY_RUN ? '[DRY RUN] Would tag' : 'Tagging'}: ${bill.Contact?.Name || '?'} | ${bill.InvoiceNumber || bill.Reference || bill.InvoiceID} | $${bill.Total} → ${jobName}`);

    if (!DRY_RUN) {
      // Build updated line items with tracking
      const updatedLineItems = (bill.LineItems || []).map(li => ({
        LineItemID: li.LineItemID,
        Description: li.Description,
        Quantity: li.Quantity,
        UnitAmount: li.UnitAmount,
        AccountCode: li.AccountCode,
        TaxType: li.TaxType,
        Tracking: [
          { TrackingCategoryID: jobCategory.TrackingCategoryID, TrackingOptionID: option.id },
        ],
      }));

      const updateRes = await xeroRequest({
        method: 'POST',
        path: `/api.xro/2.0/Invoices/${bill.InvoiceID}`,
        accessToken: AT,
        body: {
          Invoices: [{
            InvoiceID: bill.InvoiceID,
            LineItems: updatedLineItems,
          }],
        },
      });

      if (updateRes.status === 200) {
        console.log(`   ✅ Tagged successfully`);
        tagged++;
      } else {
        console.log(`   ❌ Failed: ${JSON.stringify(updateRes.body?.Elements?.[0]?.ValidationErrors || updateRes.body)}`);
        skipped++;
      }

      // Small delay to avoid rate limiting
      await new Promise(r => setTimeout(r, 200));
    } else {
      tagged++;
    }
  }

  // 5. Summary
  console.log(`\n${'─'.repeat(50)}`);
  console.log(`📊 Summary${DRY_RUN ? ' (DRY RUN)' : ''}:`);
  console.log(`   ${DRY_RUN ? 'Would tag' : 'Tagged'}:      ${tagged}`);
  console.log(`   Already tagged: ${alreadyTagged}`);
  console.log(`   No match:       ${notMatched}`);
  console.log(`   Skipped/errors: ${skipped}`);
  console.log(`   Total bills:    ${allBills.length}`);
  if (DRY_RUN) {
    console.log(`\n   Run without --dry-run to apply changes.`);
  }
  console.log();
}

main().catch(err => {
  console.error('\n❌ Fatal error:', err.message);
  process.exit(1);
});
