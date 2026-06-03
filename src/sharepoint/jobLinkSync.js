/**
 * For each active Monday job (stage = "In Progress" or "Not begun") whose
 * SharePoint Link column is empty, find a matching folder in the SharePoint
 * "ACTIVE JOBS" library and write its webUrl back to Monday.
 *
 * - Reads jobs via the existing fetchJobs() (already exposes sharePointUrl)
 * - Lists folders in config.filing.libraries["ACTIVE JOBS"]
 * - Match: case-insensitive, longest-substring-of-jobKey-in-folder-name wins
 * - Writes via Monday GraphQL change_column_value mutation
 * - Idempotent: never overwrites an existing link
 */

require('isomorphic-fetch');
const https = require('https');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

const { config } = require('../config');
const { fetchJobs } = require('../dashboard/sources/monday');

const JOB_TRACKER_BOARD_ID = 5027051539;
const LINK_COLUMN_ID = 'link_mm36d0ea';

function encodePath(p) {
  return p.split('/').map((seg) => encodeURIComponent(seg)).join('/');
}

async function listActiveJobsFolders() {
  const driveId = config.filing.driveId;
  const path = config.filing.libraries['ACTIVE JOBS'];
  if (!driveId || !path) {
    throw new Error('Missing config.filing.driveId or config.filing.libraries["ACTIVE JOBS"]');
  }

  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Missing MS_* credentials');
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

  const res = await graph
    .api(`/drives/${driveId}/root:/${encodePath(path)}:/children`)
    .select('name,folder,webUrl,id')
    .top(200)
    .get();

  return (res.value || [])
    .filter((it) => !!it.folder)
    .map((it) => ({ name: it.name, webUrl: it.webUrl, id: it.id }));
}

function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;
  const dp = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) dp[j] = j;
  for (let i = 1; i <= a.length; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1]
        ? prev
        : Math.min(prev, dp[j], dp[j - 1]) + 1;
      prev = tmp;
    }
  }
  return dp[b.length];
}

// Length-scaled tolerance for typos. Conservative for short keys (where one
// substitution shifts meaning — "Tom" vs "Tim") and more forgiving for long
// names (where a missing letter is more likely a typo).
function fuzzyTolerance(len) {
  if (len <= 4) return 0;
  if (len <= 7) return 1;
  return 2;
}

function matchFolder(jobKey, folders) {
  if (!jobKey) return null;
  const key = String(jobKey).toLowerCase().trim();
  if (!key) return null;
  const keyTokens = key.split(/\s+/).filter(Boolean);

  let bestSubstr = null;
  let bestFuzzy = null;
  let bestFuzzyDist = Infinity;

  for (const f of folders) {
    const name = String(f.name || '');
    const lower = name.toLowerCase();

    // Tier 1 — substring (preferred, e.g. "Baldivis" in
    // "Focus - Parkland Heights - Baldivis").
    if (lower.includes(key)) {
      if (!bestSubstr || name.length < bestSubstr.name.length) bestSubstr = f;
      continue;
    }

    // Tier 2 — per-token Levenshtein, to catch typos like "Eneabba" vs
    // "Eneaba - Relay Room". For multi-word jobKeys, every key-token must
    // find a near match.
    const tokens = lower.split(/[^a-z0-9]+/).filter(Boolean);
    let worstDist = 0;
    let allMatched = true;
    for (const kt of keyTokens) {
      const tol = fuzzyTolerance(kt.length);
      let bestForKt = Infinity;
      for (const t of tokens) {
        const d = levenshtein(kt, t);
        if (d < bestForKt) bestForKt = d;
        if (bestForKt === 0) break;
      }
      if (bestForKt > tol) { allMatched = false; break; }
      if (bestForKt > worstDist) worstDist = bestForKt;
    }
    if (allMatched && worstDist < bestFuzzyDist) {
      bestFuzzy = f;
      bestFuzzyDist = worstDist;
    }
  }

  // Substring beats fuzzy. Otherwise the closest fuzzy match wins.
  return bestSubstr || bestFuzzy;
}

function gqlMonday(apiKey, query, variables) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ query, variables });
    const req = https.request(
      {
        hostname: 'api.monday.com',
        path: '/v2',
        method: 'POST',
        headers: {
          Authorization: apiKey,
          'Content-Type': 'application/json',
          'API-Version': '2024-10',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => {
          try { resolve(JSON.parse(data)); }
          catch (e) { reject(new Error(`Monday parse error: ${e.message} body=${data.slice(0, 200)}`)); }
        });
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function writeLink(apiKey, itemId, url) {
  // Monday's link column accepts a JSON-encoded { url, text } value, passed
  // as a JSON-typed variable.
  const valueObj = { url, text: 'SharePoint Folder' };
  const mutation = `
    mutation ($boardId: ID!, $itemId: ID!, $colId: String!, $val: JSON!) {
      change_column_value(board_id: $boardId, item_id: $itemId, column_id: $colId, value: $val) { id }
    }
  `;
  const result = await gqlMonday(apiKey, mutation, {
    boardId: String(JOB_TRACKER_BOARD_ID),
    itemId: String(itemId),
    colId: LINK_COLUMN_ID,
    val: JSON.stringify(valueObj),
  });
  if (result.errors) {
    throw new Error('Monday mutation error: ' + JSON.stringify(result.errors));
  }
  return result.data.change_column_value.id;
}

async function syncJobLinks({ dryRun = false } = {}) {
  const apiKey = process.env.MONDAY_API_KEY;
  if (!apiKey) throw new Error('Missing MONDAY_API_KEY in .env');

  // Read jobs (uses the same fetcher as the dashboard, including jobAliases).
  const manualPath = require('path').resolve(__dirname, '../dashboard/manualAllocations.json');
  const manual = JSON.parse(require('fs').readFileSync(manualPath, 'utf8'));
  const { jobs } = await fetchJobs(manual);

  // All real jobs are eligible for a SharePoint link — including Completed
  // ones (so historical jobs that have folders still get linked). Anything
  // else (no stage / unknown stage) is skipped.
  const ELIGIBLE_STAGES = new Set(['In Progress', 'Not begun', 'Completed']);
  const linkable = jobs.filter((j) => ELIGIBLE_STAGES.has(j.stage));

  console.log(`\nLinkable jobs: ${linkable.length} (of ${jobs.length} total)`);
  console.log(`Mode: ${dryRun ? '🔍 DRY RUN (no writes)' : '✏️  LIVE'}\n`);

  const folders = await listActiveJobsFolders();
  console.log(`SharePoint folders under ${config.filing.libraries['ACTIVE JOBS']}: ${folders.length}\n`);

  const stats = { alreadyLinked: 0, noMatch: 0, wrote: 0, wouldWrite: 0, errors: 0 };

  for (const job of linkable) {
    const tag = `[${job.stage}] ${job.name}`;
    if (job.sharePointUrl) {
      console.log(`  ⏭   ${tag}`);
      console.log(`        already linked → ${job.sharePointUrl}`);
      stats.alreadyLinked++;
      continue;
    }

    const match = matchFolder(job.jobKey, folders);
    if (!match) {
      console.log(`  ❓  ${tag}`);
      console.log(`        no SharePoint folder matched jobKey="${job.jobKey}"`);
      stats.noMatch++;
      continue;
    }

    if (dryRun) {
      console.log(`  📝  ${tag}`);
      console.log(`        WOULD WRITE → ${match.webUrl}`);
      console.log(`        (folder: "${match.name}")`);
      stats.wouldWrite++;
      continue;
    }

    try {
      await writeLink(apiKey, job.id, match.webUrl);
      console.log(`  ✅  ${tag}`);
      console.log(`        wrote → ${match.webUrl}`);
      console.log(`        (folder: "${match.name}")`);
      stats.wrote++;
    } catch (err) {
      console.log(`  ❌  ${tag}`);
      console.log(`        FAILED → ${err.message}`);
      stats.errors++;
    }
  }

  console.log('\n=== Summary ===');
  console.log(`  Jobs scanned        : ${linkable.length}`);
  console.log(`  Already linked      : ${stats.alreadyLinked}`);
  console.log(`  No folder match     : ${stats.noMatch}`);
  if (dryRun) {
    console.log(`  Would write         : ${stats.wouldWrite}`);
  } else {
    console.log(`  Wrote               : ${stats.wrote}`);
    console.log(`  Errors              : ${stats.errors}`);
  }

  return stats;
}

module.exports = { syncJobLinks, matchFolder, listActiveJobsFolders };
