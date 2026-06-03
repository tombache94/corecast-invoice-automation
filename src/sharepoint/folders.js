/**
 * List folder names under the SharePoint ACTIVE JOBS and Active Quotes
 * libraries. These folder names are the source of truth for job/quote
 * identity in CoreCast — any folder that exists here represents a real
 * job or quote that downstream attribution should recognise.
 *
 * Used by src/dashboard/compose.js to build a dynamic JOB_KEYWORDS list
 * for matching Xero Reference fields against jobs/quotes, replacing the
 * old hardcoded keyword table.
 */

require('isomorphic-fetch');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const { config } = require('../config');

function encodePath(p) {
  return p.split('/').map((seg) => encodeURIComponent(seg)).join('/');
}

function createGraphClient() {
  const { tenantId, clientId, clientSecret } = config.microsoft;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Missing MS_* credentials');
  }
  const msal = new ConfidentialClientApplication({
    auth: { clientId, clientSecret, authority: `https://login.microsoftonline.com/${tenantId}` },
  });
  return Client.init({
    authProvider: async (done) => {
      try {
        const r = await msal.acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        done(null, r.accessToken);
      } catch (e) {
        done(e, null);
      }
    },
  });
}

async function listFoldersIn(graph, driveId, libraryPath) {
  const res = await graph
    .api(`/drives/${driveId}/root:/${encodePath(libraryPath)}:/children`)
    .select('name,folder,webUrl,id')
    .top(200)
    .get();
  return (res.value || [])
    .filter((it) => !!it.folder)
    .map((it) => ({ name: it.name, webUrl: it.webUrl, id: it.id }));
}

async function listJobAndQuoteFolders() {
  const driveId = config.filing.driveId;
  const jobsPath = config.filing.libraries['ACTIVE JOBS'];
  const quotesPath = config.filing.libraries['Active Quotes'];
  if (!driveId || !jobsPath || !quotesPath) {
    throw new Error('Missing SharePoint config — driveId / ACTIVE JOBS / Active Quotes path');
  }

  const graph = createGraphClient();
  const [jobsResult, quotesResult] = await Promise.allSettled([
    listFoldersIn(graph, driveId, jobsPath),
    listFoldersIn(graph, driveId, quotesPath),
  ]);

  const folders = [];
  const status = { jobs: { ok: false, count: 0, error: null }, quotes: { ok: false, count: 0, error: null } };

  if (jobsResult.status === 'fulfilled') {
    status.jobs.ok = true;
    status.jobs.count = jobsResult.value.length;
    for (const f of jobsResult.value) folders.push({ ...f, library: 'ACTIVE JOBS' });
  } else {
    status.jobs.error = jobsResult.reason?.message || String(jobsResult.reason);
  }

  if (quotesResult.status === 'fulfilled') {
    status.quotes.ok = true;
    status.quotes.count = quotesResult.value.length;
    for (const f of quotesResult.value) folders.push({ ...f, library: 'Active Quotes' });
  } else {
    status.quotes.error = quotesResult.reason?.message || String(quotesResult.reason);
  }

  return { folders, status };
}

module.exports = { listJobAndQuoteFolders };
