/**
 * SharePoint source — Active Quotes folder.
 *
 * STATUS: stubbed. The existing MS_* app credentials returned 403 when probed
 * against this drive, so this source returns an empty list and a flag that
 * the dashboard composer can surface as a warning. Once Sites.Read.All
 * (or Files.Read.All) is granted to the app and admin-consented in Azure,
 * replace `fetchQuotes` with the real Graph call:
 *
 *     GET /drives/{driveId}/root:/Active Quotes:/children
 *
 * driveId: b!4B_dBjoPKEeuMh4XWxgxCQcyBxLlmwFLgtOWOYvkPTQGCZjiE9CmRoLS7JMod59t
 *
 * The script `scripts/probe-sharepoint.js` re-checks access — once it
 * succeeds, swap this stub out for the real implementation.
 */

require('isomorphic-fetch');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

const DRIVE_ID =
  'b!4B_dBjoPKEeuMh4XWxgxCQcyBxLlmwFLgtOWOYvkPTQGCZjiE9CmRoLS7JMod59t';

async function fetchQuotes() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    return { quotes: [], available: false, reason: 'MS_* credentials missing' };
  }

  const msal = new ConfidentialClientApplication({
    auth: { clientId, clientSecret, authority: `https://login.microsoftonline.com/${tenantId}` },
  });
  const graph = Client.init({
    authProvider: async (done) => {
      try {
        const result = await msal.acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        done(null, result.accessToken);
      } catch (err) { done(err, null); }
    },
  });

  // Path discovered via scripts/probe-sharepoint.js — Active Quotes lives
  // inside an emoji-prefixed parent folder.
  const QUOTES_PATH = 'CORECAST/📋 QUOTES/Active Quotes';
  try {
    // $expand=listItem($expand=fields) brings in the SharePoint custom
    // columns (e.g. QuoteValue) alongside the basic driveItem properties.
    const aq = await graph
      .api(`/drives/${DRIVE_ID}/root:/${QUOTES_PATH}:/children`)
      .expand('listItem($expand=fields)')
      .top(100)
      .get();
    const quotes = (aq.value || []).map((item) => {
      const fields = item.listItem?.fields || {};
      // SharePoint library has two value columns: `Value` (the one the UI
      // surfaces for new entries) and `QuoteValue` (older, kept for legacy
      // rows like Ring Beam). Prefer Value; fall back to QuoteValue.
      const rawValue = fields.Value != null && fields.Value !== ''
        ? fields.Value
        : fields.QuoteValue;
      const quoteValue = rawValue == null || rawValue === '' ? null : Number(rawValue);
      // QuoteSent is a SharePoint Yes/No column. Graph omits the field when
      // it's never been set, returns true/false when it has — so we treat
      // missing/null as "not yet sent" (i.e. needs quoting).
      const quoteSent = fields.QuoteSent === true;
      return {
        name: item.name,
        isFolder: !!item.folder,
        size: item.size || null,
        lastModified: item.lastModifiedDateTime || null,
        webUrl: item.webUrl || null,
        quoteValue: Number.isFinite(quoteValue) ? quoteValue : null,
        quoteSent,
      };
    });
    return { quotes, available: true };
  } catch (err) {
    return {
      quotes: [],
      available: false,
      reason: `Graph ${err.statusCode || ''} ${err.code || ''}: ${err.message}`,
    };
  }
}

module.exports = { fetchQuotes };
