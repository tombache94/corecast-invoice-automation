require('dotenv').config();

const config = {
  microsoft: {
    tenantId: process.env.MS_TENANT_ID,
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
  },
  mailboxes: (process.env.MAILBOXES || 'tom@corecastconcrete.com.au,accounts@corecastconcrete.com.au')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean),
  mailboxFetchLimit: Number(process.env.MAILBOX_FETCH_LIMIT || 50),
  connecteam: {
    apiKey: process.env.CONNECTEAM_API_KEY,
    baseUrl: process.env.CONNECTEAM_BASE_URL || 'https://api.connecteam.com',
    mainTimeClockId: process.env.CONNECTEAM_TIME_CLOCK_ID || '13723871', // CoreCast's main time clock ID
  },
  storage: {
    invoicesFile: process.env.INVOICES_FILE || './data/invoices.json',
  },
  timesheets: {
    watchFolder: process.env.TIMESHEETS_FOLDER || './timesheets',
  },
  alerts: {
    reportRecipient: process.env.REPORT_RECIPIENT || 'tom@corecastconcrete.com.au',
    reportSender: process.env.REPORT_SENDER || 'accounts@corecastconcrete.com.au',
  },
  matcher: {
    hoursTolerance: Number(process.env.HOURS_TOLERANCE || 0.5),
  },
  xero: {
    clientId: process.env.XERO_CLIENT_ID || '',
    clientSecret: process.env.XERO_CLIENT_SECRET || '',
    tenantId: process.env.XERO_TENANT_ID || '024c8088-8095-482b-906e-f72d9b8acaee',
    refreshToken: process.env.XERO_REFRESH_TOKEN || '',
    accessToken: process.env.XERO_ACCESS_TOKEN || '',
    tokenCacheFile: process.env.XERO_TOKEN_CACHE || './data/xero-token.json',
  },
  filing: {
    driveId:
      process.env.SHAREPOINT_DRIVE_ID ||
      'b!4B_dBjoPKEeuMh4XWxgxCQcyBxLlmwFLgtOWOYvkPTQGCZjiE9CmRoLS7JMod59t',
    libraries: {
      'ACTIVE JOBS':
        process.env.SHAREPOINT_ACTIVE_JOBS_PATH || 'CORECAST/📁 ACTIVE JOBS',
      'Active Quotes':
        process.env.SHAREPOINT_ACTIVE_QUOTES_PATH || 'CORECAST/📋 QUOTES/Active Quotes',
    },
    toSortPath: process.env.SHAREPOINT_TO_SORT_PATH || 'CORECAST/To Sort',
    accountsPayablePath:
      process.env.SHAREPOINT_ACCOUNTS_PAYABLE_PATH || 'CORECAST/Accounts Payable',
    keywordConfigFile: process.env.JOB_KEYWORDS_FILE || './config/job-keywords.json',
    supplierDomainsFile:
      process.env.SUPPLIER_DOMAINS_FILE || './config/supplier-domains.json',
    stateFile: process.env.FILING_STATE_FILE || './data/filed-attachments.json',
    // Comma-separated list of digest-email recipients. Outlook.sendMail
    // accepts an array of addresses, one email goes out per run.
    notificationRecipient:
      (process.env.FILING_NOTIFY_RECIPIENT
        || 'tom@corecastconcrete.com.au,josh@corecastconcrete.com.au')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
    notificationSender:
      process.env.FILING_NOTIFY_SENDER || 'accounts@corecastconcrete.com.au',
    fetchLimit: Number(process.env.FILING_FETCH_LIMIT || 50),
    lookbackHours: Number(process.env.FILING_LOOKBACK_HOURS || 24),
    // Sender domains to skip entirely — internal forwards from your own
    // mailboxes are admin chatter, not job documents.
    skipDomains: (process.env.FILING_SKIP_DOMAINS || 'corecastconcrete.com.au')
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter(Boolean),
  },
};

function assertXeroCreds() {
  const fs = require('fs');
  const path = require('path');
  const { clientId, clientSecret, tenantId, refreshToken, accessToken, tokenCacheFile } = config.xero;
  const missing = [];
  if (!clientId) missing.push('XERO_CLIENT_ID');
  if (!clientSecret) missing.push('XERO_CLIENT_SECRET');
  if (!tenantId) missing.push('XERO_TENANT_ID');

  // A token cached on disk (rotated by the dashboard refresh and other
  // consumers) counts as a valid credential — it's where XeroBillsClient
  // actually reads the refresh_token from at init() time. Only fall back to
  // demanding env-var tokens when no usable cache file exists.
  let cachedTokenOK = false;
  if (tokenCacheFile) {
    try {
      const abs = path.resolve(tokenCacheFile);
      if (fs.existsSync(abs)) {
        const cached = JSON.parse(fs.readFileSync(abs, 'utf8'));
        if (cached.refresh_token || cached.access_token) cachedTokenOK = true;
      }
    } catch (_) {
      // unreadable cache — treat as no cache
    }
  }
  if (!refreshToken && !accessToken && !cachedTokenOK) {
    missing.push('XERO_REFRESH_TOKEN or XERO_ACCESS_TOKEN (none found in .env or token cache file)');
  }
  if (missing.length) {
    throw new Error(
      `Missing Xero credentials: ${missing.join(', ')}. ` +
        `Register a custom-connection app at https://developer.xero.com, run scripts/xero-auth.js once, and set these in .env.`,
    );
  }
}

function assertMicrosoftCreds() {
  const { tenantId, clientId, clientSecret } = config.microsoft;
  const missing = [];
  if (!tenantId) missing.push('MS_TENANT_ID');
  if (!clientId) missing.push('MS_CLIENT_ID');
  if (!clientSecret) missing.push('MS_CLIENT_SECRET');
  if (missing.length) {
    throw new Error(
      `Missing Microsoft Graph credentials: ${missing.join(', ')}. ` +
        `Register an Azure AD app, grant Mail.Read application permission, and set these in .env.`,
    );
  }
}

module.exports = { config, assertMicrosoftCreds, assertXeroCreds };
