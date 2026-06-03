#!/usr/bin/env node
/**
 * Probe the current M365 app token to see which Microsoft Graph
 * application permissions (roles) have been granted-with-consent.
 *
 * Decodes the JWT's `roles` claim (no signature verification — we just
 * want to read the claims). Also tests Mail.Send by attempting a
 * benign sendMail with malformed payload — a permission failure
 * surfaces differently than a payload failure, so we can tell which
 * is which.
 */

require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');

function decodeJwt(token) {
  const parts = token.split('.');
  if (parts.length !== 3) throw new Error('Not a JWT');
  const payload = Buffer.from(parts[1], 'base64url').toString('utf8');
  return JSON.parse(payload);
}

(async () => {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    console.error('Missing MS_* env vars');
    process.exit(1);
  }

  const msal = new ConfidentialClientApplication({
    auth: { clientId, clientSecret, authority: `https://login.microsoftonline.com/${tenantId}` },
  });
  const result = await msal.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  const claims = decodeJwt(result.accessToken);

  console.log('=== App registration ===');
  console.log(`  App ID    : ${claims.appid || claims.azp || '(missing)'}`);
  console.log(`  Tenant ID : ${claims.tid}`);
  console.log(`  App name  : ${claims.app_displayname || '(not in token)'}`);
  console.log(`  Token type: ${claims.idtyp || '?'}  (app = application, user = delegated)`);
  console.log('');
  console.log('=== Granted application permissions (roles claim) ===');
  const roles = Array.isArray(claims.roles) ? claims.roles : [];
  if (!roles.length) {
    console.log('  (none granted!)');
  } else {
    for (const r of roles.sort()) console.log(`  • ${r}`);
  }
  console.log('');
  console.log('=== Mail-relevant capability check ===');
  const hasSend = roles.includes('Mail.Send');
  const hasSendShared = roles.includes('Mail.Send.Shared');
  const hasRead = roles.includes('Mail.Read') || roles.includes('Mail.ReadBasic.All') || roles.includes('Mail.ReadWrite');
  const hasReadWrite = roles.includes('Mail.ReadWrite');
  console.log(`  Mail.Send         : ${hasSend ? '✅' : '❌'}  ${hasSend ? '(can send mail as any user — needed for outbound notifications)' : '(MISSING — cannot send mail)'}`);
  console.log(`  Mail.Send.Shared  : ${hasSendShared ? '✅' : '❌'}  (send on behalf of users who delegated access)`);
  console.log(`  Mail.Read / .ReadWrite : ${hasRead ? '✅' : '❌'}  ${hasRead ? '(used by file-attachments + Cowork inbox scanning)' : ''}`);
  console.log(`  Mail.ReadWrite    : ${hasReadWrite ? '✅' : '❌'}  (needed to mark messages as read, move to folders, etc.)`);
  console.log('');

  if (hasSend) {
    console.log('=== Verifying send actually works (best-effort) ===');
    // We don't actually send mail here — just confirm the endpoint accepts
    // our auth. A 401 = permission denied; a 400/422 = permission OK but
    // payload bad (we use a malformed payload deliberately).
    try {
      const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent('tom@corecastconcrete.com.au')}/sendMail`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${result.accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ /* deliberately empty — Graph will 400 */ }),
      });
      if (res.status === 401 || res.status === 403) {
        console.log(`  ❌  HTTP ${res.status} — endpoint refused auth. Permission claim says Mail.Send but real call failed.`);
      } else {
        console.log(`  ✅  HTTP ${res.status} — endpoint accepted auth (expected 4xx because payload was deliberately empty). Send pipe is wired.`);
      }
    } catch (err) {
      console.log(`  ⚠️  Probe call failed: ${err.message}`);
    }
  } else {
    console.log('=== How to add Mail.Send ===');
    console.log('');
    console.log('  1. Open the Azure portal: https://portal.azure.com');
    console.log('  2. Go to: Microsoft Entra ID → App registrations → [your app]');
    console.log(`     (App ID: ${claims.appid || claims.azp})`);
    console.log('  3. Left sidebar → API permissions → Add a permission');
    console.log('  4. Choose: Microsoft Graph → Application permissions');
    console.log('  5. Find and tick: "Mail.Send"');
    console.log('  6. Click "Add permissions"');
    console.log('  7. Back on API permissions, click "Grant admin consent for [tenant]"');
    console.log('     (requires you to be a Global Admin or Privileged Role Admin)');
    console.log('  8. Wait ~60 seconds for the change to propagate');
    console.log('  9. Re-run this probe script to verify Mail.Send appears in the roles list');
    console.log('');
    console.log('  Note: Mail.Send is "send mail as any user" by default. To restrict it');
    console.log('  to a specific mailbox (e.g. accounts@corecastconcrete.com.au only),');
    console.log('  configure an Application Access Policy via Exchange Online PowerShell:');
    console.log('  https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access');
  }
})().catch((err) => {
  console.error('Fatal:', err.message);
  if (process.env.DEBUG) console.error(err.stack);
  process.exit(1);
});
