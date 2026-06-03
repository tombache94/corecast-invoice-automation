/**
 * One-off probe — verifies that the existing MS_* credentials can read the
 * Active Quotes folder in SharePoint. Lists at most the first 50 entries by
 * name only (no contents fetched). Run once, expect either success or a 403.
 */

require('isomorphic-fetch');
require('dotenv').config();

const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

const DRIVE_ID =
  'b!4B_dBjoPKEeuMh4XWxgxCQcyBxLlmwFLgtOWOYvkPTQGCZjiE9CmRoLS7JMod59t';

async function main() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    console.error('Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET in .env');
    process.exit(1);
  }

  const msal = new ConfidentialClientApplication({
    auth: { clientId, clientSecret, authority: `https://login.microsoftonline.com/${tenantId}` },
  });

  let scopesGranted = '(unknown)';
  const graph = Client.init({
    authProvider: async (done) => {
      try {
        const result = await msal.acquireTokenByClientCredential({
          scopes: ['https://graph.microsoft.com/.default'],
        });
        scopesGranted = result.scopes ? result.scopes.join(' ') : '(none reported)';
        done(null, result.accessToken);
      } catch (err) {
        done(err, null);
      }
    },
  });

  console.log('Probing SharePoint drive', DRIVE_ID);
  console.log('Tenant:', tenantId);
  console.log('App (client) ID:', clientId);
  console.log();

  // 1) Confirm the drive exists and we can read its metadata.
  let drive;
  try {
    drive = await graph.api(`/drives/${DRIVE_ID}`).get();
    console.log('✅ /drives/{id} succeeded');
    console.log('   Drive name :', drive.name);
    console.log('   Drive type :', drive.driveType);
    console.log('   Owner      :', drive.owner?.user?.displayName || drive.owner?.group?.displayName || '(unknown)');
    console.log('   Web URL    :', drive.webUrl);
    console.log('   Token scope claim:', scopesGranted);
  } catch (err) {
    console.log('❌ /drives/{id} failed');
    console.log('   Status :', err.statusCode);
    console.log('   Code   :', err.code);
    console.log('   Message:', err.message);
    console.log('   Token scope claim:', scopesGranted);
    process.exit(2);
  }
  console.log();

  // 2) List the root of the drive to see what folders are at the top level.
  try {
    const root = await graph
      .api(`/drives/${DRIVE_ID}/root/children`)
      .select('name,folder,file,size,lastModifiedDateTime')
      .top(50)
      .get();
    console.log(`✅ Root listing — ${root.value.length} entries:`);
    for (const item of root.value) {
      const kind = item.folder ? `folder(${item.folder.childCount ?? '?'})` : `file(${item.size ?? '?'})`;
      console.log(`   ${kind}  ${item.name}`);
    }
  } catch (err) {
    console.log('❌ Root listing failed');
    console.log('   Status :', err.statusCode);
    console.log('   Code   :', err.code);
    console.log('   Message:', err.message);
    process.exit(3);
  }
  console.log();

  // 3) List CORECAST contents to find where Active Quotes lives.
  try {
    const cc = await graph
      .api(`/drives/${DRIVE_ID}/root:/CORECAST:/children`)
      .select('name,folder,file,size,lastModifiedDateTime')
      .top(50)
      .get();
    console.log(`✅ /CORECAST — ${cc.value.length} entries:`);
    for (const item of cc.value) {
      const kind = item.folder ? `folder(${item.folder.childCount ?? '?'})` : `file(${item.size ?? '?'})`;
      console.log(`   ${kind}  ${item.name}`);
    }
  } catch (err) {
    console.log('❌ /CORECAST listing failed:', err.statusCode, err.code, err.message);
  }
  console.log();

  // 4) Try the Active Quotes folder under CORECAST.
  try {
    const aq = await graph
      .api(`/drives/${DRIVE_ID}/root:/CORECAST/Active Quotes:/children`)
      .select('name,folder,file,size,lastModifiedDateTime')
      .top(50)
      .get();
    console.log(`✅ /CORECAST/Active Quotes — ${aq.value.length} entries:`);
    for (const item of aq.value) {
      const kind = item.folder ? `folder(${item.folder.childCount ?? '?'})` : `file(${item.size ?? '?'})`;
      const mod = item.lastModifiedDateTime ? item.lastModifiedDateTime.slice(0, 10) : '';
      console.log(`   ${kind}  ${item.name}  ${mod}`);
    }
  } catch (err) {
    console.log('❌ "/CORECAST/Active Quotes" listing failed');
    console.log('   Status :', err.statusCode);
    console.log('   Code   :', err.code);
    console.log('   Message:', err.message);
    process.exit(4);
  }
}

main().catch((err) => {
  console.error('Fatal:', err.message);
  process.exit(99);
});
