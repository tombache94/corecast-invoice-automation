const https = require('https');
const url = process.argv[2] || 'https://corecastconcrete.com.au';
https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (res) => {
  console.log('status      :', res.statusCode);
  console.log('server      :', res.headers.server);
  console.log('x-powered-by:', res.headers['x-powered-by']);
  console.log('---');
  let data = '';
  res.on('data', (c) => (data += c));
  res.on('end', () => {
    const fingerprints = [
      'lovable', 'vite', '/assets/', '/_next/', 'gatsby',
      'id="root"', 'id="__next"', 'data-vite', 'manifest.webmanifest',
      'webflow', 'wp-content', 'shopify', 'squarespace',
    ];
    console.log('substring hits:');
    for (const f of fingerprints) {
      const count = (data.match(new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi')) || []).length;
      if (count) console.log('  ' + f + '  → ' + count);
    }
    console.log('---');
    const scripts = [...data.matchAll(/<script[^>]*src="([^"]+)"[^>]*>/g)].map((m) => m[1]);
    const links   = [...data.matchAll(/<link[^>]*href="([^"]+\.(?:css|js))"[^>]*>/g)].map((m) => m[1]);
    console.log('scripts (first 6):');
    scripts.slice(0, 6).forEach((s) => console.log('  ' + s));
    console.log('asset links (first 6):');
    links.slice(0, 6).forEach((l) => console.log('  ' + l));
    console.log('---');
    console.log('html length:', data.length, 'chars');
    console.log('---');
    console.log('first 700 chars:');
    console.log(data.slice(0, 700));
  });
}).on('error', (e) => console.error('error:', e.message));
