const fs = require('fs');
const path = require('path');

function loadDomains(filePath) {
  const abs = path.resolve(filePath);
  const raw = fs.readFileSync(abs, 'utf8');
  const parsed = JSON.parse(raw);
  if (!parsed.domains || typeof parsed.domains !== 'object') {
    throw new Error(`${abs}: missing top-level "domains" object`);
  }
  // Normalise keys to lowercase so lookup is case-insensitive without per-call work.
  const normalised = {};
  for (const [domain, dest] of Object.entries(parsed.domains)) {
    normalised[domain.toLowerCase()] = dest || null;
  }
  return normalised;
}

function extractDomain(emailAddress) {
  if (!emailAddress || typeof emailAddress !== 'string') return null;
  const at = emailAddress.lastIndexOf('@');
  if (at === -1) return null;
  return emailAddress.slice(at + 1).toLowerCase();
}

function matchDomain(emailAddress, domains) {
  const domain = extractDomain(emailAddress);
  if (!domain) return null;
  if (Object.prototype.hasOwnProperty.call(domains, domain)) {
    return { domain, destination: domains[domain] };
  }
  return null;
}

module.exports = { loadDomains, matchDomain, extractDomain };
