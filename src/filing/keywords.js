const fs = require('fs');
const path = require('path');

function loadKeywords(filePath) {
  const abs = path.resolve(filePath);
  const raw = fs.readFileSync(abs, 'utf8');
  const parsed = JSON.parse(raw);
  if (!parsed.keywords || typeof parsed.keywords !== 'object') {
    throw new Error(`${abs}: missing top-level "keywords" object`);
  }
  for (const [kw, dest] of Object.entries(parsed.keywords)) {
    if (!dest || typeof dest !== 'object') {
      throw new Error(`Keyword "${kw}" has invalid destination (must be object)`);
    }
    if (!dest.library || !dest.folder) {
      throw new Error(`Keyword "${kw}" missing library or folder`);
    }
  }
  return parsed.keywords;
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Length-scaled tolerance for typos. Conservative for short keywords (where
// one substitution shifts meaning — "Tom" vs "Tim") and more forgiving for
// longer names (a missing/transposed letter is more likely a typo).
function fuzzyTolerance(len) {
  if (len <= 4) return 0;     // too short — exact only
  if (len <= 6) return 1;
  if (len <= 10) return 2;    // covers most job/location names (Eneabba=7, Baldivis=8, Mandurah=8, Joondalup=9)
  return 3;
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

function tokenize(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
}

// Single-word keywords only. Multi-word keywords (e.g. "Wattle Grove") stay
// in the regex tier — fuzzy multi-word matching would need per-token Levenshtein
// AND positional logic, which isn't worth the complexity for v1.
function fuzzyMatchKeyword(kw, subject) {
  const kwLower = kw.toLowerCase();
  if (/\s/.test(kwLower)) return null;
  const tol = fuzzyTolerance(kwLower.length);
  if (tol === 0) return null;
  const tokens = tokenize(subject);
  let best = null;
  for (const tok of tokens) {
    if (tok === kwLower) return { matched: tok, distance: 0 };
    if (Math.abs(tok.length - kwLower.length) > tol) continue;
    const dist = levenshtein(kwLower, tok);
    if (dist <= tol && (!best || dist < best.distance)) {
      best = { matched: tok, distance: dist };
    }
  }
  return best;
}

function matchSubject(subject, keywords) {
  if (!subject) return { primary: null, all: [] };
  const matches = [];
  for (const [kw, dest] of Object.entries(keywords)) {
    // Tier 1 — exact regex word-boundary match.
    const re = new RegExp(`\\b${escapeRegex(kw)}\\b`, 'i');
    if (re.test(subject)) {
      matches.push({ keyword: kw, destination: dest, exact: true });
      continue;
    }
    // Tier 2 — fuzzy single-token match (catches typos like "Enneaba" → "Eneabba").
    const fuzzy = fuzzyMatchKeyword(kw, subject);
    if (fuzzy) {
      matches.push({
        keyword: kw,
        destination: dest,
        exact: false,
        matched: fuzzy.matched,
        distance: fuzzy.distance,
      });
    }
  }
  // Exact beats fuzzy. Then longer keyword wins on conflict
  // ("Wattle Grove" beats a hypothetical "Grove").
  matches.sort((a, b) => {
    if (a.exact !== b.exact) return a.exact ? -1 : 1;
    return b.keyword.length - a.keyword.length;
  });
  return {
    primary: matches[0] || null,
    all: matches,
  };
}

module.exports = { loadKeywords, matchSubject, fuzzyMatchKeyword, levenshtein };
