/**
 * Rule-based email → job classifier.
 *
 * Scores each routing rule against the email's sender domain/address,
 * subject, and bodyPreview. Returns the highest-scoring active job match,
 * or null if no rule scored above 0. The matchAnyJobName field is used to
 * resolve a rule's target job by name-substring match against the live
 * Monday active-jobs list (so we don't have to hardcode item IDs in the
 * config).
 */

function classifyEmail(message, activeJobs, routing) {
  const { rules, xeroNotifications, scoring } = routing;

  const fromDomain = extractDomain(message.from || '');
  const fromAddress = (message.from || '').toLowerCase();
  const subjectLower = (message.subject || '').toLowerCase();
  const previewLower = (message.preview || '').toLowerCase();
  const haystack = `${subjectLower} ${previewLower}`;

  // Xero notification — handled separately; for v1 we just tag as xero, no
  // job-routing (extracting invoice → job mapping is v2). Treat as unmatched
  // so it surfaces in the unclassified bucket.
  if (xeroNotifications && xeroNotifications.senderDomains.some((d) => fromDomain === d)) {
    return null;
  }

  let best = null;
  for (const rule of rules) {
    let score = 0;
    if (rule.senderDomains.some((d) => fromDomain === d)) score += scoring.senderDomain;
    if (rule.senderAddresses.some((a) => a.toLowerCase() === fromAddress)) score += scoring.senderAddress;
    if (rule.keywords.some((k) => haystack.includes(k.toLowerCase()))) score += scoring.keyword;
    if (score === 0) continue;

    const job = activeJobs.find((j) =>
      rule.matchAnyJobName.some((needle) => j.name.toLowerCase().includes(needle.toLowerCase())),
    );
    if (!job) continue;

    score += scoring.matchAnyJobName;
    if (!best || score > best.score) {
      best = { jobId: job.id, jobName: job.name, score, rule };
    }
  }
  return best;
}

function extractDomain(address) {
  const at = address.indexOf('@');
  if (at === -1) return '';
  return address.slice(at + 1).toLowerCase();
}

module.exports = { classifyEmail, extractDomain };
