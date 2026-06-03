const { config, assertMicrosoftCreds } = require('../config');
const { OutlookClient } = require('../email/outlook');
const { SharePointClient } = require('../sharepoint/client');
const { FilingStore } = require('./state');
const { loadKeywords, matchSubject } = require('./keywords');
const { loadDomains, matchDomain, extractDomain } = require('./domains');

function resolveLibraryPath(libraryName) {
  const p = config.filing.libraries[libraryName];
  if (!p) {
    throw new Error(
      `Unknown library "${libraryName}". Configured libraries: ${Object.keys(config.filing.libraries).join(', ')}`,
    );
  }
  return p;
}

function pickKeywordDestination(match) {
  if (!match.primary) return null;
  const libraryPath = resolveLibraryPath(match.primary.destination.library);
  return {
    tier: 'keyword',
    label: `${match.primary.destination.library} / ${match.primary.destination.folder}`,
    keyword: match.primary.keyword,
    folderPath: `${libraryPath}/${match.primary.destination.folder}`,
  };
}

function pickDomainDestination(domainHit) {
  if (!domainHit) return null;
  // Per-domain override takes precedence over the global Accounts Payable path.
  if (domainHit.destination && domainHit.destination.library && domainHit.destination.folder) {
    const libraryPath = resolveLibraryPath(domainHit.destination.library);
    return {
      tier: 'domain',
      label: `${domainHit.destination.library} / ${domainHit.destination.folder}  [supplier: ${domainHit.domain}]`,
      domain: domainHit.domain,
      folderPath: `${libraryPath}/${domainHit.destination.folder}`,
    };
  }
  return {
    tier: 'domain',
    label: `Accounts Payable  [supplier: ${domainHit.domain}]`,
    domain: domainHit.domain,
    folderPath: config.filing.accountsPayablePath,
  };
}

function isSafeAttachment(att) {
  // Skip inline images / signatures and non-file attachments. We only file
  // user-attached documents (PDFs, DOCX, XLSX, etc).
  if (att['@odata.type'] !== '#microsoft.graph.fileAttachment') return false;
  if (att.isInline) return false;
  if (!att.contentBytes) return false;
  return true;
}

function decodeAttachment(att) {
  return Buffer.from(att.contentBytes, 'base64');
}

function formatDigestEntry(e) {
  const where = e.tier === 'keyword'
    ? `job folder (${e.label})`
    : e.tier === 'domain'
    ? `Accounts Payable (${e.label})`
    : `To Sort`;
  const link = e.webUrl ? `\n      ${e.webUrl}` : '';
  return `  • ${e.attachmentName}\n      from ${e.from || 'unknown'} — "${e.subject || '(no subject)'}"\n      → ${where}${link}`;
}

async function sendDigest(outlook, digest) {
  const totals = {
    keyword: digest.filter((d) => d.tier === 'keyword').length,
    domain:  digest.filter((d) => d.tier === 'domain').length,
    toSort:  digest.filter((d) => d.tier === 'to_sort').length,
  };
  const today = new Date().toISOString().slice(0, 10);
  const subject = totals.toSort > 0
    ? `[CoreCast] Filing summary ${today} — ${totals.toSort} need review`
    : `[CoreCast] Filing summary ${today} — all auto-filed`;

  const sections = [];
  sections.push(
    `Attachment filer ran ${new Date().toISOString()}.`,
    `  Filed to job folders : ${totals.keyword}`,
    `  Filed to Accts Payable: ${totals.domain}`,
    `  Sent to "To Sort"    : ${totals.toSort}` +
      (totals.toSort > 0 ? '   ← needs manual filing' : ''),
    '',
  );

  if (totals.toSort > 0) {
    sections.push('=== To Sort — needs manual review ===');
    for (const e of digest.filter((d) => d.tier === 'to_sort')) {
      sections.push(formatDigestEntry(e));
    }
    sections.push(
      '',
      'For each one: move it in SharePoint to the correct job folder, then',
      'either add the job keyword to config/job-keywords.json or add the',
      'sender domain to config/supplier-domains.json so future emails route',
      'automatically.',
      '',
    );
  }

  if (totals.keyword > 0) {
    sections.push('=== Filed to job folders ===');
    for (const e of digest.filter((d) => d.tier === 'keyword')) {
      sections.push(formatDigestEntry(e));
    }
    sections.push('');
  }

  if (totals.domain > 0) {
    sections.push('=== Filed to Accounts Payable ===');
    for (const e of digest.filter((d) => d.tier === 'domain')) {
      sections.push(formatDigestEntry(e));
    }
    sections.push('');
  }

  await outlook.sendMail(config.filing.notificationSender, {
    to: config.filing.notificationRecipient,
    subject,
    text: sections.join('\n'),
  });
}

async function fileAttachments(args = {}, deps = {}) {
  assertMicrosoftCreds();

  const dryRun = !!args.dryRun;
  const limit = args.limit || config.filing.fetchLimit;
  const lookbackHours = args.lookbackHours ?? config.filing.lookbackHours;
  // --since wins; otherwise compute a rolling window from lookbackHours.
  // Pass --since '' (empty) or set lookbackHours=0 to disable the window.
  let sinceISO = args.sinceISO || null;
  let sinceSource = 'explicit';
  if (!sinceISO && lookbackHours > 0) {
    sinceISO = new Date(Date.now() - lookbackHours * 3600 * 1000).toISOString();
    sinceSource = `lookback ${lookbackHours}h`;
  } else if (!sinceISO) {
    sinceSource = 'none (full inbox window up to limit)';
  }
  const mailboxes = args.mailbox ? [args.mailbox] : config.mailboxes;

  const outlook = deps.outlook || new OutlookClient(config.microsoft);
  const sharepoint =
    deps.sharepoint ||
    new SharePointClient({ ...config.microsoft, driveId: config.filing.driveId });
  const store = deps.store || new FilingStore(config.filing.stateFile);
  const keywords = deps.keywords || loadKeywords(config.filing.keywordConfigFile);
  const domains = deps.domains || loadDomains(config.filing.supplierDomainsFile);

  const stats = {
    messagesScanned: 0,
    attachmentsConsidered: 0,
    alreadyFiled: 0,
    skippedOwnDomain: 0,
    matchedByKeyword: 0,
    matchedByDomain: 0,
    sentToSort: 0,
    errors: 0,
  };

  // Track unique unmatched sender domains so Tom can populate supplier-domains.json.
  const unmatchedSenderTally = new Map(); // domain -> { count, examples: [{from, subject}] }
  // Collected at filing time, sent as a single digest at end-of-run.
  const digest = []; // { tier, label, attachmentName, from, subject, webUrl }
  const skipDomains = new Set(config.filing.skipDomains || []);

  console.log(
    `\nFiling attachments — dryRun=${dryRun} mailboxes=[${mailboxes.join(', ')}] limit=${limit}` +
      (sinceISO ? ` since=${sinceISO} (${sinceSource})` : ` since=none`) +
      `\n`,
  );
  console.log(
    `Loaded ${Object.keys(keywords).length} job keywords from ${config.filing.keywordConfigFile}`,
  );
  console.log(
    `Loaded ${Object.keys(domains).length} supplier domains from ${config.filing.supplierDomainsFile}\n`,
  );

  for (const mailbox of mailboxes) {
    let summaries;
    try {
      summaries = await outlook.listRecentMessages(mailbox, { sinceISO, top: limit });
    } catch (err) {
      console.log(`! ${mailbox}: list failed — ${err.message}\n`);
      stats.errors++;
      continue;
    }

    console.log(`[${mailbox}] ${summaries.length} message(s) scanned`);

    for (const s of summaries) {
      stats.messagesScanned++;
      if (!s.hasAttachments) continue;

      let full;
      try {
        full = await outlook.getMessage(mailbox, s.id, { includeAttachments: true });
      } catch (err) {
        console.log(`  ! ${s.id}: fetch failed — ${err.message}`);
        stats.errors++;
        continue;
      }

      const attachments = (full.attachments || []).filter(isSafeAttachment);
      if (!attachments.length) continue;

      const senderEmail = full.from?.emailAddress?.address || '';
      const senderDomain = extractDomain(senderEmail);
      if (senderDomain && skipDomains.has(senderDomain)) {
        console.log(`  ⏭   ${attachments.length} attachment(s) from ${senderEmail} — skipped (own domain)`);
        stats.skippedOwnDomain += attachments.length;
        continue;
      }

      const keywordMatch = matchSubject(s.subject, keywords);
      const keywordDest = pickKeywordDestination(keywordMatch);
      const domainDest = keywordDest
        ? null
        : pickDomainDestination(matchDomain(senderEmail, domains));
      const dest = keywordDest || domainDest;

      for (const att of attachments) {
        stats.attachmentsConsidered++;

        const identity = {
          internetMessageId: full.internetMessageId || s.internetMessageId || null,
          attachmentName: att.name,
          attachmentSize: att.size,
          mailbox,
          messageId: s.id,
          attachmentId: att.id,
        };
        if (await store.hasBeenFiled(identity)) {
          stats.alreadyFiled++;
          continue;
        }

        const targetPath = dest ? dest.folderPath : config.filing.toSortPath;
        const outcome = dest ? (dest.tier === 'keyword' ? 'matched_keyword' : 'matched_domain') : 'to_sort';
        const label = dest
          ? `→ ${dest.label}`
          : `→ To Sort (no match)`;

        console.log(`  ${att.name}  [${senderEmail || 'unknown'}]  ${label}`);

        if (dest && dest.tier === 'keyword') stats.matchedByKeyword++;
        else if (dest && dest.tier === 'domain') stats.matchedByDomain++;
        else {
          stats.sentToSort++;
          const dom = extractDomain(senderEmail);
          if (dom) {
            const cur = unmatchedSenderTally.get(dom) || { count: 0, examples: [] };
            cur.count++;
            if (cur.examples.length < 2) {
              cur.examples.push({ from: senderEmail, subject: s.subject || '' });
            }
            unmatchedSenderTally.set(dom, cur);
          }
        }

        if (dryRun) {
          continue;
        }

        try {
          const uploaded = await sharepoint.uploadFile(
            targetPath,
            att.name,
            decodeAttachment(att),
          );
          await store.record({
            ...identity,
            subject: s.subject,
            from: senderEmail,
            keyword: keywordDest ? keywordDest.keyword : null,
            destination: targetPath,
            outcome,
            webUrl: uploaded?.webUrl || null,
            notes: keywordMatch.all.length > 1
              ? `also matched keywords: ${keywordMatch.all.slice(1).map((m) => m.keyword).join(', ')}`
              : domainDest
              ? `routed by supplier domain: ${domainDest.domain}`
              : '',
          });

          digest.push({
            tier: dest ? dest.tier : 'to_sort',
            label: dest ? dest.label : 'To Sort',
            attachmentName: att.name,
            from: senderEmail,
            subject: s.subject || '',
            webUrl: uploaded?.webUrl || null,
          });
        } catch (err) {
          // Roll back the optimistic stat increment so the summary reflects reality.
          if (dest && dest.tier === 'keyword') stats.matchedByKeyword--;
          else if (dest && dest.tier === 'domain') stats.matchedByDomain--;
          else stats.sentToSort--;

          // 409 nameAlreadyExists = the file is already in the destination
          // folder, almost certainly because a parallel filing process
          // (e.g. Cowork) uploaded it first. Treat as benign "already
          // filed", record in our state file so we never retry it, and
          // don't count it as an error.
          const isConflict =
            err.statusCode === 409 ||
            /nameAlreadyExists/i.test(err.message || '') ||
            /nameAlreadyExists/i.test(err.code || '');
          if (isConflict) {
            console.log(`    ⏭   already exists in destination (likely filed by a parallel process)`);
            stats.alreadyFiled++;
            try {
              await store.record({
                ...identity,
                subject: s.subject,
                from: senderEmail,
                keyword: keywordDest ? keywordDest.keyword : null,
                destination: targetPath,
                outcome: 'already_existed',
                webUrl: null,
                notes: '409 from upload — parallel process already filed this',
              });
            } catch (_) {
              // best-effort state write
            }
            continue;
          }

          console.log(`    ! upload failed — ${err.statusCode || ''} ${err.code || ''} ${err.message}`);
          stats.errors++;
        }
      }
    }

    console.log('');
  }

  console.log('=== Summary ===');
  console.log(`  Messages scanned:           ${stats.messagesScanned}`);
  console.log(`  Attachments considered:     ${stats.attachmentsConsidered}`);
  console.log(`  Skipped (own domain):       ${stats.skippedOwnDomain}`);
  console.log(`  Already filed (skipped):    ${stats.alreadyFiled}`);
  console.log(`  Filed to job folder:        ${stats.matchedByKeyword}`);
  console.log(`  Filed to Accounts Pay'l:    ${stats.matchedByDomain}`);
  console.log(`  Sent to "To Sort":          ${stats.sentToSort}`);
  console.log(`  Errors:                     ${stats.errors}`);
  if (dryRun) console.log(`  (dry run — no uploads or notifications were sent)`);

  // Send a single digest email summarising everything filed this run.
  // Skip when nothing was filed (most runs will be quiet), in dry-run mode,
  // OR when called with skipDigestEmail (refresh-dashboard.js consumes the
  // digest array itself and rolls it into the dashboard refresh email).
  const recipients = Array.isArray(config.filing.notificationRecipient)
    ? config.filing.notificationRecipient.join(', ')
    : config.filing.notificationRecipient;
  if (!dryRun && !args.skipDigestEmail && digest.length > 0) {
    try {
      await sendDigest(outlook, digest);
      console.log(`\n  📧 digest sent to ${recipients} (${digest.length} entries)`);
    } catch (err) {
      console.log(`\n  ! digest email failed — ${err.message}`);
    }
  } else if (dryRun && digest.length > 0) {
    console.log(`\n  (dry run — would send digest with ${digest.length} entries to ${recipients})`);
  } else if (args.skipDigestEmail && digest.length > 0) {
    console.log(`\n  (digest with ${digest.length} entries handed back to caller — no separate email)`);
  }

  if (unmatchedSenderTally.size > 0) {
    console.log(`\n=== Unmatched sender domains (would go to To Sort) ===`);
    const sorted = [...unmatchedSenderTally.entries()].sort((a, b) => b[1].count - a[1].count);
    for (const [domain, info] of sorted) {
      console.log(`  ${domain}  (${info.count} attachment${info.count === 1 ? '' : 's'})`);
      for (const ex of info.examples) {
        console.log(`     e.g. ${ex.from}  —  ${ex.subject}`);
      }
    }
    console.log(
      `\n  To route any of these to Accounts Payable, add to config/supplier-domains.json:`,
    );
    console.log(`    { "domains": { "supplier.com.au": {} } }`);
  }

  return { stats, digest };
}

module.exports = {
  fileAttachments,
  pickKeywordDestination,
  pickDomainDestination,
  isSafeAttachment,
};
