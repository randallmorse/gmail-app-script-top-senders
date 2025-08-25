/**
 * Fast Top Senders using GmailApp batching (recommended)
 * - Batches with GmailApp.getMessagesForThreads() for up to 100 threads/call
 * - Filters messages by explicit after:/before: dates if present in QUERY
 * - Optional: exclude no-reply senders
 * - Optional: also write a Top Domains sheet
 * - Writes to sheets in a single bulk setValues()
 */

const CONFIG = {
  QUERY: 'in:inbox after:2025/07/01',  // <-- adjust as needed
  THREAD_PAGE_SIZE: 500,               // threads fetched per page via GmailApp.search()
  THREAD_CHUNK_SIZE: 100,              // getMessagesForThreads() in batches (<= 100 recommended)
  TOP_N_SENDERS: 0,                    // 0 = no limit; set to e.g. 100 for Top 100 senders
  EXCLUDE_NOREPLY: true,               // skip "no-reply" / "donotreply" style senders
  WRITE_SENDERS_SHEET: true,           // write the per-sender results
  WRITE_DOMAINS_SHEET: true,           // also write aggregated per-domain results
  SENDERS_SHEET_NAME: 'Top Senders',
  DOMAINS_SHEET_NAME: 'Top Domains',
  CLEAR_SHEETS: true,                  // clear previous content before writing
  LOG_EVERY: 2000,                     // log every N processed messages
  WRITE_TIMESTAMP: true                // add timestamp in C1:D1 of each sheet
};

/**
 * Entry point
 */
function getTopSendersFast() {
  const t0 = Date.now();

  // Parse explicit date bounds only if QUERY contains after:YYYY/MM/DD (or -) and/or before:YYYY/MM/DD (or -)
  const bounds = parseDateBounds_(CONFIG.QUERY); // {after: Date|null, before: Date|null}
  const enforceBounds = Boolean(bounds.after || bounds.before);

  const senderCounts = Object.create(null); // { email: count }
  const domainCounts = Object.create(null); // { domain: count }

  let totalProcessed = 0;
  let start = 0;

  log_('Starting (GmailApp)…');
  log_(`Query: "${CONFIG.QUERY}"`);
  if (bounds.after) log_(`Message filter: >= ${bounds.after.toISOString()}`);
  if (bounds.before) log_(`Message filter:  < ${bounds.before.toISOString()}`);

  try {
    while (true) {
      // 1) Search threads page-by-page
      const threads = GmailApp.search(CONFIG.QUERY, start, CONFIG.THREAD_PAGE_SIZE);
      if (!threads || threads.length === 0) {
        if (start === 0) log_('No threads found for the query.');
        break;
      }

      // 2) Fetch messages for up to 100 threads per call for speed
      for (let i = 0; i < threads.length; i += CONFIG.THREAD_CHUNK_SIZE) {
        const chunk = threads.slice(i, i + CONFIG.THREAD_CHUNK_SIZE);

        let threadMsgs;
        try {
          threadMsgs = GmailApp.getMessagesForThreads(chunk); // 2D array [thread][message]
        } catch (e) {
          log_('Error in getMessagesForThreads(): ' + e);
          continue; // continue with next chunk
        }

        // 3) Count senders (per message), with optional date filtering and no-reply exclusion
        for (const msgs of threadMsgs) {
          for (const msg of msgs) {
            const msgDate = msg.getDate(); // Date
            if (enforceBounds) {
              if (bounds.after && msgDate < bounds.after) continue;
              if (bounds.before && msgDate >= bounds.before) continue;
            }

            const fromVal = msg.getFrom(); // "Name <email@domain>" or "email@domain"
            const email = extractEmail_(fromVal);
            if (!email) continue;

            if (CONFIG.EXCLUDE_NOREPLY && isNoReply_(email)) continue;

            // Count sender
            senderCounts[email] = (senderCounts[email] || 0) + 1;

            // Count domain
            const domain = extractDomain_(email);
            if (domain) domainCounts[domain] = (domainCounts[domain] || 0) + 1;

            totalProcessed++;
            if (CONFIG.LOG_EVERY > 0 && totalProcessed % CONFIG.LOG_EVERY === 0) {
              const rate = (totalProcessed / ((Date.now() - t0) / 1000)).toFixed(1);
              log_(`Processed ${totalProcessed} messages… @ ${rate} msg/s`);
            }
          }
        }
      }

      start += threads.length;
      if (threads.length < CONFIG.THREAD_PAGE_SIZE) break; // last page
    }

    log_(`Finished. Total messages processed: ${totalProcessed}`);

    // 4) Sort & (optionally) limit Top N for senders
    if (CONFIG.WRITE_SENDERS_SHEET) {
      let sendersSorted = Object.entries(senderCounts).sort(([, a], [, b]) => b - a);
      if (CONFIG.TOP_N_SENDERS && CONFIG.TOP_N_SENDERS > 0) {
        sendersSorted = sendersSorted.slice(0, CONFIG.TOP_N_SENDERS);
      }
      writeTableToSheet_(CONFIG.SENDERS_SHEET_NAME, ['Sender', 'Email Count'], sendersSorted);
    }

    // 5) Also write Top Domains (unlimited by default; adjust here if you want a TOP_N_DOMAINS)
    if (CONFIG.WRITE_DOMAINS_SHEET) {
      const domainsSorted = Object.entries(domainCounts).sort(([, a], [, b]) => b - a);
      writeTableToSheet_(CONFIG.DOMAINS_SHEET_NAME, ['Domain', 'Email Count'], domainsSorted);
    }

  } catch (err) {
    log_('Overall error: ' + err);
    throw err; // make visible in Executions
  }
}

/** ----------------- Helpers ----------------- **/

function writeTableToSheet_(sheetName, headerRow, rows) {
  const sheet = ensureSheet_(sheetName);
  if (CONFIG.CLEAR_SHEETS) sheet.clearContents();

  const data = [headerRow].concat(rows);
  if (data.length === 1) {
    // Only header
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  } else {
    sheet.getRange(1, 1, data.length, headerRow.length).setValues(data);
  }

  if (CONFIG.WRITE_TIMESTAMP) {
    sheet.getRange(1, headerRow.length + 1, 1, 2).setValues([['Generated At', new Date()]]);
  }

  log_(`Wrote ${rows.length} rows to "${sheetName}".`);
}

function extractEmail_(fromValue) {
  if (!fromValue) return null;
  // Prefer the address inside angle brackets
  const angle = fromValue.match(/<([^>]+)>/);
  if (angle && angle[1]) return normalizeEmail_(angle[1]);

  // Otherwise, find any email-like token
  const emailLike = fromValue.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return emailLike ? normalizeEmail_(emailLike[0]) : null;
}

function normalizeEmail_(s) {
  return String(s).trim().replace(/\s+/g, '').toLowerCase();
}

function extractDomain_(email) {
  const at = email.indexOf('@');
  return at > 0 ? email.slice(at + 1).toLowerCase() : null;
}

function isNoReply_(email) {
  const local = email.split('@')[0] || '';
  return /(no[-_.\s]?reply|do[-_.\s]?not[-_.\s]?reply)/i.test(local);
}

/**
 * Parses explicit after:/before: date bounds from the query in YYYY/MM/DD or YYYY-MM-DD.
 * If you use "newer_than:Xd" or "older_than:Xd", there is no explicit date to parse.
 */
function parseDateBounds_(query) {
  const parse = (s) => {
    if (!s) return null;
    // Supports YYYY/MM/DD or YYYY-MM-DD
    const parts = s.includes('-') ? s.split('-') : s.split('/');
    if (parts.length !== 3) return null;
    const [y, m, d] = parts.map(Number);
    if (!y || !m || !d) return null;
    // Interpret as local timezone midnight
    return new Date(y, m - 1, d);
  };
  const afterMatch = query.match(/\bafter:([0-9]{4}[\/-][0-9]{2}[\/-][0-9]{2})\b/);
  const beforeMatch = query.match(/\bbefore:([0-9]{4}[\/-][0-9]{2}[\/-][0-9]{2})\b/);
  return { after: parse(afterMatch && afterMatch[1]), before: parse(beforeMatch && beforeMatch[1]) };
}

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function log_(m) {
  Logger.log(`[TopSenders] ${m}`);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Top Senders')
    .addItem('Run (Fast)', 'getTopSendersFast')
    .addToUi();
}
