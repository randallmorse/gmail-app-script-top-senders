const CONFIG = {
  // Your date range
  QUERY: 'in:inbox after:2025/12/01 before:2025/12/31', 
  MAX_MESSAGES: 10000, 
  EXCLUDE_NOREPLY: true,
  SENDERS_SHEET_NAME: 'Top Senders',
  DOMAINS_SHEET_NAME: 'Top Domains'
};

function getTopSendersAdvanced() {
  const t0 = Date.now();
  const senderCounts = Object.create(null);
  const domainCounts = Object.create(null);
  
  let pageToken = null;
  let totalProcessed = 0;
  let errorCount = 0;

  console.log(`Starting search: ${CONFIG.QUERY}`);

  try {
    do {
      const response = Gmail.Users.Messages.list('me', {
        q: CONFIG.QUERY,
        pageToken: pageToken,
        maxResults: 500 // Max allowed per page
      });

      if (!response.messages || response.messages.length === 0) break;

      // Process messages in this page
      for (const msg of response.messages) {
        try {
          // Fetch only the 'From' header
          const detail = Gmail.Users.Messages.get('me', msg.id, {
            format: 'metadata',
            metadataHeaders: ['From']
          });

          if (!detail || !detail.payload || !detail.payload.headers) {
            errorCount++;
            continue; 
          }

          const fromHeader = detail.payload.headers.find(h => h.name === 'From');
          if (!fromHeader) continue;

          const email = extractEmail_(fromHeader.value);
          if (!email) continue;
          if (CONFIG.EXCLUDE_NOREPLY && isNoReply_(email)) continue;

          // Update counts
          senderCounts[email] = (senderCounts[email] || 0) + 1;
          const domain = email.split('@')[1];
          if (domain) domainCounts[domain] = (domainCounts[domain] || 0) + 1;

          totalProcessed++;

          // Log progress every 500 messages
          if (totalProcessed % 500 === 0) {
            console.log(`Processed ${totalProcessed} messages...`);
          }

        } catch (msgError) {
          // This catches the "Empty response" error for a single message
          errorCount++;
          // If we hit a lot of errors, wait a second
          if (errorCount % 5 === 0) Utilities.sleep(500); 
          continue; 
        }
      }

      pageToken = response.nextPageToken;

      // Safety break to avoid hitting the 6-minute Apps Script limit
      const elapsedMins = (Date.now() - t0) / 1000 / 60;
      if (elapsedMins > 5) {
        console.warn("Approaching script timeout. Stopping early.");
        break;
      }

    } while (pageToken && totalProcessed < CONFIG.MAX_MESSAGES);

    // Write results
    writeToSheet_(CONFIG.SENDERS_SHEET_NAME, senderCounts, ['Sender', 'Count']);
    writeToSheet_(CONFIG.DOMAINS_SHEET_NAME, domainCounts, ['Domain', 'Count']);

    console.log(`Finished. Processed: ${totalProcessed}, Errors/Skipped: ${errorCount}, Time: ${(Date.now() - t0)/1000}s`);

  } catch (e) {
    console.error("Critical Error: " + e.toString());
  }
}

/** ----------------- Helpers ----------------- **/

function extractEmail_(fromValue) {
  // Handles "Name <email@domain.com>" or "email@domain.com"
  const match = fromValue.match(/<([^>]+)>/) || fromValue.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  return match ? match[1].toLowerCase().trim() : null;
}

function isNoReply_(email) {
  return /no[-_]?reply|donotreply/i.test(email);
}

function writeToSheet_(name, countsObj, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();

  const data = Object.entries(countsObj)
    .sort((a, b) => b[1] - a[1]) // Sort by count descending
    .map(([key, val]) => [key, val]);

  if (data.length > 0) {
    sheet.getRange(1, 1, 1, 2).setValues([headers]);
    sheet.getRange(2, 1, data.length, 2).setValues(data);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}