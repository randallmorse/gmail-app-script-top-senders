/**
 * ARCHIVE INBOX BY DOMAINS (Consolidated List)
 * This script will archive every email currently in your Inbox from the 270+ domains listed.
 */

const DOMAINS_TO_ARCHIVE = [
  // --- Replace with your list ---
  "yourlisthere.com", "yourlisthere.net", "yourlisthere.org", "yourlisthere.us"
];

function bulkArchiveDomains() {
  const BATCH_SIZE = 25; // Adjusted batch size for safer processing
  let totalArchived = 0;

  // 1. Remove duplicates from the array just in case
  const uniqueDomains = [...new Set(DOMAINS_TO_ARCHIVE)];
  
  console.log(`Starting archive for ${uniqueDomains.length} unique domains.`);

  for (let i = 0; i < uniqueDomains.length; i += BATCH_SIZE) {
    const chunk = uniqueDomains.slice(i, i + BATCH_SIZE);
    
    // Construct query: in:inbox (from:domain1.com OR from:domain2.com...)
    const query = `in:inbox (from:${chunk.join(' OR from:')})`;
    
    console.log(`Checking batch: ${chunk[0]}... and others`);

    let threads;
    do {
      // Find up to 100 threads at a time
      threads = GmailApp.search(query, 0, 100);
      if (threads.length > 0) {
        GmailApp.moveThreadsToArchive(threads);
        totalArchived += threads.length;
        console.log(`Archived ${threads.length} threads in this batch.`);
      }
    } while (threads.length === 100); 
    // If threads.length is 100, there are likely more; loop again for the same chunk
  }

  console.log(`Finished! Total threads archived: ${totalArchived}`);
  //SpreadsheetApp.getUi().alert(`Success! ${totalArchived} threads moved to Archive.`);
}