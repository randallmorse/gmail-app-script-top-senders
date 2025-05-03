function getTopSenders() {
  var senders = {};
  var query = 'in:inbox after:2025/04/20';
  //var query = 'in:inbox';
  var threads, pageToken;

  try {
    do {
      var results = Gmail.Users.Threads.list('me', { q: query, pageToken: pageToken });
      threads = results.threads || [];

      threads.forEach(function(thread) {
        try {
          var messages = GmailApp.getThreadById(thread.id).getMessages();
          messages.forEach(function(message) {
            var sender = message.getFrom();
            senders[sender] = (senders[sender] || 0) + 1;
          });
        } catch (threadError) {
          Logger.log('Error processing thread: ' + thread.id + ' - ' + threadError);
        }
      });

      pageToken = results.nextPageToken;
    } while (pageToken);

    // Sort senders by frequency
    var sortedSenders = Object.entries(senders).sort(([, a], [, b]) => b - a);

    // Output results to a spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    sheet.clearContents(); // Clear previous results

    sheet.appendRow(['Sender', 'Email Count']); // Header row

    for (var k = 0; k < sortedSenders.length; k++) {
      sheet.appendRow([sortedSenders[k][0], sortedSenders[k][1]]);
    }

    Logger.log('Top senders processed successfully.');
    Logger.log(sortedSenders);

  } catch (error) {
    Logger.log('An error occurred: ' + error);
  }
}
