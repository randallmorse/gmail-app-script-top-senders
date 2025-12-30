function getTopSenders() {
  var senders = {};
  // Define your query here. Examples:
  // 'in:inbox after:2025/04/20' - Emails in inbox after a specific date
  // 'in:inbox' - All emails in inbox (can be very slow if large)
  // 'is:unread' - Unread emails
  // var query = 'in:inbox after:2025/04/20';

  var query = 'in:inbox';

  var messagesListResponse, pageToken;
  var messagesProcessed = 0; // Counter for logging progress

  try {
    Logger.log('Starting to fetch messages...');

    // Use a do-while loop to handle pagination of message lists
    do {
      // List messages in batches using the Gmail Advanced Service
      // maxResults can be up to 500
      messagesListResponse = Gmail.Users.Messages.list('me', {
        q: query,
        pageToken: pageToken,
        maxResults: 500 // Fetch up to 500 message IDs per call
      });

      var messagesList = messagesListResponse.messages || [];

      if (messagesList.length === 0) {
          Logger.log('No messages found with the current query.');
          break; // Exit loop if no messages are returned on the first page
      }

      // Process the fetched batch of message IDs
      messagesList.forEach(function(messageItem) {
         try {
           // Fetch ONLY the 'From' header for each message ID
           var messageDetails = Gmail.Users.Messages.get('me', messageItem.id, {
             format: 'metadata', // Request only metadata
             metadataHeaders: ['From'] // Specifically request the 'From' header
           });

           // Find the 'From' header in the metadata
           var fromHeader = messageDetails.payload.headers.find(function(header) {
             return header.name === 'From';
           });

           if (fromHeader) {
             var sender = fromHeader.value;

             // Optional: Clean up sender string (e.g., remove name and keep just email)
             // This regex attempts to extract the email address from "Name <email@example.com>" or just use the string if no angle brackets
             var emailMatch = sender.match(/<([^>]+)>/);
             var cleanSender = emailMatch ? emailMatch[1] : sender.trim();

             // Count the sender
             senders[cleanSender] = (senders[cleanSender] || 0) + 1;
             messagesProcessed++;

             // Log progress periodically to see it's working
             if (messagesProcessed % 1000 === 0) {
               Logger.log('Processed ' + messagesProcessed + ' messages...');
             }

           } else {
             Logger.log('Warning: Could not find From header for message ID: ' + messageItem.id);
           }

         } catch (messageDetailError) {
           // Log error for a specific message but try to continue with others
           Logger.log('Error processing message details for ID: ' + messageItem.id + ' - ' + messageDetailError);
         }
      });

      // Get the token for the next page of results
      pageToken = messagesListResponse.nextPageToken;

    } while (pageToken); // Continue the loop if there's a next page token

    Logger.log('Finished fetching message data. Total messages processed: ' + messagesProcessed);

    // --- Process and Output Results ---

    // Convert the senders object into an array of [sender, count] pairs
    var sortedSenders = Object.entries(senders);

    // Sort senders by count in descending order
    sortedSenders.sort(([, a], [, b]) => b - a);

    Logger.log('Sorting complete. Preparing to write to sheet.');

    // Get the active spreadsheet and sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    // Clear previous results (optional, but good practice)
    sheet.clearContents();

    // Prepare data for bulk writing
    var dataToWrite = [['Sender', 'Email Count']]; // Header row

    // Add sorted sender data
    sortedSenders.forEach(function(item) {
      dataToWrite.push([item[0], item[1]]);
    });

    // Write data to the sheet in one go using setValues (much faster than appendRow in a loop)
    if (dataToWrite.length > 1) { // Check if there's data besides the header
      var range = sheet.getRange(1, 1, dataToWrite.length, 2);
      range.setValues(dataToWrite);
    } else {
      // Just write the header if no data
      sheet.appendRow(['Sender', 'Email Count']);
    }


    Logger.log('Top senders list written to sheet successfully.');

  } catch (error) {
    // Log any errors that occur outside the inner loops
    Logger.log('An overall error occurred: ' + error);
  }
}
