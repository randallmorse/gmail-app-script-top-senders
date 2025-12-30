Gmail Top Senders Purge Scripts

This workflow analyzes a Google Sheets spreadsheet to identify your top email senders, allows you to review and filter out unwanted domains, and then purges emails from those domains using Apps Script.
The process runs in two scripts: one for identifying top senders, and another for purging emails from selected domains.
Timeouts are expected; the purge script may run in multiple passes and can be scheduled to run repeatedly until complete.

Prerequisites

Google account with access to Google Sheets and Google Apps Script
The spreadsheet containing your emails (or a reference sheet with sender data)
Basic familiarity with Google Apps Script
Two JavaScript files:
top_senders.js (identifies top senders)
purge_emails.js (purges emails for specified domains)
Optional: a temporary trigger to run the purge script at intervals (e.g., every 10 minutes)

What you’ll need to configure

PURGE_DOMAINS: a list of domains you want to purge (copied from Step 2)
API/configuration settings required by your scripts (documented within the files)

Step-by-step Instructions

Step 1 — Identify Top Senders

Open your Google Sheet containing the data you want to analyze.
Go to Extensions > Apps Script.
Paste the contents of top_senders.js into the Apps Script editor.
Save and run the function that identifies top senders.
If prompted for authorization, grant the required permissions.
Document any settings or API configuration required by top_senders.js (as noted inside the script).

Step 2 — Filter Domains to Purge

Review the list of domains produced by top_senders.js.
Quickly skim to identify domains you don’t care about purging.
Copy the domains you want to purge and prepare them for the purge script.
Paste the copied domain list into the header section of purge_emails.js (as directed by the script’s comments or header instructions).

Step 3 — Purge Emails by Domain

Open https://script.google.com/
Create a new Apps Script project.
Paste the contents of purge_emails.js into the new project.
Update DOMAINS_TO_ARCHIVE with the list you copied in Step 2 (example format below).
Run the purge script.
If prompted, authorize the script with the necessary permissions.
Document any required API keys or settings referenced by purge_emails.js.

Notes and tips

It’s normal for the purge script to fail on the first run due to Google Apps Script quotas and timeouts. This is expected behavior.
In practice, large purges (e.g., tens of thousands of emails) may take several attempts. A typical run can time out after several minutes; you may need to re-run until you observe progress (e.g., a decreasing inbox count or a log of removals per batch).
If you see it making progress, you can set a temporary time-based trigger to run the purge script every 10 minutes until completion. Check back the next morning to confirm completion.
Always ensure you have a backup or a way to restore data before purging.

Example configurations

DOMAINS_TO_ARCHIVE example (place this at the top of purge_emails.js as a constant): 

```
const DOMAINS_TO_ARCHIVE = [ “example1.com”, “example2.org”, “spamdomain.net” ];
```

Optional: API/configuration notes

Ensure any required API keys or OAuth scopes are enabled for both Apps Script projects.

Document where credentials are stored and how they are refreshed.

Troubleshooting

Script authorization errors: re-authorize the script and grant the necessary scopes.
No progress after a run: check for errors in the Apps Script execution log; verify PURGE_DOMAINS is correctly populated and that the script is targeting the right mailbox/spreadsheet.
Timeouts: implement a time-based trigger to run the script periodically, and monitor the inbox count to confirm progress.
If you’re purging a very large mailbox, consider splitting the domain list into smaller batches and running each batch separately.

Safety and best practices

Test on a small subset of data first to validate behavior.
Maintain an offline or cloud backup of emails before purging.
Use clear domain lists and avoid accidental purges of legitimate domains.