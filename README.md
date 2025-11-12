# Gmail_Analysis
This script analyzes your Gmail and writes a concise report into a Google Sheet. It scans messages (excluding spam and trash), aggregates key metrics, and presents organized breakdowns so you can quickly understand your inbox activity.


# Main Functions
## Gmail querying
Searches your Gmail using a configurable query and ignores spam/trash.
## Date handling
Identifies the oldest message date and focuses on messages from the last 365 days.
## Data aggregation
Computes totals such as number of threads, messages, unread counts, threads with attachments, and approximate total email size in bytes.
## Breakdowns
Summarizes messages by sender and by sender domain to show top contacts and domains.
## Output
Writes results into a Google Sheet with sections for overall metrics, sender summaries, domain summaries, and label summaries.
## Error handling
Logs errors and checks permissions to ensure it can access and write to the target spreadsheet.


# Why it’s useful
Provides quick visibility into inbox volume, storage usage, and who/which domains contact you most often.
Automates reporting into a readable Google Sheet for review or further analysis.


# How to use
## 1. Create a Google Sheet
### 1.1 Go to https://drive.google.com and sign in with your Google account.
### 1.2 Click the New button (left side) and choose Google Sheets to create a blank spreadsheet.
### 1.3 Rename the sheet: click the title at the top left and enter name "Gmail Analysis".
## 2. Open Google Apps Script
### 2.1 In the Google Sheet, open the menu Extensions → Apps Script.
### 2.2 A new tab opens with the Apps Script editor.
## 3. Add Your Script Code
### 3.1 In the Apps Script editor, remove any placeholder code in the script file.
### 3.2 Paste your script code into the editor.
### 3.3 If you have set the name of the Google sheet as "Gmail Analysis", then there shouldn't be any need to change the code else, change the variable in the code to match the name of the Google sheet

## 4. Save, Authorize, and Run
### 4.1 Save the project using the disk icon or File → Save.
### 4.2 To run the script, select the function you want (for example, "buildReport.gs") and click the Run button (▶).
###  4.3 On the first run you must authorize the script: follow the prompts to grant required permissions (e.g., access to Gmail and Google Sheets).
###  4.4 After granting permissions, run the function again to execute the script.
### 4.5 View Results
### 5.1 Return to the Google Sheet tab.
###  5.2 The script should populate the sheet with the report or data created by your script.
## Troubleshooting
If the script fails, open the Apps Script console and check the error messages.
Ensure you granted the necessary permissions (Gmail, Sheets, etc.).
Confirm the SPREADSHEET_ID in your script is correct and points to the sheet you created.
Verify the sheet name or ranges used in the script match those in your spreadsheet.

