# Mailbox PST Export Script with Yearly Segmentation

This PowerShell script automates the export of Exchange mailboxes to PST files, segmented by year, using a mailbox list from a CSV file.

## Features

- 🔄 Processes mailboxes listed in a CSV file
- 📅 Splits export requests by year (last 10 years)
- 📨 Uses `New-MailboxExportRequest` for Exchange environments
- 📁 Exports PST files to custom paths
- 📊 Generates a simple HTML report for monitoring (optional)

## Prerequisites

- Exchange Management Shell
- Mailbox export permissions
- UNC file path for export destination
- Exchange server must be configured to allow PST exports

## CSV Format

The input CSV should be named `mailboxes_export_list.csv` with the following structure:

```csv
Mailbox,OutputPath
user1@domain.com,\\server\PSTExports\
user2@domain.com,\\server\PSTExports\
```

## How It Works

The script:
1. Loads mailbox and export path data from the CSV.
2. Creates 10 yearly date ranges.
3. Loops through each mailbox and creates one PST export request per year.
4. Optionally, generates an HTML report with export statuses.

## Example Output Path

```
\\server\PSTExports\user1@domain.com - 2020.pst
\\server\PSTExports\user1@domain.com - 2021.pst
...
```

## Monitoring Exports

You can monitor the export requests via:

```powershell
Get-MailboxExportRequest | Get-MailboxExportRequestStatistics
```

To remove completed requests:

```powershell
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest
```

## Author

**FA-IT Consulting**

---

Feel free to customize or extend the script as needed.

