<#
.SYNOPSIS
    This script automates the export of mailbox content to PST files for the past 10 years.

.DESCRIPTION
    - Mailboxes and export paths are read from a CSV file (columns: Mailbox, OutputPath).
    - For each mailbox, the script creates one export request per year over the last 10 years.
    - Each PST file will contain messages received during the corresponding year.
    - Optionally, export request statistics can be monitored and completed requests cleaned up.

.AUTHOR
    FA-IT Consulting
    May 2025
#>

# Load mailboxes and export paths from the CSV file
$mailboxes = Import-Csv -Path "C:\Scripts\mailboxes_export_list.csv"

# Get the current year
$currentYear = (Get-Date).Year

# Create date ranges for the past 10 years
$dateRanges = @()
for ($i = 0; $i -lt 10; $i++) {
    $year = $currentYear - $i
    $startDate = Get-Date -Day 1 -Month 1 -Year $year
    $endDate = Get-Date -Day 1 -Month 1 -Year ($year + 1)
    $dateRanges += @{ Year = $year; Start = $startDate; End = $endDate }
}

# Loop through each mailbox
foreach ($entry in $mailboxes) {
    $mailbox = $entry.Mailbox
    $outputPath = $entry.OutputPath

    foreach ($range in $dateRanges) {
        $year = $range.Year
        $startDate = $range.Start
        $endDate = $range.End
        $outputFile = Join-Path -Path $outputPath -ChildPath "$mailbox - $year.pst"

        Write-Host "Creating export request for $mailbox ($year)..."

        New-MailboxExportRequest -Mailbox $mailbox `
                                 -FilePath $outputFile `
                                 -ContentFilter {(Received -ge $startDate) -and (Received -lt $endDate)}
    }
}

# Monitor export requests (optional)
Get-MailboxExportRequest | Get-MailboxExportRequestStatistics

# Clean up completed requests (optional)
# Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest
