# Exchange Online Mailbox Size Report

## Overview
This repository contains a PowerShell script that generates mailbox size reports for **Exchange Online** using a **CSV-driven UPN list as input**.

The script normalises Exchange Online's `TotalItemSize` byte-quantified strings into numeric GB values during execution, so the exported report is sortable and analysis-ready without any post-processing in Excel.

## Why This Exists

In tenant-to-tenant migration projects, early visibility into mailbox sizes is required for batch planning, licensing decisions, and storage analysis.

Exchange Online mailbox statistics are returned as formatted strings that are not immediately usable for calculations or sorting. This script normalises those values during execution so the exported report can be used directly without additional Excel formulas or manual cleanup.

## Key Capabilities
- CSV-driven input using a single UPN column
- Primary and archive mailbox statistics in a single report
- Numeric GB column for sorting and aggregation
- Per-mailbox try/catch so individual failures do not stop the run
- Friendly text size column retained for readability
- Sanitized sample data for reference

## Repository Structure
```
.
├── scripts/
│   └── Invoke-EXOMailboxSizeReport.ps1
├── sample-data/
│   └── Mailboxes.sample.csv
└── README.md
```

## Prerequisites
- PowerShell 5.1 or PowerShell 7.x
- ExchangeOnlineManagement v3 module installed
- An account with the View-Only Recipients role (or higher) in Exchange Online

## CSV Input Format

### Mailboxes CSV
```csv
UPN
alex.lee@northshore.example
nina.kapoor@northshore.example
shared.finance@northshore.example
```

## Usage Examples

### Standard Run
```powershell
./Invoke-EXOMailboxSizeReport.ps1 `
  -InputCsv "./sample-data/Mailboxes.sample.csv" `
  -CSVPath  "./MailboxSizeReport_YYYYMMDD_HHMMSS.csv"
```

### Quiet Run (No Progress Bar)
```powershell
./Invoke-EXOMailboxSizeReport.ps1 `
  -InputCsv "./sample-data/Mailboxes.sample.csv" `
  -CSVPath  "./MailboxSizeReport_YYYYMMDD_HHMMSS.csv" `
  -NoProgress
```

## Reporting
Each run produces a timestamped CSV report capturing:
- UserPrincipalName
- PrimaryTotalItemSize (friendly text)
- PrimarySizeGB (numeric)
- ArchiveEnabled
- ArchiveTotalItemSize (friendly text)
- ArchiveSizeGB (numeric)
- ArchiveStatsError

## Safety Notes
The script is read-only against Exchange Online and produces a report file only. Always review the output before using it for batch planning or licensing decisions. Sample data and identifiers in this repository are sanitized.

## Disclaimer
Provided as-is for reference and learning purposes.

## Blog Post

A full write-up of the script, the output format, and usage examples is at [AroraMSP: Exchange Online mailbox size reporting with PowerShell](https://aroramsp.com/blog/exo-mailbox-size-report).
