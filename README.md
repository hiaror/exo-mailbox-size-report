## Exchange Online Mailbox Size Report

This repository contains a PowerShell script that generates mailbox size reports for Exchange Online using a CSV driven input.

The script was created during a tenant to tenant migration project where early visibility into mailbox sizes was required for planning and decision making.

---

## What the script does

- Reads a list of mailboxes from a CSV file
- Retrieves primary and archive mailbox statistics from Exchange Online
- Converts mailbox size values into numeric GB values within PowerShell
- Exports a clean CSV report ready for sorting and analysis

No post processing in Excel is required.

---

## Why this approach

Exchange Online mailbox statistics are returned as formatted strings that are not immediately usable for calculations or sorting.

This script normalizes size values during execution so the exported report can be used directly without additional Excel formulas or manual cleanup.

This makes it easier to:
- Sort mailboxes by size
- Identify large mailboxes early
- Use the output for migration planning

---

## Folder structure

scripts/
  Invoke-EXOMailboxSizeReport.ps1

sample-data/
  Mailboxes.sample.csv

---

## Input CSV format

The input CSV must contain a column named UPN.

Example:

UPN
user1@domain.com
sharedmailbox@domain.com

---

## Example execution

.\Invoke-EXOMailboxSizeReport.ps1 `
  -InputCsv ".\migrationmailboxes.csv" `
  -CSVPath ".\MailboxSizeReport_YYYYMMDD_HHMMSS.csv"

---

## Notes

- The script uses Exchange Online PowerShell
- Progress is displayed during execution
- Archive mailbox properties are loaded explicitly
- Errors are handled per mailbox without stopping the entire run

---

## Disclaimer

This repository contains sanitized sample data only.

The script is provided as a reference implementation and should be reviewed before use in any production environment.
