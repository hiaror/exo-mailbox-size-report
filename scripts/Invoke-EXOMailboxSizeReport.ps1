<#
Invoke-EXOMailboxSizeReport.ps1

Release Notes:
1) CSV-driven Exchange Online mailbox size report using a UPN-only input column
2) Produces both a friendly text size and a numeric GB column for sorting
3) Handles archive mailboxes via an explicit ArchiveGuid check on Get-EXOMailbox
4) Per-mailbox try/catch keeps the run going on individual failures

Usage examples (SANITIZED):

Standard run:
.\Invoke-EXOMailboxSizeReport.ps1 `
  -InputCsv ".\sample-data\Mailboxes.sample.csv" `
  -CSVPath  ".\MailboxSizeReport_YYYYMMDD_HHMMSS.csv"

Quiet run (no progress bar):
.\Invoke-EXOMailboxSizeReport.ps1 `
  -InputCsv ".\sample-data\Mailboxes.sample.csv" `
  -CSVPath  ".\MailboxSizeReport_YYYYMMDD_HHMMSS.csv" `
  -NoProgress
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$InputCsv,

    [Parameter(Mandatory = $false)]
    [string]$CSVPath = (Join-Path $PWD ("MailboxSizeReport_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [switch]$NoProgress
)

# ----------------------------
# Helpers
# ----------------------------

function Get-BytesFromSizeString {
    param([string]$SizeString)

    if ($SizeString -match "\(([\d,]+)\s+bytes\)") {
        return [double]($matches[1] -replace ",", "")
    }

    return 0
}

# ----------------------------
# Connect and process mailboxes
# ----------------------------

Connect-ExchangeOnline -ShowBanner:$false | Out-Null

$users = Import-Csv -Path $InputCsv
$results = New-Object System.Collections.Generic.List[object]

$total = $users.Count
$i = 0

foreach ($u in $users) {

    $upn = (($u.UPN) + "").Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    $i++

    if (-not $NoProgress) {
        # Progress update (throttled for speed)
        if (($i % 10) -eq 0 -or $i -eq $total) {
            Write-Progress -Activity "Scanning mailboxes" -Status "$i / $total : $upn" -PercentComplete (($i / $total) * 100)
        }
    }

    try {
        # Load archive properties explicitly (fix for ArchiveEnabled detection)
        $mbx = Get-EXOMailbox -Identity $upn -PropertySets Archive -ErrorAction Stop

        # Primary stats
        $pStats = Get-EXOMailboxStatistics -Identity $upn -ErrorAction Stop
        $pBytes = Get-BytesFromSizeString ([string]$pStats.TotalItemSize)
        $pGB = [math]::Round(($pBytes / 1GB), 2)
        $pText = "{0:N2} GB" -f $pGB

        # Archive stats (only if archive exists)
        $archiveEnabled = $false
        $aGB = 0
        $aText = ""
        $archiveError = ""

        if ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [guid]::Empty) {
            $archiveEnabled = $true

            $aStats = Get-EXOMailboxStatistics -Identity $upn -Archive -ErrorAction Stop
            $aBytes = Get-BytesFromSizeString ([string]$aStats.TotalItemSize)
            $aGB = [math]::Round(($aBytes / 1GB), 2)
            $aText = "{0:N2} GB" -f $aGB
        }

        $results.Add([pscustomobject]@{
            UserPrincipalName    = $upn
            PrimaryTotalItemSize = $pText
            PrimarySizeGB        = $pGB
            ArchiveEnabled       = $archiveEnabled
            ArchiveTotalItemSize = $aText
            ArchiveSizeGB        = $aGB
            ArchiveStatsError    = $archiveError
        }) | Out-Null
    }
    catch {
        $results.Add([pscustomobject]@{
            UserPrincipalName    = $upn
            PrimaryTotalItemSize = ""
            PrimarySizeGB        = ""
            ArchiveEnabled       = ""
            ArchiveTotalItemSize = ""
            ArchiveSizeGB        = ""
            ArchiveStatsError    = $_.Exception.Message
        }) | Out-Null
    }
}

if (-not $NoProgress) {
    Write-Progress -Activity "Scanning mailboxes" -Completed
}

$results | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

Write-Host "Completed. Report exported to: $CSVPath"
