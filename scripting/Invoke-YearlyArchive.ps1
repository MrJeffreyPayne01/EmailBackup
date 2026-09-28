<#
.SYNOPSIS
    Runs Move-OutlookMailToArchive.ps1 across every year-based archive PST in order.

.DESCRIPTION
    Move-OutlookMailToArchive.ps1's -BeforeYear is an exclusive upper bound with no lower
    bound, so each pass takes everything the previous pass left behind. That makes pass
    ORDER load-bearing: oldest first. Running 2023 before Pre-2020 would sweep decades of
    mail into 2023Mail.pst.

    This script enforces that ordering so it cannot be got wrong by hand.

    Passes: Pre-2020, then 2020, 2021, 2022, 2023 (bounded by -ThroughYear).

.PARAMETER SourceFolderPath
    Folder to archive from. Pass the full store-qualified path if your mail is not in the
    default delivery store, e.g. 'me@gmail.com\Inbox'.

.PARAMETER ArchiveDirectory
    Directory holding the PST files.

.PARAMETER ThroughYear
    Last year to process. Defaults to 2023.

.PARAMETER MaxItemsPerPass
    Cap on items considered per pass. Use a small number for a rehearsal.

.PARAMETER IncludeSubfolders
    Recurse into subfolders, mirroring the tree into each PST.

.EXAMPLE
    .\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -WhatIf

.EXAMPLE
    .\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItemsPerPass 25

.EXAMPLE
    .\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -Confirm:$false
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [string]$SourceFolderPath,

    [string]$ArchiveDirectory = 'E:\PSTArchiveFiles',

    [ValidateRange(2020, 2100)]
    [int]$ThroughYear = 2023,

    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxItemsPerPass = [int]::MaxValue,

    [switch]$IncludeSubfolders
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$archiveScript = Join-Path $PSScriptRoot 'Move-OutlookMailToArchive.ps1'
if (-not (Test-Path -LiteralPath $archiveScript)) {
    throw "Move-OutlookMailToArchive.ps1 was not found beside this script ($PSScriptRoot)."
}

# Ordered oldest-first. Pre-2020 must run before any single-year pass.
$passes = @(
    [pscustomobject]@{ Pst = 'Pre2020Mail.pst'; Folder = 'Pre-2020'; BeforeYear = 2020 }
)
foreach ($year in 2020..$ThroughYear) {
    $passes += [pscustomobject]@{
        Pst        = "${year}Mail.pst"
        Folder     = "$year"
        BeforeYear = $year + 1
    }
}

Write-Host "`nArchive plan for $SourceFolderPath" -ForegroundColor White
$passes | Format-Table @{ N = 'Pass'; E = { $_.Folder } },
                       @{ N = 'PST'; E = { $_.Pst } },
                       @{ N = 'Takes mail received before'; E = { "$($_.BeforeYear)-01-01" } } |
    Out-String | Write-Host

$totals = [System.Collections.Generic.List[object]]::new()

foreach ($pass in $passes) {
    $pstPath = Join-Path $ArchiveDirectory $pass.Pst

    Write-Host ('=' * 72) -ForegroundColor DarkGray
    Write-Host "Pass: $($pass.Folder)  ->  $pstPath" -ForegroundColor White
    Write-Host ('=' * 72) -ForegroundColor DarkGray

    $splat = @{
        PstPath          = $pstPath
        TargetFolderName = $pass.Folder
        SourceFolderPath = $SourceFolderPath
        BeforeYear       = $pass.BeforeYear
        Confirm          = $false          # confirmation is handled once, here
    }
    if ($MaxItemsPerPass -ne [int]::MaxValue) { $splat.MaxItems = $MaxItemsPerPass }
    if ($IncludeSubfolders) { $splat.IncludeSubfolders = $true }
    if ($WhatIfPreference) { $splat.WhatIf = $true }

    if (-not $PSCmdlet.ShouldProcess("$SourceFolderPath -> $($pass.Folder) in $($pass.Pst)", 'Archive pass')) {
        continue
    }

    try {
        & $archiveScript @splat
        $totals.Add([pscustomobject]@{ Pass = $pass.Folder; Status = 'Completed' })
    }
    catch {
        Write-Warning "Pass '$($pass.Folder)' failed: $($_.Exception.Message)"
        $totals.Add([pscustomobject]@{ Pass = $pass.Folder; Status = "Failed: $($_.Exception.Message)" })
    }
}

Write-Host "`nSummary" -ForegroundColor White
$totals | Format-Table -AutoSize | Out-String | Write-Host
