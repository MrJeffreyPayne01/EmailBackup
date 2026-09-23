<#
.SYNOPSIS
    Moves mail between two Outlook folders, with optional date filtering and batching.

.DESCRIPTION
    A general-purpose mover, used for example to restore items from Gmail's Trash back to
    the Inbox. Unlike Move-OutlookMailToArchive.ps1 it does not touch PST files: both the
    source and the target are ordinary folders you name explicitly.

    Safety features that matter on large runs:
      * Items are collected as EntryID/StoreID pairs before any move, because moving an
        item mutates the collection it came from.
      * Work proceeds in batches, re-querying the source each time, so an interruption
        costs one batch rather than the whole run.
      * Consecutive failures abort the run. A dead Outlook process returns
        "RPC server is unavailable" for every item, and without this guard the script
        would log thousands of identical warnings and report a successful-looking finish
        having moved nothing.

.PARAMETER SourceFolderPath
    Store-qualified source path, e.g. 'me@gmail.com\[Gmail]\Trash'.

.PARAMETER TargetFolderPath
    Store-qualified destination path, e.g. 'me@gmail.com\Inbox'.

.PARAMETER SinceDate
    Only move items received on or after this date.

.PARAMETER BeforeDate
    Only move items received strictly before this date.

.PARAMETER MaxItems
    Stop after considering this many items.

.PARAMETER BatchSize
    Items per batch. Smaller batches recover more gracefully from an Outlook crash.

.PARAMETER MaxConsecutiveFailures
    Abort after this many failures in a row. Almost always means Outlook has died.

.EXAMPLE
    .\Move-OutlookMailBetweenFolders.ps1 -SourceFolderPath 'me@gmail.com\[Gmail]\Trash' `
        -TargetFolderPath 'me@gmail.com\Inbox' -SinceDate '2024-01-01' -MaxItems 500 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string]$SourceFolderPath,
    [Parameter(Mandatory)][string]$TargetFolderPath,
    [datetime]$SinceDate,
    [datetime]$BeforeDate,
    [ValidateRange(1, [int]::MaxValue)][int]$MaxItems = [int]::MaxValue,
    [ValidateRange(1, 5000)][int]$BatchSize = 500,
    [ValidateRange(1, 1000)][int]$MaxConsecutiveFailures = 15,
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$olMailItem = 43

function Get-OutlookNamespace {
    $app = New-Object -ComObject Outlook.Application
    $ns = $app.GetNamespace('MAPI')
    $ns.Logon($null, $null, $false, $false)
    $ns
}

function Get-ChildFolder {
    param([Parameter(Mandatory)] $Parent, [Parameter(Mandatory)][string] $Name)
    for ($i = 1; $i -le $Parent.Folders.Count; $i++) {
        $c = $Parent.Folders.Item($i)
        if ($c.Name -eq $Name) { return $c }
    }
    return $null
}

function Resolve-Folder {
    param([Parameter(Mandatory)] $Namespace, [Parameter(Mandatory)][string] $Path)

    $segments = @($Path.Split('\') | Where-Object { $_ -ne '' })
    if ($segments.Count -eq 0) { throw 'Folder path is empty.' }

    $current = $null
    for ($i = 1; $i -le $Namespace.Folders.Count; $i++) {
        $root = $Namespace.Folders.Item($i)
        if ($root.Name -eq $segments[0]) { $current = $root; break }
    }
    if ($current) {
        $segments = @($segments | Select-Object -Skip 1)
    }
    else {
        $current = $Namespace.GetDefaultFolder(6).Parent
    }

    foreach ($seg in $segments) {
        $next = Get-ChildFolder -Parent $current -Name $seg
        if (-not $next) { throw "Folder '$seg' not found under '$($current.Name)'." }
        $current = $next
    }
    $current
}

$ns     = Get-OutlookNamespace
$source = Resolve-Folder -Namespace $ns -Path $SourceFolderPath
$target = Resolve-Folder -Namespace $ns -Path $TargetFolderPath

if ($source.EntryID -eq $target.EntryID) { throw 'Source and target are the same folder.' }

# Built at script scope: $PSBoundParameters inside a function refers to that function's own
# parameters, which silently dropped the date filter and moved the wrong items.
$clauses = @()
if ($PSBoundParameters.ContainsKey('SinceDate')) {
    $clauses += '"urn:schemas:httpmail:datereceived" >= ''{0:yyyy-MM-dd HH:mm}''' -f $SinceDate
}
if ($PSBoundParameters.ContainsKey('BeforeDate')) {
    $clauses += '"urn:schemas:httpmail:datereceived" < ''{0:yyyy-MM-dd HH:mm}''' -f $BeforeDate
}
$filter = if ($clauses.Count -gt 0) { '@SQL=' + ($clauses -join ' AND ') } else { $null }

if (($PSBoundParameters.ContainsKey('SinceDate') -or $PSBoundParameters.ContainsKey('BeforeDate')) -and -not $filter) {
    throw 'A date filter was requested but could not be built. Refusing to run unfiltered.'
}

Write-Host ''
Write-Host "Source      : $($source.FolderPath)"
Write-Host "Destination : $($target.FolderPath)"
Write-Host "Filter      : $(if ($filter) { $filter } else { '(none - all mail items)' })"
Write-Host "Batch size  : $BatchSize"
Write-Host ''

$moved = 0
$processed = 0
$log = [System.Collections.Generic.List[object]]::new()
$consecutiveFailures = 0
$aborted = $false

while ($processed -lt $MaxItems -and -not $aborted) {
    $items = $source.Items
    # No Sort: sorting the parent collection then restricting makes indexed access throw
    # "Index was outside the bounds of the array" on IMAP folders.
    $scope = if ($filter) { $items.Restrict($filter) } else { $items }

    $remaining = [Math]::Min($BatchSize, $MaxItems - $processed)
    $batch = [System.Collections.Generic.List[object]]::new()

    for ($i = $scope.Count; $i -ge 1; $i--) {
        if ($batch.Count -ge $remaining) { break }
        try { $item = $scope.Item($i) } catch { continue }
        if ($item.Class -ne $olMailItem) { continue }
        $batch.Add([pscustomobject]@{
            EntryID      = $item.EntryID
            StoreID      = $source.StoreID
            Subject      = $item.Subject
            ReceivedTime = $item.ReceivedTime
            SenderName   = $item.SenderName
        })
    }

    if ($batch.Count -eq 0) {
        Write-Host 'No further matching items.' -ForegroundColor Cyan
        break
    }

    Write-Host "Batch: $($batch.Count) item(s)..." -ForegroundColor Cyan

    foreach ($entry in $batch) {
        $processed++
        $label = '{0:yyyy-MM-dd}  {1}' -f $entry.ReceivedTime, $entry.Subject

        if (-not $PSCmdlet.ShouldProcess($label, "Move to $($target.FolderPath)")) { continue }

        try {
            $item = $ns.GetItemFromID($entry.EntryID, $entry.StoreID)
            [void]$item.Move($target)
            $moved++
            $consecutiveFailures = 0
            $log.Add($entry)
        }
        catch {
            $consecutiveFailures++
            if ($consecutiveFailures -le 3) {
                Write-Warning "Failed to move '$($entry.Subject)': $($_.Exception.Message)"
            }
            if ($consecutiveFailures -ge $MaxConsecutiveFailures) {
                Write-Error ("Aborting: $consecutiveFailures consecutive failures. " +
                             'Outlook has most likely stopped responding - check it is running with a visible window, then re-run.') `
                             -ErrorAction Continue
                $aborted = $true
                break
            }
        }
    }

    if ($WhatIfPreference) { break }   # nothing actually moves, so the source never shrinks
}

Write-Host ''
Write-Host "Moved $moved item(s)." -ForegroundColor $(if ($aborted) { 'Red' } else { 'Green' })

if ($log.Count -gt 0) {
    if (-not $LogPath) {
        $LogPath = Join-Path $PSScriptRoot ('RestoreLog-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date))
    }
    $log | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to $LogPath"
}

if ($aborted) { exit 1 }
