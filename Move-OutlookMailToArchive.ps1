<#
.SYNOPSIS
    Moves Outlook mail older than a cut-off year into a folder inside an archive PST file.

.DESCRIPTION
    Automates classic Outlook (desktop) through COM to:
      1. Attach the archive PST if it is not already mounted.
      2. Create the destination folder in that PST if it does not exist.
      3. Find mail in the source folder received before -BeforeYear.
      4. Move those items into the PST folder.

    Requires classic Outlook. The "new Outlook" client exposes no COM API and cannot
    open PST files, so it is ignored even if it is installed.

    Run with -WhatIf first. Moves are not undoable in bulk.

.PARAMETER PstPath
    Full path of the archive PST file. Created by Outlook if it does not exist.

.PARAMETER TargetFolderName
    Folder inside the PST that receives the mail.

.PARAMETER SourceFolderPath
    Backslash-delimited folder path to archive from, e.g. "Inbox",
    "Inbox\Receipts", or "me@gmail.com\[Gmail]\All Mail".
    Defaults to the default account's Inbox.

.PARAMETER BeforeYear
    Items received strictly before 1 January of this year are archived.

.PARAMETER SinceYear
    Optional lower bound. Only items received on or after 1 January of this year are
    archived. Without it -BeforeYear is open-ended and will sweep up everything older,
    which matters when the source still holds many years of mail.

.PARAMETER IncludeSubfolders
    Also process subfolders of the source, recreating the folder tree under the target.

.PARAMETER MaxItems
    Stop after considering this many items, oldest first. Bounds the scan itself, so it
    also limits a -WhatIf run. Useful for a cautious first run.

.PARAMETER LogPath
    CSV file recording every moved item. Defaults to a timestamped file beside this script.

.PARAMETER ListFolders
    Print the Outlook folder tree and exit without moving anything.

.EXAMPLE
    .\Move-OutlookMailToArchive.ps1 -ListFolders

.EXAMPLE
    .\Move-OutlookMailToArchive.ps1 -WhatIf

.EXAMPLE
    .\Move-OutlookMailToArchive.ps1 -MaxItems 50

.EXAMPLE
    .\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'Inbox' -IncludeSubfolders -Confirm:$false
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$PstPath = 'E:\PSTArchiveFiles\Pre2020Mail.pst',
    [string]$TargetFolderName = 'Pre-2020',
    [string]$SourceFolderPath = 'Inbox',
    [ValidateRange(1990, 2100)]
    [int]$BeforeYear = 2020,
    [ValidateRange(1990, 2100)]
    [int]$SinceYear,
    [switch]$IncludeSubfolders,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxItems = [int]::MaxValue,
    [string]$LogPath,
    [switch]$ListFolders
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Outlook enum values, hard-coded so the script works with late-bound COM.
$olFolderInbox  = 6
$olStoreUnicode = 3
$olMailItem     = 43   # OlObjectClass.olMail

function Assert-ClassicOutlook {
    $exe = Get-ChildItem -Path 'C:\Program Files\Microsoft Office', 'C:\Program Files (x86)\Microsoft Office' `
        -Filter 'OUTLOOK.EXE' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -notmatch '\\Updates?\\' } |
        Select-Object -First 1
    if (-not $exe) {
        throw 'Classic Outlook (OUTLOOK.EXE) was not found. The new Outlook client cannot be scripted and cannot open PST files.'
    }
    Write-Verbose "Classic Outlook: $($exe.FullName) ($($exe.VersionInfo.ProductVersion))"
}

function Get-OutlookNamespace {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Verbose 'Attached to a running Outlook instance.'
    }
    catch {
        $app = New-Object -ComObject Outlook.Application
        Write-Verbose 'Started a new Outlook instance.'
    }
    $ns = $app.GetNamespace('MAPI')
    $ns.Logon($null, $null, $false, $false)
    $ns
}

function Get-PstRootFolder {
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)][string] $Path
    )

    function Find-Store {
        for ($i = 1; $i -le $Namespace.Folders.Count; $i++) {
            $root = $Namespace.Folders.Item($i)
            $filePath = $null
            try { $filePath = $root.Store.FilePath } catch { }
            if ($filePath -and ($filePath -eq $Path)) { return $root }
        }
        return $null
    }

    $root = Find-Store
    if ($root) {
        Write-Verbose "PST already mounted as store '$($root.Name)'."
        return $root
    }

    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        throw "The folder '$dir' does not exist. Create it or point -PstPath somewhere else."
    }

    Write-Verbose "Mounting PST '$Path'."
    $Namespace.AddStoreEx($Path, $olStoreUnicode)

    $root = Find-Store
    if (-not $root) {
        throw "Outlook did not report a mounted store for '$Path'. Check the path and that the file is not open elsewhere."
    }
    $root
}

function Get-ChildFolder {
    param(
        [Parameter(Mandatory)] $Parent,
        [Parameter(Mandatory)][string] $Name
    )
    for ($i = 1; $i -le $Parent.Folders.Count; $i++) {
        $child = $Parent.Folders.Item($i)
        if ($child.Name -eq $Name) { return $child }
    }
    return $null
}

function New-ChildFolderIfMissing {
    param(
        [Parameter(Mandatory)] $Parent,
        [Parameter(Mandatory)][string] $Name
    )
    $existing = Get-ChildFolder -Parent $Parent -Name $Name
    if ($existing) { return $existing }
    Write-Verbose "Creating folder '$Name' under '$($Parent.Name)'."
    # No Type argument: OlFolderType has no "mail" member, so the folder inherits IPF.Note from its parent.
    $Parent.Folders.Add($Name)
}

function Resolve-SourceFolder {
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)][string] $Path
    )

    $segments = $Path.Split('\') | Where-Object { $_ -ne '' }
    if ($segments.Count -eq 0) { throw 'SourceFolderPath is empty.' }

    # A leading segment that matches a top-level store name anchors the path to that store.
    $current = $null
    for ($i = 1; $i -le $Namespace.Folders.Count; $i++) {
        $root = $Namespace.Folders.Item($i)
        if ($root.Name -eq $segments[0]) { $current = $root; break }
    }

    if ($current) {
        $segments = $segments | Select-Object -Skip 1
    }
    else {
        # Otherwise resolve relative to the default account's mailbox root.
        $current = $Namespace.GetDefaultFolder($olFolderInbox).Parent
    }

    foreach ($segment in $segments) {
        $next = Get-ChildFolder -Parent $current -Name $segment
        if (-not $next) {
            throw "Folder '$segment' was not found under '$($current.Name)'. Run with -ListFolders to see valid paths."
        }
        $current = $next
    }
    $current
}

function Write-FolderTree {
    param(
        [Parameter(Mandatory)] $Folder,
        [int] $Depth = 0,
        [int] $MaxDepth = 4
    )
    Write-Host ('  ' * $Depth + $Folder.Name)
    if ($Depth -ge $MaxDepth) { return }
    for ($i = 1; $i -le $Folder.Folders.Count; $i++) {
        Write-FolderTree -Folder $Folder.Folders.Item($i) -Depth ($Depth + 1) -MaxDepth $MaxDepth
    }
}

function Get-ArchiveCandidate {
    <#
        Returns EntryID/StoreID pairs rather than live COM items. Moving an item mutates
        the Items collection it came from, so the collection must not be walked while moving.
    #>
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)][int] $Year,
        [int] $Limit = [int]::MaxValue
    )

    # datereceived is stored in UTC while ReceivedTime displays local time, so convert local
    # midnight to UTC or year boundaries land a few hours out.
    $cutoff = ([datetime]::new($Year, 1, 1, 0, 0, 0, [DateTimeKind]::Local)).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')
    $clauses = @('"urn:schemas:httpmail:datereceived" < ''{0}''' -f $cutoff)
    if ($script:SinceCutoff) {
        $clauses += '"urn:schemas:httpmail:datereceived" >= ''{0}''' -f $script:SinceCutoff
    }
    $filter = '@SQL=' + ($clauses -join ' AND ')

    $items = $Folder.Items
    $items.Sort('[ReceivedTime]', $false)   # oldest first, so a capped run takes the oldest mail
    $restricted = $items.Restrict($filter)

    $storeId = $Folder.StoreID
    $results = [System.Collections.Generic.List[object]]::new()

    for ($i = 1; $i -le $restricted.Count; $i++) {
        if ($results.Count -ge $Limit) { break }
        $item = $restricted.Item($i)
        if ($item.Class -ne $olMailItem) { continue }   # skip meeting responses, reports, etc.
        $results.Add([pscustomobject]@{
            EntryID      = $item.EntryID
            StoreID      = $storeId
            Subject      = $item.Subject
            ReceivedTime = $item.ReceivedTime
            SenderName   = $item.SenderName
            SourceFolder = $Folder.FolderPath
        })
    }
    $results
}

function Invoke-FolderArchive {
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)] $SourceFolder,
        [Parameter(Mandatory)] $TargetFolder,
        [Parameter(Mandatory)][int] $Year,
        [Parameter(Mandatory)][ref] $Processed,
        [Parameter(Mandatory)][ref] $MovedCount,
        [Parameter(Mandatory)][ref] $Log
    )

    if ($SourceFolder.EntryID -eq $TargetFolder.EntryID) {
        Write-Verbose "Skipping '$($SourceFolder.FolderPath)' - it is the destination."
        return
    }

    $remaining = $MaxItems - $Processed.Value
    if ($remaining -le 0) { return }

    Write-Host "Scanning $($SourceFolder.FolderPath) ..." -ForegroundColor Cyan
    # @() guards the unroll: a returned List of 0 or 1 items would otherwise arrive as $null or a scalar.
    $candidates = @(Get-ArchiveCandidate -Folder $SourceFolder -Year $Year -Limit $remaining)
    Write-Host "  $($candidates.Count) item(s) selected (received before $Year)." -ForegroundColor Cyan

    $index = 0
    foreach ($candidate in $candidates) {
        if ($Processed.Value -ge $MaxItems) {
            Write-Warning "Reached -MaxItems limit of $MaxItems. Stopping."
            return
        }
        $Processed.Value++
        $index++

        $label = '{0:yyyy-MM-dd}  {1}' -f $candidate.ReceivedTime, $candidate.Subject
        Write-Progress -Activity "Archiving $($SourceFolder.Name)" `
            -Status $label -PercentComplete (($index / [Math]::Max($candidates.Count, 1)) * 100)

        if (-not $script:Cmdlet.ShouldProcess($label, "Move to $($TargetFolder.FolderPath)")) {
            continue
        }

        try {
            $item = $Namespace.GetItemFromID($candidate.EntryID, $candidate.StoreID)
            [void]$item.Move($TargetFolder)
            $MovedCount.Value++
            $script:ConsecutiveFailures = 0
            $Log.Value.Add($candidate)
        }
        catch {
            $script:ConsecutiveFailures++
            if ($script:ConsecutiveFailures -le 3) {
                Write-Warning "Failed to move '$($candidate.Subject)': $($_.Exception.Message)"
            }
            if ($script:ConsecutiveFailures -ge 15) {
                # A dead Outlook returns the same RPC error for every item; bail out rather than
                # logging thousands of identical warnings and reporting a clean finish.
                throw "Aborting: $($script:ConsecutiveFailures) consecutive failures. Outlook has most likely stopped responding - confirm it is running with a visible window, then re-run."
            }
        }
    }
    Write-Progress -Activity "Archiving $($SourceFolder.Name)" -Completed
}

function Invoke-RecursiveArchive {
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)] $SourceFolder,
        [Parameter(Mandatory)] $TargetFolder,
        [Parameter(Mandatory)][int] $Year,
        [Parameter(Mandatory)][ref] $Processed,
        [Parameter(Mandatory)][ref] $MovedCount,
        [Parameter(Mandatory)][ref] $Log
    )

    Invoke-FolderArchive -Namespace $Namespace -SourceFolder $SourceFolder -TargetFolder $TargetFolder `
        -Year $Year -Processed $Processed -MovedCount $MovedCount -Log $Log

    if (-not $IncludeSubfolders) { return }

    for ($i = 1; $i -le $SourceFolder.Folders.Count; $i++) {
        if ($Processed.Value -ge $MaxItems) { return }
        $childSource = $SourceFolder.Folders.Item($i)
        if ($childSource.EntryID -eq $TargetFolder.EntryID) { continue }

        $childTarget = New-ChildFolderIfMissing -Parent $TargetFolder -Name $childSource.Name
        Invoke-RecursiveArchive -Namespace $Namespace -SourceFolder $childSource -TargetFolder $childTarget `
            -Year $Year -Processed $Processed -MovedCount $MovedCount -Log $Log
    }
}

# ---------------------------------------------------------------- main

# Captured so the helper functions share this script's -WhatIf / -Confirm state.
$script:Cmdlet = $PSCmdlet
$script:ConsecutiveFailures = 0
# Resolved here: $PSBoundParameters inside a function refers to that function's own parameters.
$script:SinceCutoff = if ($PSBoundParameters.ContainsKey('SinceYear')) { ([datetime]::new($SinceYear, 1, 1, 0, 0, 0, [DateTimeKind]::Local)).ToUniversalTime().ToString('yyyy-MM-dd HH:mm') } else { $null }

Assert-ClassicOutlook
$ns = Get-OutlookNamespace

if ($ListFolders) {
    Write-Host 'Outlook folder tree:' -ForegroundColor Yellow
    for ($i = 1; $i -le $ns.Folders.Count; $i++) {
        Write-FolderTree -Folder $ns.Folders.Item($i)
    }
    return
}

$pstRoot = Get-PstRootFolder -Namespace $ns -Path $PstPath
$target  = New-ChildFolderIfMissing -Parent $pstRoot -Name $TargetFolderName
$source  = Resolve-SourceFolder -Namespace $ns -Path $SourceFolderPath

Write-Host ''
Write-Host "Source      : $($source.FolderPath)"
Write-Host "Destination : $($target.FolderPath)  ($PstPath)"
Write-Host "Cut-off     : received before $BeforeYear-01-01"
if ($script:SinceCutoff) { Write-Host "Lower bound : received on/after $SinceYear-01-01" }
Write-Host "Subfolders  : $([bool]$IncludeSubfolders)"
Write-Host ''

$moved = 0
$processed = 0
$log = [System.Collections.Generic.List[object]]::new()

Invoke-RecursiveArchive -Namespace $ns -SourceFolder $source -TargetFolder $target `
    -Year $BeforeYear -Processed ([ref]$processed) -MovedCount ([ref]$moved) -Log ([ref]$log)

Write-Host ''
Write-Host "Moved $moved item(s)." -ForegroundColor Green

if ($log.Count -gt 0) {
    if (-not $LogPath) {
        $LogPath = Join-Path $PSScriptRoot ('ArchiveLog-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date))
    }
    $log | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to $LogPath"
}
