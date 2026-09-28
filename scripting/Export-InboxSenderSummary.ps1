<#
.SYNOPSIS
    Scans an Outlook folder and exports a per-sender message count to CSV.

.DESCRIPTION
    Automates classic Outlook (desktop) through COM to enumerate every mail item in a
    folder (the Inbox by default), grouping by sender email address. The result is a CSV
    - openable directly in Excel - sorted by message count descending, so the senders
    responsible for the most mail (frequently newsletters, notifications, or spam) sort
    to the top and are easy to triage.

    Read-only: this script only reads item properties. It moves and deletes nothing.

.PARAMETER SourceFolderPath
    Backslash-delimited folder path to scan, e.g. "Inbox" or "me@gmail.com\Inbox".
    Defaults to the default account's Inbox.

.PARAMETER IncludeSubfolders
    Also scan subfolders of the source folder.

.PARAMETER OutputPath
    CSV file to write. Defaults to a timestamped file in a "log" folder beside this script.

.EXAMPLE
    .\Export-InboxSenderSummary.ps1

.EXAMPLE
    .\Export-InboxSenderSummary.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -OutputPath 'C:\temp\senders.csv'
#>
[CmdletBinding()]
param(
    [string]$SourceFolderPath = 'Inbox',
    [switch]$IncludeSubfolders,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$olMailItem = 43   # OlObjectClass.olMail
$olFolderInbox = 6

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

function Get-ChildFolder {
    param([Parameter(Mandatory)] $Parent, [Parameter(Mandatory)][string] $Name)
    for ($i = 1; $i -le $Parent.Folders.Count; $i++) {
        $child = $Parent.Folders.Item($i)
        if ($child.Name -eq $Name) { return $child }
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
        $current = $Namespace.GetDefaultFolder($olFolderInbox).Parent
    }

    foreach ($seg in $segments) {
        $next = Get-ChildFolder -Parent $current -Name $seg
        if (-not $next) { throw "Folder '$seg' was not found under '$($current.Name)'." }
        $current = $next
    }
    $current
}

function Get-SenderAddress {
    param([Parameter(Mandatory)] $Item)

    # Exchange-delivered mail exposes an X500 DN here instead of an SMTP address;
    # resolving through the sender's Exchange user (when available) gets the real address.
    $address = $Item.SenderEmailAddress
    if ($Item.SenderEmailType -eq 'EX') {
        try {
            $exUser = $Item.Sender.GetExchangeUser()
            if ($exUser) { $address = $exUser.PrimarySmtpAddress }
        }
        catch { }
    }
    $address
}

function Add-FolderItems {
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)][switch] $Recurse,
        [Parameter(Mandatory)] $Counts,
        [Parameter(Mandatory)][ref] $Processed
    )

    Write-Host "Scanning '$($Folder.FolderPath)' ($($Folder.Items.Count) item(s))..."
    $items = $Folder.Items

    for ($i = 1; $i -le $items.Count; $i++) {
        $item = $items.Item($i)
        $Processed.Value++
        if ($Processed.Value % 500 -eq 0) {
            Write-Host "  ...$($Processed.Value) item(s) scanned"
        }

        if ($item.Class -ne $olMailItem) { continue }

        $address = Get-SenderAddress -Item $item
        if ([string]::IsNullOrWhiteSpace($address)) { $address = '(unknown)' }
        $name = $item.SenderName

        if ($Counts.ContainsKey($address)) {
            $entry = $Counts[$address]
            $entry.Count++
            if ($item.ReceivedTime -lt $entry.FirstReceived) { $entry.FirstReceived = $item.ReceivedTime }
            if ($item.ReceivedTime -gt $entry.LastReceived) { $entry.LastReceived = $item.ReceivedTime }
        }
        else {
            $Counts[$address] = [pscustomobject]@{
                SenderEmailAddress = $address
                SenderName         = $name
                Count              = 1
                FirstReceived      = $item.ReceivedTime
                LastReceived       = $item.ReceivedTime
            }
        }
    }

    if ($Recurse) {
        for ($i = 1; $i -le $Folder.Folders.Count; $i++) {
            Add-FolderItems -Folder $Folder.Folders.Item($i) -Recurse -Counts $Counts -Processed $Processed
        }
    }
}

if (-not $OutputPath) {
    $logDir = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'log'
    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }
    $OutputPath = Join-Path $logDir ('SenderSummary-{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date))
}

$ns = Get-OutlookNamespace
$folder = Resolve-Folder -Namespace $ns -Path $SourceFolderPath

$counts = @{}
$processed = 0
Add-FolderItems -Folder $folder -Recurse:$IncludeSubfolders -Counts $counts -Processed ([ref]$processed)

Write-Host ''
Write-Host "Scanned $processed item(s) across $($counts.Count) distinct sender(s)."

$counts.Values |
    Sort-Object -Property Count -Descending |
    Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Sender summary written to $OutputPath"
