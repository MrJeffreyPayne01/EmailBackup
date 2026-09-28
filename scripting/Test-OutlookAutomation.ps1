<#
.SYNOPSIS
    Pre-flight check for Outlook COM automation.

.DESCRIPTION
    Verifies everything Move-OutlookMailToArchive.ps1 depends on, and explains how to fix
    whatever is broken. Read-only: opens no mail and moves nothing.

    Checks performed:
      1. Classic Outlook is installed (the new Outlook client cannot be automated).
      2. The Outlook.Application COM class is registered.
      3. This shell's elevation, and whether it matches any running Outlook.
      4. COM activation actually succeeds.
      5. Mounted stores, their PST paths, and the default delivery store.
      6. The archive directory is present and writable.

.PARAMETER ArchiveDirectory
    Directory expected to hold the archive PST files.

.EXAMPLE
    .\Test-OutlookAutomation.ps1
#>
[CmdletBinding()]
param(
    [string]$ArchiveDirectory = 'E:\PSTArchiveFiles'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$script:Failures = 0

function Write-Result {
    param(
        [Parameter(Mandatory)][string] $Name,
        [Parameter(Mandatory)][ValidateSet('Pass', 'Warn', 'Fail')][string] $State,
        [string] $Detail,
        [string] $Remedy
    )
    $colour = @{ Pass = 'Green'; Warn = 'Yellow'; Fail = 'Red' }[$State]
    Write-Host ('[{0}] {1}' -f $State.ToUpper().PadRight(4), $Name) -ForegroundColor $colour
    if ($Detail) { Write-Host "       $Detail" }
    if ($Remedy) { Write-Host "       -> $Remedy" -ForegroundColor Cyan }
    if ($State -eq 'Fail') { $script:Failures++ }
}

Write-Host "`nOutlook automation pre-flight`n" -ForegroundColor White

# 1. Classic Outlook ---------------------------------------------------------
$classic = Get-ChildItem -Path 'C:\Program Files\Microsoft Office', 'C:\Program Files (x86)\Microsoft Office' `
    -Filter 'OUTLOOK.EXE' -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.DirectoryName -notmatch '\\Updates?\\' } |
    Select-Object -First 1

if ($classic) {
    Write-Result -Name 'Classic Outlook installed' -State 'Pass' `
        -Detail "$($classic.FullName) ($($classic.VersionInfo.ProductVersion))"
}
else {
    Write-Result -Name 'Classic Outlook installed' -State 'Fail' `
        -Detail 'OUTLOOK.EXE was not found.' `
        -Remedy 'Install classic Outlook. The new Outlook client has no COM API and cannot open PST files.'
}

# 2. New Outlook -------------------------------------------------------------
$new = Get-AppxPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue
if ($new) {
    Write-Result -Name 'New Outlook present' -State 'Warn' `
        -Detail "Version $($new.Version). It cannot be automated and cannot open PSTs." `
        -Remedy 'Close it while archiving to avoid sync conflicts on the same profile.'
}

# 3. COM registration --------------------------------------------------------
if (Test-Path 'HKLM:\SOFTWARE\Classes\Outlook.Application') {
    $curVer = (Get-ItemProperty 'HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)'
    Write-Result -Name 'Outlook.Application registered' -State 'Pass' -Detail $curVer
}
else {
    Write-Result -Name 'Outlook.Application registered' -State 'Fail' `
        -Detail 'COM class not found in the registry.' `
        -Remedy 'Run an Office Quick Repair.'
}

# 4. Elevation ---------------------------------------------------------------
$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
$running = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)

Write-Result -Name 'Shell elevation' -State 'Pass' -Detail ("Elevated: $elevated")

if ($running.Count -gt 0) {
    $ages = $running | ForEach-Object { '{0} (PID {1}, started {2})' -f $_.Name, $_.Id, $_.StartTime }
    Write-Result -Name 'Classic Outlook running' -State 'Pass' -Detail ($ages -join '; ')
}

# 5. COM activation ----------------------------------------------------------
$ns = $null
try {
    $app = New-Object -ComObject Outlook.Application
    $ns = $app.GetNamespace('MAPI')
    $ns.Logon($null, $null, $false, $false)
    Write-Result -Name 'COM activation' -State 'Pass' -Detail 'Connected to the MAPI namespace.'
}
catch {
    $message = $_.Exception.Message
    $remedy = 'Check that Outlook starts normally and is not showing a modal dialog.'
    if ($message -match '80080005') {
        # Outlook is a single-instance COM server: a non-elevated client cannot bind to an
        # elevated instance, and COM cannot start a second one.
        $remedy = 'Elevation mismatch. Close every OUTLOOK.EXE, then relaunch NOT as administrator ' +
                  '(shortcut: Properties > Advanced > untick "Run as administrator"), and retry.'
    }
    Write-Result -Name 'COM activation' -State 'Fail' -Detail $message -Remedy $remedy
}

# 6. Stores ------------------------------------------------------------------
if ($ns) {
    Write-Host "`nMounted stores:" -ForegroundColor White
    for ($i = 1; $i -le $ns.Folders.Count; $i++) {
        $root = $ns.Folders.Item($i)
        $path = '(no file)'
        try { if ($root.Store.FilePath) { $path = $root.Store.FilePath } } catch { }
        Write-Host ('  {0,-40} {1}' -f $root.Name, $path)
    }

    try {
        $defaultInbox = $ns.GetDefaultFolder(6)
        $defaultStore = $defaultInbox.Parent.Name
        Write-Host ''
        Write-Result -Name 'Default delivery store' -State 'Pass' -Detail $defaultStore `
            -Remedy "Bare 'Inbox' resolves here. Pass the full path if your mail lives in another store."
    }
    catch {
        Write-Result -Name 'Default delivery store' -State 'Warn' -Detail $_.Exception.Message
    }
}

# 7. Archive directory -------------------------------------------------------
if (Test-Path -LiteralPath $ArchiveDirectory) {
    $psts = Get-ChildItem -LiteralPath $ArchiveDirectory -Filter '*.pst' -ErrorAction SilentlyContinue
    $detail = if ($psts) {
        ($psts | ForEach-Object { '{0} ({1:N1} MB)' -f $_.Name, ($_.Length / 1MB) }) -join ', '
    } else { 'No PST files yet.' }
    Write-Result -Name 'Archive directory' -State 'Pass' -Detail "$ArchiveDirectory - $detail"

    try {
        $probe = Join-Path $ArchiveDirectory ('.writetest-{0}' -f ([guid]::NewGuid()))
        [IO.File]::WriteAllText($probe, 'x')
        Remove-Item -LiteralPath $probe -Force
        Write-Result -Name 'Archive directory writable' -State 'Pass'
    }
    catch {
        Write-Result -Name 'Archive directory writable' -State 'Fail' -Detail $_.Exception.Message `
            -Remedy 'Grant write permission, or choose another -PstPath.'
    }
}
else {
    Write-Result -Name 'Archive directory' -State 'Fail' -Detail "$ArchiveDirectory does not exist." `
        -Remedy 'Create it, or pass -ArchiveDirectory / -PstPath pointing elsewhere.'
}

Write-Host ''
if ($script:Failures -eq 0) {
    Write-Host 'Ready to archive.' -ForegroundColor Green
}
else {
    Write-Host "$script:Failures check(s) failed - resolve these before archiving." -ForegroundColor Red
    exit 1
}
