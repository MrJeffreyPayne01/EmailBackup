# EmailBackup

PowerShell automation for archiving Outlook mail into year-based PST files.

These scripts drive the **classic** Outlook desktop client over COM to move messages older than a cut-off year out of a live mailbox and into a folder inside an archive PST.

| Script | Purpose |
| --- | --- |
| [`Test-OutlookAutomation.ps1`](Test-OutlookAutomation.ps1) | Read-only pre-flight check. **Run this first.** |
| [`Move-OutlookMailToArchive.ps1`](Move-OutlookMailToArchive.ps1) | Does the archiving. One source folder, one target PST. |
| [`Invoke-YearlyArchive.ps1`](Invoke-YearlyArchive.ps1) | Runs every year pass in the correct order. |

---

## Quick start

```powershell
# 1. Pre-flight - checks everything, changes nothing
.\Test-OutlookAutomation.ps1

# 2. Dry run, 50 oldest messages
.\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItems 50 -WhatIf

# 3. Same run for real
.\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItems 50 -Confirm:$false
```

---

## Why classic Outlook

This machine has both Outlook clients installed:

| Client | Version | Usable here |
| --- | --- | --- |
| Classic Outlook (Microsoft 365) | 16.0.20326.20144 | Yes |
| New Outlook (Store app) | 1.2026.908.200 | No |

The new Outlook client exposes **no COM automation API** and **cannot open PST files** at all. Every approach to scripted PST archiving therefore requires classic Outlook, which is present at:

```
C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
```

The COM class `Outlook.Application.16` is registered, so the script can attach to it. The script verifies this on startup and aborts with a clear message if classic Outlook is missing.

**Keep new Outlook closed while the script runs.** Running both clients against the same profile at once can cause sync conflicts.

---

## Know your store layout before you start

This is the single easiest thing to get wrong. On this machine the **default delivery store is `My Outlook Data File(1)`**, not the Gmail account:

```
My Outlook Data File(1)      C:\Users\...\Outlook Files\My Outlook Data File(1).pst
mrjeffreypayne01@gmail.com   (no file - IMAP)
Outlook Data File            E:\PSTArchiveFiles\Pre2020Mail.pst
```

A bare `-SourceFolderPath 'Inbox'` resolves against the **default** store, which here is the near-empty local PST. The script would report zero candidates and appear to do nothing. Always pass the store-qualified path:

```powershell
-SourceFolderPath 'mrjeffreypayne01@gmail.com\Inbox'
```

`Test-OutlookAutomation.ps1` prints your stores and flags the default one. `Move-OutlookMailToArchive.ps1 -ListFolders` prints the full folder tree.

---

## How the archiving works

1. **Verify classic Outlook** — locates `OUTLOOK.EXE`, ignoring copies under `\Updates\`.
2. **Connect** — attaches to an already-running Outlook via `GetActiveObject`, otherwise starts a new instance, then logs on to the MAPI namespace.
3. **Mount the PST** — scans mounted stores for a matching `Store.FilePath`; if absent, calls `AddStoreEx` with `olStoreUnicode` (3). Outlook creates the file if it does not exist.
4. **Create the destination folder** — adds the target folder under the PST root if it is not already there.
5. **Select messages** — applies a DASL restriction on the source folder:

   ```sql
   @SQL="urn:schemas:httpmail:datereceived" < '2020-01-01 00:00'
   ```

   Filtering server-side via `Items.Restrict` is far faster than iterating every item, and the ISO-style DASL date avoids the locale-dependent parsing that `[ReceivedTime]` jet queries suffer from. Only items whose `Class` is `olMail` (43) are kept, so meeting responses and delivery reports are left alone. Results are sorted **oldest first**, so a capped run takes the genuinely oldest mail.
6. **Collect identifiers first** — the matching messages are captured as `EntryID` / `StoreID` pairs *before* any move happens. This is the single most important detail in the script: moving an item mutates the `Items` collection it came from, so walking that collection live causes Outlook to silently skip roughly every second message. Each item is re-fetched with `GetItemFromID` immediately before its move.
7. **Move** — each message is moved to the target folder, with per-item progress and a warning (not a halt) on individual failures.
8. **Log** — every moved message is written to a timestamped CSV.

### Folder mirroring

With `-IncludeSubfolders`, the source folder tree is recreated under the target folder in the PST, so `Inbox\Receipts\2018` lands in `Pre-2020\Receipts\2018` rather than being flattened.

---

## Prerequisites

- Classic Outlook installed and configured with the mail account
- The archive directory (default `E:\PSTArchiveFiles`) exists and is writable
- PowerShell running **as the same user, at the same elevation** as Outlook

Works under both Windows PowerShell 5.1 and PowerShell 7.

Outlook COM is not accessible across integrity levels. See [Troubleshooting](#troubleshooting) if activation fails.

If script execution is blocked:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Usage

### 1. Pre-flight

```powershell
.\Test-OutlookAutomation.ps1
```

Checks the Outlook install, COM registration, elevation, COM activation, your mounted stores and the archive directory. Changes nothing.

### 2. Discover your folder paths

IMAP and Gmail folder names are frequently not what you expect.

```powershell
.\Move-OutlookMailToArchive.ps1 -ListFolders
```

### 3. Dry run - moves nothing

```powershell
.\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItems 50 -WhatIf
```

### 4. Small live batch

```powershell
.\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItems 50 -Confirm:$false
```

Verify the results in Outlook before continuing.

### 5. Full run

```powershell
.\Move-OutlookMailToArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -IncludeSubfolders -Confirm:$false
```

---

## Parameters - `Move-OutlookMailToArchive.ps1`

| Parameter | Default | Description |
| --- | --- | --- |
| `-PstPath` | `E:\PSTArchiveFiles\Pre2020Mail.pst` | Archive PST file. Created by Outlook if missing. |
| `-TargetFolderName` | `Pre-2020` | Folder inside the PST that receives the mail. |
| `-SourceFolderPath` | `Inbox` | Source path. **Store-qualify this** - see [Know your store layout](#know-your-store-layout-before-you-start). |
| `-BeforeYear` | `2020` | Archives items received strictly before 1 January of this year. |
| `-IncludeSubfolders` | off | Recurse, mirroring the folder tree under the target. |
| `-MaxItems` | unlimited | Cap on items **considered**, oldest first. Bounds the scan itself, so it also limits a `-WhatIf` run. |
| `-LogPath` | timestamped CSV beside the script | Destination for the move log. |
| `-ListFolders` | off | Print the folder tree and exit without moving anything. |
| `-WhatIf` / `-Confirm` | — | Standard PowerShell safety switches. |

### Path resolution

If the first segment of `-SourceFolderPath` matches a top-level store name, the path is anchored to that store. Otherwise it resolves relative to the default account's mailbox root. Both of these work:

```
Inbox\Receipts
me@gmail.com\Inbox\Receipts
```

---

## Archiving into the year PSTs

`E:\PSTArchiveFiles` holds `Pre2020Mail.pst`, `2020Mail.pst`, `2021Mail.pst`, `2022Mail.pst` and `2023Mail.pst`.

`-BeforeYear` is an **exclusive upper bound with no lower bound**, so each pass takes everything the previous pass left behind. Pass order is therefore load-bearing: oldest first. Running the 2023 pass first would sweep decades of mail into `2023Mail.pst`.

[`Invoke-YearlyArchive.ps1`](Invoke-YearlyArchive.ps1) enforces the ordering so it cannot be got wrong by hand:

```powershell
# Print the plan without touching anything
.\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -WhatIf

# Rehearse with 25 items per pass
.\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -MaxItemsPerPass 25

# Full run
.\Invoke-YearlyArchive.ps1 -SourceFolderPath 'me@gmail.com\Inbox' -Confirm:$false
```

Its passes are:

| Pass | PST | Takes mail received before |
| --- | --- | --- |
| Pre-2020 | `Pre2020Mail.pst` | 2020-01-01 |
| 2020 | `2020Mail.pst` | 2021-01-01 |
| 2021 | `2021Mail.pst` | 2022-01-01 |
| 2022 | `2022Mail.pst` | 2023-01-01 |
| 2023 | `2023Mail.pst` | 2024-01-01 |

Use `-ThroughYear` to change the last year processed.

---

## Troubleshooting

### `0x80080005` CO_E_SERVER_EXEC_FAILURE - "Server execution failed"

Almost always an **elevation mismatch**. Outlook is a single-instance COM server: a non-elevated script cannot bind to an Outlook running as administrator, and COM cannot start a second instance alongside it.

Fix:

1. Close **every** `OUTLOOK.EXE`.
2. Relaunch Outlook **not** as administrator - shortcut > Properties > Advanced > untick **Run as administrator**.
3. Re-run `Test-OutlookAutomation.ps1`.

Do not work around this by elevating PowerShell; that inverts the problem the next time Outlook starts normally. A stale or hung Outlook process produces the same error, and the same restart clears it.

### "Value does not fall within the expected range" when creating a folder

`OlFolderType` has no "mail" member, so `Folders.Add($name, 0)` is invalid. The type argument must be omitted, letting the new folder inherit `IPF.Note` from its parent.

### The script reports 0 candidates

You are almost certainly pointing at the wrong store. See [Know your store layout](#know-your-store-layout-before-you-start).

### A `-WhatIf` run takes a very long time

Add `-MaxItems`. It bounds the scan itself, not just the moves, so the script stops reading COM properties once it has enough candidates.

### Execution policy blocks the script

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### "The process cannot access the file ... because it is being used by another process"

Outlook holds the PST files open. See [Unlocking the PSTs for backup](#unlocking-the-psts-for-backup).

---

## Unlocking the PSTs for backup

You cannot copy a PST while Outlook has it open. Two separate things cause the lock, and releasing only the first is not enough:

1. **The store is mounted in the profile.** `Namespace.RemoveStore($rootFolder)` detaches it. This does **not** delete the file, and the archive script re-mounts it automatically on its next run.
2. **Outlook keeps the file handle until the process exits.** Dismounting alone leaves the file locked. Outlook may also hold handles on PSTs that are not currently mounted, if it opened them earlier in the session - so expect *every* PST in the directory to be locked, not just the one you were archiving to.

A further trap: Outlook will ignore both `CloseMainWindow()` and its COM `Quit()` method while **any** process still holds a COM reference to it. An RCW keeps the server process alive. If you have PowerShell sessions with `$outlook` or `$namespace` variables still in scope, close those sessions first.

### Procedure

```powershell
# 1. Dismount any archive stores from the profile (does not delete anything)
$ns = (New-Object -ComObject Outlook.Application).GetNamespace('MAPI')
for ($i = $ns.Folders.Count; $i -ge 1; $i--) {
    $f = $ns.Folders.Item($i)
    $p = $null; try { $p = $f.Store.FilePath } catch { }
    if ($p -and $p -like 'E:\PSTArchiveFiles\*') { $ns.RemoveStore($f) }
}

# 2. Close every PowerShell session still holding Outlook COM references,
#    then make sure Outlook has actually exited
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
(Get-Process OUTLOOK -ErrorAction SilentlyContinue | Measure-Object).Count   # must be 0

# 3. Verify every PST is free before copying
Get-ChildItem 'E:\PSTArchiveFiles' -Filter *.pst | ForEach-Object {
    try {
        $s = [IO.File]::Open($_.FullName, 'Open', 'ReadWrite', 'None'); $s.Close()
        '  FREE   {0} ({1:N1} MB)' -f $_.Name, ($_.Length / 1MB)
    } catch {
        '  LOCKED {0}' -f $_.Name
    }
}
```

Step 3 is the part worth keeping: an exclusive `[IO.File]::Open` with `None` sharing is the only reliable proof the file is copyable. A plain `Test-Path` tells you nothing about locks.

Force-terminating Outlook is safe when the instance was started headlessly by the script. If you have Outlook open interactively with unsent drafts, close it normally instead.

### While copying

Let the copy finish completely before reopening Outlook. Relaunching it mid-copy re-acquires the handles and the copy fails partway through. Reliable sequence every time:

> close Outlook fully -> verify the process count is 0 -> copy -> reopen

---

## Cautions

**Gmail over IMAP.** Moving a message out of an IMAP folder into a local PST deletes it from Google's servers on the next sync. The PST becomes your only copy. Back up `E:\PSTArchiveFiles` before a full run.

**Prefer `Inbox` over `[Gmail]\All Mail`.** `All Mail` contains every message in the account, including copies already filed elsewhere, so archiving it will download and move far more than intended.

**Moves are not bulk-undoable.** Recovering means dragging messages back folder by folder. Always `-WhatIf` first.

**PST size limits.** Unicode PSTs default to a 50 GB ceiling. Splitting by year keeps each file well clear of it.

**Let Outlook finish syncing** before archiving, so the restriction sees the true state of the mailbox.

**Back up the PSTs regularly.** They are the only copy of archived mail once IMAP sync removes it from the server. See [Unlocking the PSTs for backup](#unlocking-the-psts-for-backup) for how to release the file locks first.

---

## Logs and privacy

Each run writes `ArchiveLog-yyyyMMdd-HHmmss.csv` containing the subject, sender, received time and source folder of every moved message.

These logs contain **personal mail metadata** and are excluded by [`.gitignore`](.gitignore), along with `*.pst`, `*.ost` and `*.msg`. Do not commit them. If you need to share one, strip the subject and sender columns first.

---

## Repository contents

| File | Purpose |
| --- | --- |
| `Test-OutlookAutomation.ps1` | Read-only pre-flight diagnostics |
| `Move-OutlookMailToArchive.ps1` | The archiving script |
| `Invoke-YearlyArchive.ps1` | Ordered multi-year orchestration |
| `README.md` | This document |
| `.gitignore` | Excludes archive logs, PST/OST/MSG files, editor noise |
