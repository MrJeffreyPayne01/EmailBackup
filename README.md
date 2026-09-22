# EmailBackup

PowerShell automation for archiving Outlook mail into year-based PST files.

[`Move-OutlookMailToArchive.ps1`](Move-OutlookMailToArchive.ps1) drives the **classic** Outlook desktop client over COM to move messages older than a cut-off year out of a live mailbox and into a folder inside an archive PST.

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

## How the script works

1. **Verify classic Outlook** — locates `OUTLOOK.EXE`, ignoring copies under `\Updates\`.
2. **Connect** — attaches to an already-running Outlook via `GetActiveObject`, otherwise starts a new instance, then logs on to the MAPI namespace.
3. **Mount the PST** — scans mounted stores for a matching `Store.FilePath`; if absent, calls `AddStoreEx` with `olStoreUnicode` (3). Outlook creates the file if it does not exist.
4. **Create the destination folder** — adds the target folder under the PST root if it is not already there.
5. **Select messages** — applies a DASL restriction on the source folder:

   ```sql
   @SQL="urn:schemas:httpmail:datereceived" < '2020-01-01 00:00'
   ```

   Filtering server-side via `Items.Restrict` is far faster than iterating every item, and the ISO-style DASL date avoids the locale-dependent parsing that `[ReceivedTime]` jet queries suffer from. Only items whose `Class` is `olMail` (43) are kept, so meeting responses and delivery reports are left alone.
6. **Collect identifiers first** — the matching messages are captured as `EntryID` / `StoreID` pairs *before* any move happens. This is the single most important detail in the script: moving an item mutates the `Items` collection it came from, so walking that collection live causes Outlook to silently skip roughly every second message. Each item is re-fetched with `GetItemFromID` immediately before its move.
7. **Move** — each message is moved to the target folder, with per-item progress and a warning (not a halt) on individual failures.
8. **Log** — every moved message is written to a timestamped CSV.

### Folder mirroring

With `-IncludeSubfolders`, the source folder tree is recreated under the target folder in the PST, so `Inbox\Receipts\2018` lands in `Pre-2020\Receipts\2018` rather than being flattened.

---

## Prerequisites

- Classic Outlook installed and configured with the mail account
- The archive directory (default `E:\PSTArchiveFiles`) exists and is writable
- PowerShell running **as the same user** that owns the Outlook profile

Outlook COM is not accessible across security contexts, so do **not** run this from an elevated prompt if Outlook itself runs unelevated — the script will fail to attach.

If script execution is blocked:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Usage

### 1. Discover your folder paths

IMAP and Gmail folder names are frequently not what you expect.

```powershell
.\Move-OutlookMailToArchive.ps1 -ListFolders
```

### 2. Dry run — moves nothing

```powershell
.\Move-OutlookMailToArchive.ps1 -WhatIf -Verbose
```

### 3. Small live batch

```powershell
.\Move-OutlookMailToArchive.ps1 -MaxItems 25 -Confirm:$false
```

Verify the results in Outlook before continuing.

### 4. Full run

```powershell
.\Move-OutlookMailToArchive.ps1 -IncludeSubfolders -Confirm:$false
```

---

## Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-PstPath` | `E:\PSTArchiveFiles\Pre2020Mail.pst` | Archive PST file. Created by Outlook if missing. |
| `-TargetFolderName` | `Pre-2020` | Folder inside the PST that receives the mail. |
| `-SourceFolderPath` | `Inbox` | Backslash-delimited source path, e.g. `Inbox\Receipts` or `me@gmail.com\[Gmail]\All Mail`. |
| `-BeforeYear` | `2020` | Archives items received strictly before 1 January of this year. |
| `-IncludeSubfolders` | off | Recurse, mirroring the folder tree under the target. |
| `-MaxItems` | unlimited | Stop after this many moves. Use for a cautious first run. |
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

`E:\PSTArchiveFiles` already holds `Pre2020Mail.pst`, `2020Mail.pst`, `2021Mail.pst`, `2022Mail.pst` and `2023Mail.pst`. The same script fills each one.

`-BeforeYear` is an **exclusive upper bound only** — there is no lower bound. Run the passes oldest-first so each pass only sees what the previous one left behind:

```powershell
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\Pre2020Mail.pst' -TargetFolderName 'Pre-2020' -BeforeYear 2020
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2020Mail.pst'    -TargetFolderName '2020'     -BeforeYear 2021
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2021Mail.pst'    -TargetFolderName '2021'     -BeforeYear 2022
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2022Mail.pst'    -TargetFolderName '2022'     -BeforeYear 2023
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2023Mail.pst'    -TargetFolderName '2023'     -BeforeYear 2024
```

Running them out of order will pull older mail into a newer PST.

---

## Cautions

**Gmail over IMAP.** Moving a message out of an IMAP folder into a local PST deletes it from Google's servers on the next sync. The PST becomes your only copy. Back up `E:\PSTArchiveFiles` before a full run.

**Prefer `Inbox` over `[Gmail]\All Mail`.** `All Mail` contains every message in the account, including copies already filed elsewhere, so archiving it will download and move far more than intended.

**Moves are not bulk-undoable.** Recovering means dragging messages back folder by folder. Always `-WhatIf` first.

**PST size limits.** Unicode PSTs default to a 50 GB ceiling. Splitting by year keeps each file well clear of it.

**Let Outlook finish syncing** before archiving, so the restriction sees the true state of the mailbox.

---

## Logs and privacy

Each run writes `ArchiveLog-yyyyMMdd-HHmmss.csv` containing the subject, sender, received time and source folder of every moved message.

These logs contain **personal mail metadata** and are excluded by [`.gitignore`](.gitignore). Do not commit them.

---

## Repository contents

| File | Purpose |
| --- | --- |
| `Move-OutlookMailToArchive.ps1` | The archiving script |
| `README.md` | This document |
| `.gitignore` | Excludes archive logs and PST files |
