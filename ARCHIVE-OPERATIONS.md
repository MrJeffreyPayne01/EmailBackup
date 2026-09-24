# Archive Operations Runbook

This runbook is for future sessions working on the EmailBackup repository. It records the verified machine state, the safe order of operations, and the failure modes encountered during the 2026-09-22/23 cleanup.

## Current state

As of 2026-09-24:

- Classic Outlook is installed and is the required client.
- New Outlook is pinned off for the current user with `UseNewOutlook=0` and `HideNewOutlookToggle=1`.
- Gmail is configured as IMAP, so Outlook has no local account file path for the Gmail store.
- Archive PSTs are in `E:\PSTArchiveFiles`.
- Outlook is currently closed and all PSTs were verified free for exclusive read/write access.
- The last verified Gmail Trash total was 22 items. Normal mail was archived; remaining items include non-mail or unreadable Outlook items such as `Recall: Property Manager`.

PST sizes at the last verification:

| PST | Size |
| --- | ---: |
| `Pre2020Mail.pst` | 1,573.5 MB |
| `2020Mail.pst` | 612.3 MB |
| `2021Mail.pst` | 985.0 MB |
| `2022Mail.pst` | 1,416.6 MB |
| `2023Mail.pst` | 1,298.9 MB |
| `2024Mail.pst` | 1,411.7 MB |
| `2025Mail.pst` | 1,372.0 MB |

## Before starting

1. Make a fresh copy of every PST. Never begin a large transfer with the only backup mounted in Outlook.
2. Start classic Outlook directly and wait until the Gmail folders finish synchronizing.
3. Confirm it has a visible window:

   ```powershell
   Get-Process OUTLOOK | Select-Object Id, Responding, MainWindowTitle
   ```

   An empty `MainWindowTitle` means a headless COM-started instance. Close it and launch:

   ```powershell
   Start-Process 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'
   ```

4. Run the read-only pre-flight:

   ```powershell
   .\Test-OutlookAutomation.ps1
   ```

5. Use the store-qualified Gmail paths. The default Outlook store is `My Outlook Data File(1)`, so a bare `Inbox` is the wrong source on this machine:

   ```text
   mrjeffreypayne01@gmail.com\[Gmail]\Trash
   mrjeffreypayne01@gmail.com\Inbox
   ```

## Archive commands

Use both bounds for an individual year. `-SinceYear` is inclusive and `-BeforeYear` is exclusive:

```powershell
# Pre-2020
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\Pre2020Mail.pst' -TargetFolderName 'Pre-2020' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -BeforeYear 2020 -MaxItems 2000 -Confirm:$false

# 2020
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2020Mail.pst' -TargetFolderName '2020' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2020 -BeforeYear 2021 -MaxItems 2000 -Confirm:$false

# 2021
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2021Mail.pst' -TargetFolderName '2021' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2021 -BeforeYear 2022 -MaxItems 2000 -Confirm:$false

# 2022
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2022Mail.pst' -TargetFolderName '2022' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2022 -BeforeYear 2023 -MaxItems 2000 -Confirm:$false

# 2023
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2023Mail.pst' -TargetFolderName '2023' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2023 -BeforeYear 2024 -MaxItems 2000 -Confirm:$false

# 2024
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2024Mail.pst' -TargetFolderName '2024' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2024 -BeforeYear 2025 -MaxItems 2000 -Confirm:$false

# 2025, if intentionally archiving it rather than keeping it in Gmail
.\Move-OutlookMailToArchive.ps1 -PstPath 'E:\PSTArchiveFiles\2025Mail.pst' -TargetFolderName '2025' -SourceFolderPath 'mrjeffreypayne01@gmail.com\[Gmail]\Trash' -SinceYear 2025 -BeforeYear 2026 -MaxItems 2000 -Confirm:$false
```

Repeat a bounded command until its year count reaches zero. The archive script is resumable: completed moves remain completed, and a new run only sees items still in the source folder.

## Batch sizing

- Local PST moves reached roughly 9-12 items/sec in the final passes.
- IMAP-to-IMAP moves were much slower, roughly 1 item/sec, so restore current mail to the Inbox only when that is truly intended.
- `MaxItems 2000` worked well for the stable later passes. Use `500` after an Outlook crash or when an item repeatedly fails.
- Do not run multiple Outlook COM archive commands concurrently.
- Keep the computer awake and leave Outlook visible during a long run.

## Failure handling

### RPC errors

`0x800706BA` and `0x800706BE` mean Outlook stopped responding or crashed. The archive script aborts after 15 consecutive failures. Do not keep retrying against the same dead process:

1. Stop the run.
2. Check `Get-Process OUTLOOK`.
3. If the process is headless, force-close it.
4. Launch classic Outlook visibly and wait for synchronization.
5. Re-run the same bounded command.

### Repeated unreadable item

`Recall: Property Manager` repeatedly fails with `Could not open the item`. It is not a normal mail item and should remain in Trash unless handled manually in Gmail. The script skips it and continues.

### Stale wrapper exit code

A PowerShell wrapper can report a stale `$LASTEXITCODE` even when the script moved its full batch successfully. Trust the script's `Moved N item(s)` line and verify source/destination counts directly; do not use a stale exit code alone to decide whether to rerun.

### Year-boundary dates

The DASL `datereceived` property is UTC, while Outlook displays local `ReceivedTime`. The archive script converts local New Year's midnight to UTC. Do not replace its generated filters with hard-coded UTC dates without preserving that conversion.

## After archiving

1. Let Outlook finish Gmail synchronization.
2. Close Outlook completely.
3. If a PST remains locked, check for hidden `OUTLOOK.EXE` and stop it.
4. Windows Search can also lock a PST. Stop only `SearchIndexer` temporarily if needed, then retest.
5. Verify each PST with an exclusive open:

   ```powershell
   Get-ChildItem 'E:\PSTArchiveFiles' -Filter '*.pst' | ForEach-Object {
       try {
           $s = [IO.File]::Open($_.FullName, 'Open', 'ReadWrite', 'None')
           $s.Close()
           "FREE   $($_.Name)"
       } catch {
           "LOCKED $($_.Name)"
       }
   }
   ```

6. Copy all free PSTs to the backup location before reopening Outlook.
7. Keep generated CSV logs in `log/`; they contain personal mail metadata and are gitignored.

## Git update

Only commit scripts and documentation. Never commit PSTs, CSV logs, or transcript files:

```powershell
git status --short
git add README.md ARCHIVE-OPERATIONS.md *.ps1
git commit -m "Describe the archive session"
git push origin main
```
