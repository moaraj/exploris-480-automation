# Import status and pending updates

Revision: 2026-10-20 r01. This document separates code inspected during the import from requirements that still need implementation or verification.

## Import coverage

We read four named files from `D:\Scripts` through direct file access. We verified their UTF-8 bytes against the SHA-256 values returned by the device before making public-comment redactions.

| Source filename | Original SHA-256 |
| --- | --- |
| `Exploris Raw Data Watchdog (Direct to Google Drive).ps1` | `fcb2c0d0a7e09a6336f37752852013350be1b1563b7d9ded03238c8dcccdb0b7` |
| `Exploris-Log-Uploader.ps1` | `d2d3291b6fc94d9a6247e7204171430c3961f532773d1986f451169c435749ac` |
| `mirror_experiment_folders.ps1` | `513bd91bd9b14dce2b738c45994d212295260652da413b429e4c3e0abd55c1cb` |
| `exploris_worklist_watcher.py` | `1c0a97ff56158fa74afbfb657535b035ee847ddc0b326270761d8b9a8707bc17` |

The public watchdog and log-uploader hashes differ because we removed identifying comments. We did not alter executable statements in any imported script.

We could not enumerate the directory: the device's PowerShell tool rejected workspace attachment as not writable by the host user. Direct reads succeeded, but this does not establish a complete backup of the folder or all earlier project artifacts.

## Code differences from current workflow requirements

- **Conversion:** The watchdog has `$Convert = $true` and `$PauseConvertDuringAcq = $true`. The requested acquisition-host policy is no local msConvert. Disabling conversion in a reviewed revision and verifying deployment remain pending.
- **Conversion pause:** The script checks acquisition state before starting the next conversion. It does not suspend a conversion already in progress when acquisition starts, despite stronger language in the original header.
- **Worklist names:** The Python watcher writes hash-suffixed names and generated folders. The requested output convention uses the experiment folder name plus `Exploris 480 Worklist`, with dated two-digit revisions where needed.
- **Blank settings:** The watcher prepends two 4 µL blanks at G1 and G5. It does not implement the later bracketing-blank or periodic G5 blank requirements. Confirm the intended assay-specific policy before changing generation.
- **Tray and method selection:** The watcher writes unprefixed plate wells and uses a fixed method-path default. Verify rack prefixes, method existence, blank methods, and actual executed acquisition settings before import.
- **Excel mirror:** The mirror file now exists on the PC. File presence does not establish a recurring task or a successful unattended run.
- **Folder identity:** The watchdog contains experiment-ID mapping. Duplicate-folder reconciliation and stable routing still need a real-file acceptance test.
- **Reporting and organization:** Batch reporting and results-folder organization run outside these four scripts. This repository does not yet include runnable exports of those external automations.

## Risks to review before deployment

- **Shared rclone processes:** Watchdog startup stops rclone processes whose command lines contain its configuration path. That can interrupt another script using the same configuration.
- **Remote changes:** The watchdog can mirror renames with `moveto`. `$MirrorDeletes` defaults to false, but enabling it permits deletion operations; keep it off unless the operator approves the consequences.
- **Log privacy:** The log uploader collects several text-like formats and Windows events. Review its include/exclude paths and the private Drive destination before starting it.
- **Header claims:** Original script comments describe intended operation. They do not prove resumability, complete upload integrity, conversion completeness, or acquisition safety.

## Acceptance checklist

- [ ] Enumerate the full scripts folder and reconcile remaining project-generated updates.
- [ ] Review startup tasks and running process arguments without interrupting acquisition.
- [ ] Implement upload-only acquisition-host settings in a separate reviewed change.
- [ ] Reconcile worklist filename, blank, tray, method, and repeat-injection policies.
- [ ] Test uploads and mirrors against a scratch Drive folder.
- [ ] Verify periodic execution and behavior after restart.
- [ ] Record deployment date, revision, rollback copy, and observed results.

The GitHub import does not deploy software to the PC or change scheduled tasks. Operators should complete the relevant checklist before treating a revision as production-ready.
