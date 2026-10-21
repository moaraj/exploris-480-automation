# Exploris 480 automation

PowerShell and Python scripts for transferring instrument files, collecting diagnostic logs, mirroring experiment folders, and preparing draft Xcalibur worklists. This repository starts with four scripts read from the acquisition PC on 2026-10-20 Pacific time.

## Status

Treat this as a source-control baseline, not a validated deployment package. The import preserves the PC scripts' executable behavior; it removes a workstation name and project-specific examples from comments.

**Do not run the watchdog unchanged on an acquisition PC that must not perform conversion.** The imported copy has `$Convert = $true`. Before deployment, review [known gaps and pending updates](docs/STATUS.md).

We have not established that these files match the processes running on the PC. A folder-listing error also prevented a complete inventory of `D:\Scripts`; additional scripts and external automation definitions may remain outside this repository.

## Included scripts

| File | Purpose | Important boundary |
| --- | --- | --- |
| `scripts/Exploris Raw Data Watchdog (Direct to Google Drive).ps1` | Watch instrument files, upload through rclone, reconcile folders by experiment ID, and optionally convert RAW to mzML | Conversion remains enabled in this imported copy; startup can stop other rclone processes using the same configuration |
| `scripts/Exploris-Log-Uploader.ps1` | Stage diagnostic logs and Windows events, then copy them to Drive | Review collection paths before using it; logs can contain sensitive information |
| `scripts/mirror_experiment_folders.ps1` | Create local experiment folders and copy top-level Excel workbooks from Drive | This script makes one pass; it does not install a schedule |
| `scripts/exploris_worklist_watcher.py` | Poll Drive for workbooks and write draft worklists and metadata | Uses older hash-based output names and blank settings; review before acquisition |

The repository does not contain RAW/mzML data, workbooks, sample sequences, instrument method files, reports, logs, rclone credentials, or vendor executables. References to local credential-file paths in source code do not include the credential files themselves.

## Setup and review

Use a separate checkout for review rather than cloning over a live `D:\Scripts` directory. Read [setup and safe deployment](docs/SETUP.md), inspect the differences, and test with a scratch experiment before replacing a production script.

For Python dependency installation in a separate test environment:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

These commands only prepare the Python environment. They do not start a watcher or configure Drive access.

## Validation

GitHub Actions checks Python syntax and parses PowerShell files with Windows PowerShell. These checks do not execute the instrument scripts, access Drive, or establish acquisition safety.

For local syntax checks:

```powershell
py -m py_compile scripts\exploris_worklist_watcher.py
powershell.exe -NoProfile -File tests\check-powershell-syntax.ps1
```

## Updates

Record changes in [CHANGELOG.md](CHANGELOG.md) using dated two-digit revisions. Keep proposed behavior changes separate from the imported baseline, and record deployment verification separately from a Git commit.

No license grant accompanies this initial import. Confirm ownership and choose a license before describing the repository as open-source.
