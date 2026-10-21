# Setup and safe deployment

Review a separate checkout before touching the acquisition PC's live scripts. The initial import preserves existing behavior, including settings that require correction before production use.

## Requirements

- **PowerShell:** Use Windows PowerShell 5.1 for the imported `.ps1` scripts.
- **Python:** The watcher imports `openpyxl` and uses Windows `msvcrt` locking. Use a supported Windows Python installation and an isolated environment.
- **rclone:** Install the executable separately and configure a private Drive remote outside the checkout.
- **ProteoWizard:** Only a designated conversion host should need msConvert under the requested workflow. Do not enable conversion on the acquisition host.
- **Instrument method:** Supply an existing, qualified local method file. This repository does not include vendor methods.

## Configuration locations

The PowerShell scripts keep settings near the top of each file. Review source directories, remote destinations, credential-file paths, staging locations, and logging paths before execution.

The Python watcher accepts `--remote`, `--local-root`, `--workspace`, `--rclone`, `--config`, and `--method`. Its `--once` option performs a polling pass; it is not a dry run and can download workbooks and write output files.

Keep credentials outside the checkout. Do not copy `rclone.conf`, environment files, service-account JSON, or tokens into a commit.

## Review sequence

1. Compare the proposed revision with the live file and preserve a rollback copy outside the repository.
2. Confirm acquisition is idle before any deployment or process restart.
3. Correct conversion and worklist-policy gaps described in `STATUS.md`.
4. Point a test copy at scratch directories and a private test remote.
5. Run syntax checks before testing behavior.
6. Inspect produced worklists, uploaded files, and logs.
7. Review the exact diff before replacing a live script or changing a scheduled task.

The log uploader's `-Discover` option lists candidate logs, but the script also creates its working directories. Its `-Install` option registers and starts a scheduled task; do not use it as an inspection command.

## Future commits

Stage named source and documentation files rather than running a broad add from an operational directory. Inspect both file names and staged content before pushing:

```powershell
git status --short
git diff --cached --stat
git diff --cached
```

The `.gitignore` file reduces accidental additions but does not protect content embedded inside a script or files already tracked. Review each public change for credentials, sample identities, proprietary sequences, and private endpoints.
