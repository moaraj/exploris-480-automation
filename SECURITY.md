# Public repository safety

Keep credentials, instrument data, sample metadata, proprietary sequences, diagnostic logs, and patient information out of this public repository. Inspect staged content before each push; ignore rules alone cannot establish that a commit is safe.

Do not attach raw production logs to public issues. Report suspected secrets privately to the repository owner and revoke exposed credentials before discussing a redacted incident.

Do not run imported scripts against production paths without reviewing their settings and side effects. Syntax checks do not validate instrument operation, data completeness, or scientific results.
