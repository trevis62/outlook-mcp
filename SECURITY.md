# Security Policy

## Reporting a Vulnerability

If you discover a security vulnerability in outlook-mcp, please report it responsibly.

**Do NOT open a public GitHub issue for security vulnerabilities.**

### How to Report

Use [GitHub Security Advisories](https://github.com/mpalermiti/outlook-mcp/security/advisories/new) to privately report vulnerabilities.

### Response Timeline

- **Acknowledgment:** Within 48 hours
- **Initial assessment:** Within 7 days
- **Fix timeline:** Depends on severity, typically within 30 days

### Scope

The following are in scope for security reports:

- Token leakage or credential exposure
- Authentication bypass
- Input injection (OData, KQL, path traversal)
- Unauthorized file system access
- Supply chain vulnerabilities in dependencies
- Symlink attacks on config files

### Out of Scope

- Social engineering attacks
- Denial of service against Microsoft Graph API endpoints
- Issues in Microsoft Graph API itself
- Issues requiring physical access to the machine

### Security Design

outlook-mcp is designed with security in mind:

- **Tokens:** Stored via azure-identity's persistent cache — OS Keychain on macOS, DPAPI on Windows, and gnome-keyring on Linux *when libsecret/PyGObject is importable*. On Linux without it (common with `uv tool install`, whose isolated venv can't see system PyGObject), the cache falls back to a `0600` plaintext file under `~/.IdentityService/` and the server logs a one-time warning at startup. See the README for how to get encrypted storage on Linux.
- **Input validation:** All Graph IDs, emails, dates, KQL queries, and folder names are validated before use
- **Atomic writes:** Config files are written atomically to prevent corruption
- **Symlink rejection:** Config loader refuses symlinked files
- **No telemetry:** Zero data sent to any third party
- **No caching:** Email and calendar data is never written to disk
- **Egress restriction:** Caller-supplied delta cursors are full URLs; they are validated against `graph.microsoft.com` before a bearer token is attached, so a cursor cannot redirect credentials to another host
- **Filesystem confinement:** Attachment downloads are confined to `download_dir` and attachment sends to `attachment_source_dirs` (empty by default, i.e. fails closed), enforced on the symlink-resolved path

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest  | Yes       |
| < Latest | Best effort |
