"""Filesystem path confinement for the attachment tools.

Both attachment tools take host paths from the caller, and the caller is an
agent that routinely has attacker-authored message content in its context.
That makes the two obvious primitives dangerous:

- ``outlook_download_attachment`` writes attacker-controlled bytes to a
  caller-chosen path — an arbitrary file write (``~/.zshrc``,
  ``~/.ssh/authorized_keys``) and therefore host code execution.
- ``outlook_send_with_attachments`` reads a caller-chosen file and mails it
  anywhere — an arbitrary file read (``~/.ssh/id_rsa``, the token cache).

Both are confined here with the same technique: resolve the path with
symlinks followed, then require the *resolved* path to sit under an
allowlisted root. Resolving first is what defeats ``..`` traversal and
planted symlinks; a string prefix check on the unresolved path would not.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from outlook_mcp.config import Config


def _real(path: str | Path) -> Path:
    """Expand ``~``, make absolute, and resolve symlinks."""
    return Path(os.path.realpath(os.path.expanduser(str(path))))


def _is_within(child: Path, parent: Path) -> bool:
    """True if ``child`` is ``parent`` or sits underneath it.

    Uses path-component comparison via ``relative_to`` rather than a string
    prefix test, so ``/a/outbox-evil`` does not match a root of ``/a/outbox``.
    """
    try:
        child.relative_to(parent)
    except ValueError:
        return False
    return True


def download_root(config: Config) -> Path:
    """Return the download sandbox, creating it 0700 if absent."""
    root = Path(os.path.expanduser(config.download_dir))
    root.mkdir(parents=True, exist_ok=True)
    root.chmod(0o700)
    return _real(root)


def resolve_download_path(config: Config, save_path: str) -> Path:
    """Resolve ``save_path`` to a location inside the download sandbox.

    Relative paths resolve against the sandbox. Absolute paths are permitted
    only if they already point inside it. Parent directories are created.

    Raises:
        ValueError: if the resolved path escapes the sandbox.
    """
    root = download_root(config)

    candidate = Path(os.path.expanduser(save_path))
    if not candidate.is_absolute():
        candidate = root / candidate

    # Resolve the *parent* rather than the target: the file itself normally
    # doesn't exist yet, but every directory on the way to it does, and a
    # symlink planted among them is exactly the escape we're blocking.
    parent = _real(candidate.parent)
    if not _is_within(parent, root):
        raise ValueError(
            f"save_path resolves outside the download directory ({root}): {save_path}"
        )

    parent.mkdir(parents=True, exist_ok=True)
    return parent / candidate.name


def require_allowed_source(config: Config, path: str) -> Path:
    """Return ``path`` if it is a regular file inside an allowed source dir.

    Fails closed: an empty ``attachment_source_dirs`` allows nothing.

    Raises:
        FileNotFoundError: if the file does not exist.
        ValueError: if no directory is configured, if the resolved path is
            outside every configured directory, or if it isn't a regular file.
    """
    if not config.attachment_source_dirs:
        raise ValueError(
            "No attachment_source_dirs configured, so no file may be attached. "
            "Add the directories you want to send from to attachment_source_dirs "
            "in ~/.outlook-mcp/config.json."
        )

    resolved = _real(path)

    # Allowlist first, existence second: reversing these would make the tool
    # report whether any path on the host exists, before establishing that
    # the caller is even entitled to ask about it.
    roots = [_real(d) for d in config.attachment_source_dirs]
    if not any(_is_within(resolved, root) for root in roots):
        raise ValueError(
            f"Attachment path is not in an allowed source directory: {path}"
        )

    if not resolved.exists():
        raise FileNotFoundError(f"Attachment file not found: {path}")

    if not resolved.is_file():
        raise ValueError(f"Attachment path is not a regular file: {path}")

    return resolved
