"""Config file management for outlook-mcp."""

import os
import stat
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field, field_validator

from outlook_mcp.permissions import VALID_CATEGORIES

DEFAULT_TENANT_ID = "consumers"
DEFAULT_CONFIG_DIR = os.path.expanduser("~/.outlook-mcp")
CONFIG_DIR_ENV = "OUTLOOK_MCP_CONFIG_DIR"


def get_config_dir() -> str:
    """Return the config directory, overridable via ``OUTLOOK_MCP_CONFIG_DIR``.

    Pointing a second server instance at its own directory gives it its own
    config.json and AuthenticationRecord — the way to serve several accounts.
    """
    return os.path.expanduser(os.environ.get(CONFIG_DIR_ENV) or DEFAULT_CONFIG_DIR)


class AccountConfig(BaseModel):
    """Configuration for a single account."""

    name: str
    client_id: str
    tenant_id: str = DEFAULT_TENANT_ID


class Config(BaseModel):
    """Outlook MCP server configuration."""

    client_id: str | None = Field(default=None, description="Azure AD app client ID (BYOID)")
    tenant_id: str = Field(default=DEFAULT_TENANT_ID)
    read_only: bool = Field(default=False)
    allow_categories: list[str] = Field(
        default_factory=list,
        description=(
            "Optional whitelist of write-tool categories. Empty list = fully open "
            "(all writes allowed when read_only=False). Non-empty = only the listed "
            "categories are permitted."
        ),
    )
    download_dir: str = Field(
        default="~/.outlook-mcp/downloads",
        description=(
            "Only directory outlook_download_attachment may write into. Attachment "
            "bytes are attacker-controlled, so writes are confined here rather than "
            "allowed anywhere on the host."
        ),
    )
    attachment_source_dirs: list[str] = Field(
        default_factory=list,
        description=(
            "Directories outlook_send_with_attachments / outlook_attach_to_draft may "
            "read files from. Empty list = no file may be attached (fail closed). "
            "Note this is the inverse of allow_categories, where empty means open: "
            "an empty allowlist here must not grant read access to the whole host."
        ),
    )
    graph_scopes: list[str] | None = Field(
        default=None,
        description=(
            "Override the delegated Graph scopes requested at token acquisition. "
            "Defaults to the read-write set matching the app registration the "
            "README documents. Set this only if you registered an Azure app "
            "granting a narrower set (e.g. read-only permissions) — the scopes "
            "requested must be ones your app has actually been consented, or "
            "token acquisition falls into an interactive flow."
        ),
    )
    allow_unencrypted_token_cache: bool = Field(
        default=False,
        description=(
            "Permit azure-identity to write the OAuth token cache to a plaintext "
            "file when no OS keyring is reachable (Linux without libsecret). Off by "
            "default: a plaintext refresh token is a long-lived mailbox credential."
        ),
    )
    timezone: str = Field(default="UTC", description="IANA timezone for relative date computations")
    accounts: list[AccountConfig] = Field(default_factory=list)
    default_account: str | None = Field(default=None)

    @field_validator("allow_categories")
    @classmethod
    def _validate_allow_categories(cls, value: list[str]) -> list[str]:
        """Reject unknown category names at config load time."""
        unknown = [c for c in value if c not in VALID_CATEGORIES]
        if unknown:
            valid_list = ", ".join(sorted(VALID_CATEGORIES))
            raise ValueError(
                f"Unknown permission categories: {unknown}. Valid categories: {valid_list}"
            )
        return value


def _ensure_dir(dir_path: str) -> Path:
    """Create config directory with 0700 permissions."""
    path = Path(dir_path)
    path.mkdir(parents=True, exist_ok=True)
    path.chmod(0o700)
    return path


def _atomic_write(file_path: Path, data: str) -> None:
    """Write file atomically with fsync, set 0600 permissions."""
    dir_path = file_path.parent
    fd, tmp_path = tempfile.mkstemp(dir=str(dir_path), suffix=".tmp")
    try:
        with os.fdopen(fd, "w") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp_path, stat.S_IRUSR | stat.S_IWUSR)  # 0600
        os.replace(tmp_path, str(file_path))
    except Exception:
        os.unlink(tmp_path)
        raise


def save_config(config: Config, config_dir: str | None = None) -> None:
    """Save config to disk."""
    config_dir = config_dir or get_config_dir()
    dir_path = _ensure_dir(config_dir)
    file_path = dir_path / "config.json"
    _atomic_write(file_path, config.model_dump_json(indent=2))


def load_config(config_dir: str | None = None) -> Config:
    """Load config from disk. Returns defaults if no config file exists."""
    config_dir = config_dir or get_config_dir()
    file_path = Path(config_dir) / "config.json"

    if not file_path.exists():
        return Config()

    if file_path.is_symlink():
        raise PermissionError(f"Refusing to load symlinked config: {file_path}")

    mode = file_path.stat().st_mode & 0o777
    if mode != 0o600:
        file_path.chmod(0o600)

    data = file_path.read_text()
    return Config.model_validate_json(data)
