"""OAuth2 authentication via azure-identity device code flow."""

from __future__ import annotations

import hashlib
import importlib.util
import logging
import os
import sys
from pathlib import Path

from azure.identity import (
    AuthenticationRecord,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)

from outlook_mcp.config import DEFAULT_CONFIG_DIR, Config, get_config_dir
from outlook_mcp.errors import AuthRequiredError, UnencryptedCacheError

logger = logging.getLogger(__name__)

# Process-local latch so the unencrypted-fallback warning fires at most
# once per run — _make_credential is called from both login_interactive
# and try_cached_token, often multiple times during startup.
_warned_unencrypted_fallback = False

# Scope names; graph_token_scopes() qualifies these for token acquisition.
SCOPES_READWRITE = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
]

# Not requested by default — the README's app registration grants only the
# ReadWrite variants. Provided as the list to copy into config.graph_scopes
# if you register a read-only Azure app and want a token that cannot write.
SCOPES_READONLY = [
    "Mail.Read",
    "Calendars.Read",
    "Contacts.Read",
    "Tasks.Read",
    "User.Read",
]

CACHE_NAME = "outlook-mcp"
AUTH_RECORD_FILE = "auth_record.json"

# What azure-identity itself uses on macOS: every cache, whatever its `name`,
# lives in this one Keychain item — `name` only picks the signal/lock file.
_KEYCHAIN_SERVICE = "Microsoft.Developer.IdentityService"
_KEYCHAIN_DEFAULT_ACCOUNT = "MSALCache"
_NON_CAE_SUFFIX = ".nocae"

# OSStatus codes worth naming when the macOS Keychain refuses a read/write.
# msal_extensions' KeychainError stringifies to "", so without this the
# server logs "Authentication failed: " with no reason at all.
_KEYCHAIN_STATUS_HINTS = {
    -128: "the Keychain access prompt was denied or cancelled",
    -25293: "Keychain authorization failed",
    -25300: "the token cache item does not exist yet",
    -25308: "Keychain refused to show an access prompt (errSecInteractionNotAllowed)",
}


def token_cache_name() -> str:
    """Per-instance token cache name, derived from the config directory.

    Two server instances (one per account, via OUTLOOK_MCP_CONFIG_DIR) must
    not share a cache: a sign-in or cache miss in one would otherwise disturb
    the other. The default directory keeps the historical name so existing
    installs don't have to re-authenticate.
    """
    config_dir = os.path.realpath(get_config_dir())
    if config_dir == os.path.realpath(os.path.expanduser(DEFAULT_CONFIG_DIR)):
        return CACHE_NAME
    digest = hashlib.sha256(config_dir.encode()).hexdigest()[:12]
    return f"{CACHE_NAME}-{digest}"


def _macos_token_cache(name: str):
    """Build a Keychain-backed cache in its own Keychain item for `name`.

    azure-identity hardcodes the Keychain account to "MSALCache", so on macOS
    TokenCachePersistenceOptions(name=...) cannot isolate two instances — they
    would write the same Keychain item while locking different files. We build
    the persistence ourselves and hand it over through DeviceCodeCredential's
    private ``_cache`` argument (pinned by tests/test_auth.py). The default
    cache maps to exactly the item azure-identity would have used.
    """
    import msal_extensions

    account = _KEYCHAIN_DEFAULT_ACCOUNT if name == CACHE_NAME else name
    signal_file = os.path.expanduser(os.path.join("~", ".IdentityService", name + _NON_CAE_SUFFIX))
    persistence = msal_extensions.KeychainPersistence(signal_file, _KEYCHAIN_SERVICE, account)
    return msal_extensions.PersistedTokenCache(persistence)


def describe_auth_failure(exc: BaseException) -> str:
    """Return a human-readable reason for a token-acquisition failure.

    Walks the exception chain so the underlying cause surfaces even when
    azure-identity wraps it in a message-less ClientAuthenticationError.
    """
    chain: list[BaseException] = []
    current: BaseException | None = exc
    while current is not None and all(current is not e for e in chain):
        chain.append(current)
        current = current.__cause__ or current.__context__
    for err in chain:
        status = getattr(err, "exit_status", None)
        if isinstance(status, int):
            hint = _KEYCHAIN_STATUS_HINTS.get(status, "see Apple's OSStatus codes")
            return f"macOS Keychain error {status}: {hint}"
    for err in chain:
        text = str(err).strip()
        if text and text.rstrip(":") != "Authentication failed":
            return f"{type(err).__name__}: {text}"
    return f"{type(exc).__name__} (no details)"


def _unencrypted_fallback_will_be_used() -> bool:
    """Return True if msal_extensions will fall back to plaintext caching.

    Mirrors msal_extensions' libsecret-availability check: macOS uses
    Keychain and Windows uses DPAPI, both always encrypted, so only
    Linux is at risk — and only when PyGObject/libsecret isn't
    importable in the current Python environment (the failure mode
    reported in #7 for `uv tool install`).
    """
    if sys.platform != "linux":
        return False
    return importlib.util.find_spec("gi") is None


GRAPH_RESOURCE = "https://graph.microsoft.com/"

# Token acquisition uses explicit delegated scopes, NOT `.default`.
#
# `.default` means "every permission statically configured for this app" and
# works for Entra work accounts, but on the /consumers endpoint it yields a
# token Graph rejects with 403 ErrorAccessDenied on every mailbox endpoint —
# /me succeeds, /me/messages does not. Personal accounts are this project's
# primary target, so `.default` is unusable here.
#
# Every token request in the codebase must use *this same list*: MSAL caches
# access tokens keyed by scope, so asking for a different set anywhere causes
# a cache miss and a background interactive auth the server cannot answer.
# That is why GraphClient carries its scopes and the raw-httpx paths reuse
# them instead of hardcoding their own.


def graph_token_scopes(config: Config | None = None) -> list[str]:
    """Fully-qualified Graph scopes for token acquisition.

    Deliberately independent of ``read_only``. The app registration the README
    documents grants only the ReadWrite variants, and scope matching is
    literal — requesting Mail.Read when Mail.ReadWrite was consented misses
    the MSAL cache and drops into an interactive device-code flow the server
    cannot answer. ``read_only`` is enforced by ``check_permission`` instead.

    Users who want token-level restriction register a narrower Azure app and
    name its scopes in ``config.graph_scopes``.
    """
    if config is not None and config.graph_scopes:
        return list(config.graph_scopes)
    return [f"{GRAPH_RESOURCE}{name}" for name in SCOPES_READWRITE]


def _auth_record_path() -> Path:
    return Path(get_config_dir()) / AUTH_RECORD_FILE


def _save_auth_record(record: AuthenticationRecord) -> None:
    """Persist AuthenticationRecord to disk."""
    path = _auth_record_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(record.serialize())
    path.chmod(0o600)


def _load_auth_record() -> AuthenticationRecord | None:
    """Load AuthenticationRecord from disk, or None if not found."""
    path = _auth_record_path()
    if not path.exists():
        return None
    try:
        return AuthenticationRecord.deserialize(path.read_text())
    except Exception:
        logger.warning("Failed to load auth record from %s", path)
        return None


class AuthManager:
    """Manages OAuth2 authentication for Microsoft Graph."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.credential: DeviceCodeCredential | None = None
        self._credentials: dict[str, DeviceCodeCredential] = {}
        self._active_account: str | None = config.default_account

    def get_scopes(self) -> list[str]:
        """Return individual scopes for display/consent purposes.

        Derived from the token scopes so the consent prompt can never
        advertise something different from what is actually requested.
        """
        return [s.rsplit("/", 1)[-1] for s in self.get_token_scopes()]

    def get_token_scopes(self) -> list[str]:
        """Return scopes for token acquisition — must match what the SDK requests."""
        return graph_token_scopes(self.config)

    def is_authenticated(self) -> bool:
        """Check if we have an active credential."""
        return self.credential is not None

    def _make_credential(
        self,
        prompt_callback=None,
        auth_record: AuthenticationRecord | None = None,
        interactive: bool = False,
    ) -> DeviceCodeCredential:
        """Create a DeviceCodeCredential with persistent cache.

        Non-interactive credentials (everything the MCP server uses) never
        start a device-code flow on a cache miss: that flow would block the
        stdio server for up to ``timeout`` seconds and print its prompt to
        stdout, the JSON-RPC channel. They raise AuthenticationRequiredError
        instead, which tools surface as AuthRequiredError.
        """
        global _warned_unencrypted_fallback
        allow_unencrypted = self.config.allow_unencrypted_token_cache

        # Fail closed: a plaintext refresh token is a long-lived mailbox
        # credential, so downgrading to one has to be the user's decision.
        if not allow_unencrypted and _unencrypted_fallback_will_be_used():
            raise UnencryptedCacheError()

        cache_name = token_cache_name()
        cache_options = TokenCachePersistenceOptions(
            name=cache_name,
            allow_unencrypted_storage=allow_unencrypted,
        )
        if (
            allow_unencrypted
            and not _warned_unencrypted_fallback
            and _unencrypted_fallback_will_be_used()
        ):
            logger.warning(
                "allow_unencrypted_token_cache is set, so the token cache "
                "will be stored unencrypted on disk: "
                "PyGObject/libsecret is not importable in this Python "
                "environment (common with `uv tool install` on Linux — "
                "the tool's isolated venv can't see system PyGObject). "
                "To enable encrypted caching via libsecret/gnome-keyring, "
                "install the system packages "
                "(apt: `gnome-keyring libsecret-1-0 python3-gi`) and "
                "re-create the venv with `--system-site-packages`. See "
                "https://github.com/mpalermiti/outlook-mcp/issues/7."
            )
            _warned_unencrypted_fallback = True
        kwargs = {
            "client_id": self.config.client_id,
            "tenant_id": self.config.tenant_id,
            "timeout": 900,
            "disable_automatic_authentication": not interactive,
        }
        if sys.platform == "darwin":
            kwargs["_cache"] = _macos_token_cache(cache_name)
        else:
            # libsecret and DPAPI already key the cache on its name.
            kwargs["cache_persistence_options"] = cache_options
        if prompt_callback:
            kwargs["prompt_callback"] = prompt_callback
        if auth_record:
            kwargs["authentication_record"] = auth_record
        return DeviceCodeCredential(**kwargs)

    def login_interactive(self, scopes: list[str]) -> None:
        """Run the device code flow interactively in the terminal.

        Uses get_token() which respects the token cache — if a valid
        cached token exists, completes silently. Otherwise triggers the
        device code flow. Saves the AuthenticationRecord for silent
        token refresh by the MCP server.

        Intended for CLI use (`outlook-mcp auth`), not MCP tools.
        """
        if not self.config.client_id:
            raise ValueError(
                "client_id is not configured. Register an Azure AD app and set "
                "client_id in ~/.outlook-mcp/config.json."
            )

        def _on_device_code(verification_uri: str, user_code: str, expires_on: object) -> None:
            print(f"Visit:  {verification_uri}")
            print(f"Code:   {user_code}")
            print()
            print("Waiting for you to complete sign-in in your browser...")

        cred = self._make_credential(prompt_callback=_on_device_code, interactive=True)
        # get_token() uses cache first, falls back to interactive.
        cred.get_token(*self.get_token_scopes())

        # Save the auth record for silent refresh by the MCP server
        record = getattr(cred, "_auth_record", None)
        if record:
            _save_auth_record(record)

        self.credential = cred
        print("Authenticated successfully.")

    def try_cached_token(self, scopes: list[str]) -> bool:
        """Try to get a token silently using a saved AuthenticationRecord.

        Returns True if a valid token was obtained without user interaction.
        Used by the MCP server on startup and by `outlook-mcp status`.
        """
        if not self.config.client_id:
            return False

        record = _load_auth_record()
        if record is None:
            return False

        try:
            cred = self._make_credential(auth_record=record)
            cred.get_token(*self.get_token_scopes())
            self.credential = cred
            return True
        except Exception as exc:
            logger.warning(
                "Cached token refresh failed (%s) — re-run `outlook-mcp auth`.",
                describe_auth_failure(exc),
            )
            return False

    def get_credential(self) -> DeviceCodeCredential:
        """Get the current credential, raising if not authenticated."""
        if self.credential is None:
            raise AuthRequiredError()
        return self.credential

    def list_accounts(self) -> list[dict]:
        """List configured accounts with auth status."""
        accounts = []
        for acc in self.config.accounts:
            accounts.append(
                {
                    "name": acc.name,
                    "client_id": acc.client_id[:8] + "...",
                    "tenant_id": acc.tenant_id,
                    "authenticated": acc.name in self._credentials,
                    "active": acc.name == self._active_account,
                }
            )
        if self.config.client_id and not self.config.accounts:
            accounts.append(
                {
                    "name": "default",
                    "client_id": self.config.client_id[:8] + "...",
                    "tenant_id": self.config.tenant_id,
                    "authenticated": self.credential is not None,
                    "active": True,
                }
            )
        return accounts

    def switch_account(self, name: str) -> dict:
        """Switch active account."""
        for acc in self.config.accounts:
            if acc.name == name:
                self._active_account = name
                if name in self._credentials:
                    self.credential = self._credentials[name]
                else:
                    self.credential = None
                return {"status": "switched", "account": name}
        raise ValueError(f"Account '{name}' not found in config")

    def logout(self) -> dict[str, str]:
        """Clear in-memory credentials and auth record."""
        self.credential = None
        path = _auth_record_path()
        if path.exists():
            path.unlink()
        return {"status": "logged_out", "message": "Credentials cleared."}
