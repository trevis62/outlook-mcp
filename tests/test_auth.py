"""Tests for auth module."""

import logging
import sys
from unittest.mock import patch

import pytest

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager, _unencrypted_fallback_will_be_used
from outlook_mcp.config import Config
from outlook_mcp.errors import AuthRequiredError, UnencryptedCacheError


@pytest.fixture(autouse=True)
def _reset_unencrypted_warning_latch():
    """Reset the once-per-process warning latch between tests."""
    auth_module._warned_unencrypted_fallback = False
    yield
    auth_module._warned_unencrypted_fallback = False


def test_auth_manager_init():
    """AuthManager initializes with config."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.config is config
    assert auth.credential is None


def test_auth_scopes_default():
    """Default scopes include read-write."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.ReadWrite" in scopes
    assert "Mail.Send" in scopes
    assert "Calendars.ReadWrite" in scopes
    # offline_access is reserved — MSAL adds it automatically
    assert "offline_access" not in scopes


def test_auth_scopes_read_only():
    """read_only no longer narrows the displayed consent scopes.

    It previously advertised Mail.Read etc., but those are not what gets
    requested — the documented app registration grants only the ReadWrite
    variants, and asking for scopes that were never consented drops the
    server into an interactive flow. Display now mirrors the real request so
    the consent prompt cannot advertise something different from what is
    asked for. Token-level narrowing is opt-in via config.graph_scopes.
    """
    config = Config(client_id="test-id", read_only=True)
    auth = AuthManager(config)
    scopes = auth.get_scopes()
    assert "Mail.ReadWrite" in scopes
    assert scopes == AuthManager(Config(client_id="test-id")).get_scopes()


def test_auth_scopes_follow_graph_scopes_override():
    """A narrower app registration is reflected in the consent display."""
    config = Config(
        client_id="test-id",
        graph_scopes=["https://graph.microsoft.com/Mail.Read"],
    )
    assert AuthManager(config).get_scopes() == ["Mail.Read"]


def test_auth_not_authenticated():
    """is_authenticated returns False before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    assert auth.is_authenticated() is False


def test_auth_get_credential_raises_when_not_authenticated():
    """get_credential raises AuthRequiredError before login."""
    config = Config(client_id="test-id")
    auth = AuthManager(config)
    with pytest.raises(AuthRequiredError):
        auth.get_credential()


def test_login_interactive_requires_client_id():
    """login_interactive raises if client_id is not configured."""
    config = Config()  # No client_id
    auth = AuthManager(config)
    with pytest.raises(ValueError, match="client_id"):
        auth.login_interactive(auth.get_scopes())


def test_try_cached_token_returns_false_without_client_id():
    """try_cached_token returns False if client_id is not set."""
    config = Config()
    auth = AuthManager(config)
    assert auth.try_cached_token(auth.get_scopes()) is False


class TestUnencryptedFallbackDetection:
    """_unencrypted_fallback_will_be_used mirrors msal_extensions' check."""

    def test_macos_is_never_fallback(self):
        """macOS uses Keychain — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "darwin"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_windows_is_never_fallback(self):
        """Windows uses DPAPI — fallback is impossible."""
        with patch.object(auth_module.sys, "platform", "win32"):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_with_gi_available(self):
        """Linux with PyGObject importable uses libsecret — no fallback."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=object()),
        ):
            assert _unencrypted_fallback_will_be_used() is False

    def test_linux_without_gi_uses_fallback(self):
        """Linux without PyGObject (issue #7) triggers the fallback path."""
        with (
            patch.object(auth_module.sys, "platform", "linux"),
            patch.object(auth_module.importlib.util, "find_spec", return_value=None),
        ):
            assert _unencrypted_fallback_will_be_used() is True


class TestUnencryptedFallbackWarning:
    """_make_credential emits a warning at most once when fallback is in use."""

    def test_warning_fires_once_when_fallback_active(self, caplog):
        """A single warning is logged on the first credential build.

        Only reachable once the user has opted into plaintext caching —
        without the opt-in, _make_credential raises instead of warning.
        """
        config = Config(client_id="test-id", allow_unencrypted_token_cache=True)
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=True,
            ),
        ):
            auth._make_credential()
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert len(fallback_warnings) == 1
        assert fallback_warnings[0].levelno == logging.WARNING

    def test_no_warning_when_fallback_inactive(self, caplog):
        """No fallback warning is logged when encrypted storage is available."""
        config = Config(client_id="test-id")
        auth = AuthManager(config)

        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
        ):
            auth._make_credential()

        fallback_warnings = [r for r in caplog.records if "unencrypted" in r.getMessage().lower()]
        assert fallback_warnings == []


class TestUnencryptedCacheOptIn:
    """Plaintext token caching must be a deliberate choice, not a silent fallback."""

    def test_refuses_to_build_credential_when_cache_would_be_plaintext(self):
        """Default config must fail closed rather than write a token in the clear."""
        auth = AuthManager(Config(client_id="test-id"))

        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True):
            with pytest.raises(UnencryptedCacheError):
                auth._make_credential()

    def test_builds_credential_when_user_opts_in(self):
        """Explicit opt-in is honored — some Linux setups have no keyring at all."""
        auth = AuthManager(Config(client_id="test-id", allow_unencrypted_token_cache=True))

        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=True):
            assert auth._make_credential() is not None

    def test_builds_credential_when_encrypted_storage_is_available(self):
        """macOS/Windows and Linux-with-libsecret are unaffected."""
        auth = AuthManager(Config(client_id="test-id"))

        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False):
            assert auth._make_credential() is not None

    def test_library_is_also_told_not_to_store_unencrypted(self):
        """Defense in depth: our platform heuristic could be wrong.

        Even if _unencrypted_fallback_will_be_used misjudges the platform,
        azure-identity itself must refuse to write a plaintext cache unless
        the user opted in.
        """
        auth = AuthManager(Config(client_id="test-id"))

        with (
            patch(
                "outlook_mcp.auth._unencrypted_fallback_will_be_used",
                return_value=False,
            ),
            patch("outlook_mcp.auth.TokenCachePersistenceOptions") as mock_opts,
        ):
            auth._make_credential()

        assert mock_opts.call_args.kwargs["allow_unencrypted_storage"] is False


class TestDelegatedTokenScopes:
    """Token acquisition must use explicit delegated scopes, not `.default`.

    `.default` yields a token that Microsoft Graph rejects with 403
    ErrorAccessDenied on every mailbox endpoint for personal Microsoft
    accounts (verified against a live outlook.com account: /me returns 200,
    /me/messages returns 403, while the same account with explicitly-scoped
    tokens returns 200 for both). Personal accounts are this project's
    primary target, so `.default` must not appear in a token request.
    """

    def test_read_write_scopes_are_explicit_and_graph_qualified(self):
        auth = AuthManager(Config(client_id="test-id"))
        scopes = auth.get_token_scopes()

        assert all(s.startswith("https://graph.microsoft.com/") for s in scopes)
        assert not any(s.endswith("/.default") for s in scopes)
        assert "https://graph.microsoft.com/Mail.ReadWrite" in scopes
        assert "https://graph.microsoft.com/Mail.Send" in scopes
        assert "https://graph.microsoft.com/Calendars.ReadWrite" in scopes

    def test_read_only_does_not_narrow_token_scopes(self):
        """read_only must not change which scopes are requested.

        Tempting to request Mail.Read here, but the app registration the
        README documents grants only the ReadWrite variants. Scope matching
        is literal — Mail.ReadWrite does not satisfy a request for Mail.Read
        — so narrowing produces an MSAL cache miss and drops the server into
        an interactive device-code flow it cannot answer. Verified against a
        live account: it hangs.

        read_only stays what it already was: server-side enforcement via
        check_permission, not a token-scope restriction.
        """
        rw = AuthManager(Config(client_id="test-id")).get_token_scopes()
        ro = AuthManager(Config(client_id="test-id", read_only=True)).get_token_scopes()
        assert ro == rw

    def test_graph_scopes_override_is_honored(self):
        """Escape hatch for an app registration granting only read permissions.

        Someone who wants token-level restriction registers a narrower Azure
        app and names the scopes here, rather than having them inferred.
        """
        custom = ["https://graph.microsoft.com/Mail.Read"]
        auth = AuthManager(Config(client_id="test-id", graph_scopes=custom))
        assert auth.get_token_scopes() == custom

    def test_token_scopes_match_the_advertised_consent_scopes(self):
        """Consent display and token request must not drift apart.

        If they diverge, the user consents to one set and the server
        requests another, which is exactly the cache-miss / silent-reauth
        failure the previous `.default` approach was working around.
        """
        auth = AuthManager(Config(client_id="test-id"))
        display = set(auth.get_scopes())
        requested = {s.rsplit("/", 1)[-1] for s in auth.get_token_scopes()}
        assert display == requested


class TestPerInstanceTokenCache:
    """Each OUTLOOK_MCP_CONFIG_DIR gets its own token cache.

    Two instances sharing one cache let a cache miss or sign-in in one
    disturb the other; on macOS they even shared one Keychain item.
    """

    def test_default_dir_keeps_historical_cache_name(self, monkeypatch):
        """Existing single-instance installs must not have to re-authenticate."""
        monkeypatch.delenv("OUTLOOK_MCP_CONFIG_DIR", raising=False)
        assert auth_module.token_cache_name() == auth_module.CACHE_NAME

    def test_other_dirs_get_distinct_stable_names(self, tmp_path, monkeypatch):
        names = []
        for sub in ("a", "b", "a"):
            monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(tmp_path / sub))
            names.append(auth_module.token_cache_name())
        assert names[0] != auth_module.CACHE_NAME
        assert names[0].startswith(auth_module.CACHE_NAME + "-")
        assert names[0] != names[1]
        assert names[0] == names[2]

    def _build_macos_cache(self, monkeypatch):
        import msal_extensions

        calls = []
        monkeypatch.setattr(
            msal_extensions, "KeychainPersistence", lambda *a: calls.append(a) or object()
        )
        monkeypatch.setattr(msal_extensions, "PersistedTokenCache", lambda p: p)
        with patch.object(auth_module.sys, "platform", "darwin"):
            auth_module.AuthManager(Config(client_id="test-id"))._make_credential()
        return calls[0]

    def test_macos_default_dir_maps_to_azure_identity_keychain_item(self, monkeypatch):
        monkeypatch.delenv("OUTLOOK_MCP_CONFIG_DIR", raising=False)
        signal, service, account = self._build_macos_cache(monkeypatch)
        assert service == "Microsoft.Developer.IdentityService"
        assert account == "MSALCache"
        assert signal.endswith("/.IdentityService/outlook-mcp.nocae")

    def test_macos_other_dir_uses_its_own_keychain_item(self, tmp_path, monkeypatch):
        monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(tmp_path))
        name = auth_module.token_cache_name()
        signal, service, account = self._build_macos_cache(monkeypatch)
        assert account == name
        assert signal.endswith(f"/.IdentityService/{name}.nocae")

    @pytest.mark.skipif(sys.platform != "darwin", reason="Keychain persistence is macOS-only")
    def test_macos_default_matches_what_azure_identity_builds(self, monkeypatch):
        """Guards the no-re-auth promise against azure-identity changing its layout."""
        from azure.identity import TokenCachePersistenceOptions
        from azure.identity._persistent_cache import _load_persistent_cache

        monkeypatch.delenv("OUTLOOK_MCP_CONFIG_DIR", raising=False)
        ours = auth_module._macos_token_cache(auth_module.CACHE_NAME)._persistence
        theirs = _load_persistent_cache(
            TokenCachePersistenceOptions(name="outlook-mcp")
        )._persistence
        assert ours.get_location() == theirs.get_location()
        assert ours._service_name == theirs._service_name
        assert ours._account_name == theirs._account_name

    def test_device_code_credential_still_accepts_private_cache_kwarg(self):
        """We rely on DeviceCodeCredential's private `_cache` argument on macOS.

        If an azure-identity upgrade drops it, fail here rather than silently
        falling back to the shared Keychain item.
        """
        import msal
        from azure.identity import DeviceCodeCredential

        cache = msal.TokenCache()
        cred = DeviceCodeCredential(client_id="test-id", _cache=cache)
        assert cred._cache is cache
        assert cred._custom_cache is True


class TestNoInteractiveFallbackInServer:
    """The server must never start a device-code flow on a cache miss.

    That flow blocks the stdio server (initialize times out) and prints its
    prompt to stdout, which is the JSON-RPC channel.
    """

    def _credential(self, interactive):
        with patch("outlook_mcp.auth._unencrypted_fallback_will_be_used", return_value=False):
            return AuthManager(Config(client_id="test-id"))._make_credential(
                interactive=interactive
            )

    def test_server_credential_disables_automatic_authentication(self):
        assert self._credential(interactive=False)._disable_automatic_authentication is True

    def test_cli_login_credential_allows_device_code_flow(self):
        assert self._credential(interactive=True)._disable_automatic_authentication is False

    def test_cache_miss_raises_instead_of_prompting(self):
        import msal
        from azure.identity import AuthenticationRequiredError, DeviceCodeCredential

        cred = DeviceCodeCredential(
            client_id="test-id", _cache=msal.TokenCache(), disable_automatic_authentication=True
        )
        with patch.object(
            DeviceCodeCredential, "_request_token", side_effect=AssertionError("prompted")
        ):
            with pytest.raises(AuthenticationRequiredError):
                cred.get_token("https://graph.microsoft.com/Mail.Read")

    def test_try_cached_token_builds_non_interactive_credential(self, tmp_path, monkeypatch):
        monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(tmp_path))
        (tmp_path / auth_module.AUTH_RECORD_FILE).write_text("{}")
        auth = AuthManager(Config(client_id="test-id"))
        with (
            patch.object(auth_module, "_load_auth_record", return_value=object()),
            patch.object(auth, "_make_credential") as make,
        ):
            assert auth.try_cached_token(auth.get_token_scopes()) is True
        assert make.call_args.kwargs.get("interactive", False) is False


class _FakeKeychainError(OSError):
    """Mimics msal_extensions.osx.KeychainError: str() is empty."""

    def __init__(self, exit_status):
        super().__init__()
        self.exit_status = exit_status


def _wrapped(cause):
    from azure.core.exceptions import ClientAuthenticationError

    try:
        try:
            raise cause
        except Exception as ex:
            raise ClientAuthenticationError(message=f"Authentication failed: {ex}") from ex
    except ClientAuthenticationError as outer:
        return outer


class TestDescribeAuthFailure:
    """Token failures must name their cause instead of logging a blank reason."""

    def test_keychain_error_code_is_surfaced(self):
        reason = auth_module.describe_auth_failure(_wrapped(_FakeKeychainError(-25308)))
        assert "-25308" in reason
        assert "errSecInteractionNotAllowed" in reason

    def test_unknown_keychain_code_still_reported(self):
        reason = auth_module.describe_auth_failure(_wrapped(_FakeKeychainError(-1)))
        assert "macOS Keychain error -1" in reason

    def test_uses_first_message_with_content(self):
        reason = auth_module.describe_auth_failure(_wrapped(RuntimeError("lock busy")))
        assert reason.endswith("Authentication failed: lock busy")

    def test_no_details_anywhere(self):
        reason = auth_module.describe_auth_failure(_wrapped(RuntimeError()))
        assert "no details" in reason

    def test_try_cached_token_logs_the_reason(self, caplog, monkeypatch):
        auth = AuthManager(Config(client_id="test-id"))
        failing = type(
            "C",
            (),
            {"get_token": lambda *a: (_ for _ in ()).throw(_wrapped(_FakeKeychainError(-128)))},
        )()
        with (
            caplog.at_level(logging.WARNING, logger="outlook_mcp.auth"),
            patch.object(auth_module, "_load_auth_record", return_value=object()),
            patch.object(auth, "_make_credential", return_value=failing),
        ):
            assert auth.try_cached_token(auth.get_token_scopes()) is False
        assert "-128" in caplog.text
