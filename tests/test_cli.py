"""Tests for the CLI entry points."""

from unittest.mock import patch

from outlook_mcp.auth import AUTH_RECORD_FILE
from outlook_mcp.cli import cmd_logout
from outlook_mcp.config import Config


def _patch_paths(tmp_path):
    """Point the auth record at tmp_path and stub out config loading."""
    return (
        patch.dict("os.environ", {"OUTLOOK_MCP_CONFIG_DIR": str(tmp_path)}),
        patch("outlook_mcp.cli.load_config", return_value=Config(client_id="test")),
    )


class TestCmdLogout:
    def test_deletes_the_persisted_auth_record(self, tmp_path, capsys):
        """`outlook-mcp logout` must actually revoke silent re-auth.

        The auth record is what lets the MCP server refresh a token with no
        user interaction, so leaving it behind means a user who ran logout
        is still logged in on the next server start.
        """
        record = tmp_path / AUTH_RECORD_FILE
        record.write_text('{"username": "user@example.com"}')

        p_dir, p_cfg = _patch_paths(tmp_path)
        with p_dir, p_cfg:
            cmd_logout()

        assert not record.exists()

    def test_is_idempotent_when_no_record_exists(self, tmp_path):
        """Logging out twice must not raise."""
        p_dir, p_cfg = _patch_paths(tmp_path)
        with p_dir, p_cfg:
            cmd_logout()
            cmd_logout()

    def test_still_tells_user_about_the_os_credential_store(self, tmp_path, capsys):
        """The refresh token lives in the keychain, which we can't clear.

        azure-identity exposes no cache-clear API, so the user has to remove
        it by hand — saying so is the difference between a partial logout the
        user knows about and one they don't.
        """
        p_dir, p_cfg = _patch_paths(tmp_path)
        with p_dir, p_cfg:
            cmd_logout()

        out = capsys.readouterr().out
        assert "Keychain" in out or "credential store" in out
