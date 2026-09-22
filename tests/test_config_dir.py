"""Tests for OUTLOOK_MCP_CONFIG_DIR — running one server instance per account."""

import os

from outlook_mcp.auth import AUTH_RECORD_FILE, _auth_record_path
from outlook_mcp.config import Config, get_config_dir, load_config, save_config


def test_config_dir_defaults_to_home(monkeypatch):
    monkeypatch.delenv("OUTLOOK_MCP_CONFIG_DIR", raising=False)
    assert get_config_dir() == os.path.expanduser("~/.outlook-mcp")


def test_config_dir_from_env_is_expanded(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", "~/.outlook-mcp-pobox")
    assert get_config_dir() == os.path.expanduser("~/.outlook-mcp-pobox")


def test_empty_env_falls_back_to_default(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", "")
    assert get_config_dir() == os.path.expanduser("~/.outlook-mcp")


def test_load_and_save_config_follow_env(tmp_path, monkeypatch):
    """A second instance must read its own config.json, not the default one."""
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(tmp_path))
    save_config(Config(client_id="pobox-app"))
    assert (tmp_path / "config.json").exists()
    assert load_config().client_id == "pobox-app"


def test_auth_record_follows_env(tmp_path, monkeypatch):
    """Each instance keeps its own AuthenticationRecord, which selects the
    account's refresh token out of the shared OS token cache."""
    monkeypatch.setenv("OUTLOOK_MCP_CONFIG_DIR", str(tmp_path))
    assert _auth_record_path() == tmp_path / AUTH_RECORD_FILE
