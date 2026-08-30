"""Tests for filesystem path confinement.

The attachment tools take host filesystem paths straight from the caller.
Since the caller is an agent that may be acting on attacker-authored
message content, a download must not be able to write outside its
sandbox, and a send must not be able to read arbitrary files off the host.
"""

import os
from pathlib import Path

import pytest

from outlook_mcp.config import Config
from outlook_mcp.paths import require_allowed_source, resolve_download_path


def _cfg(tmp_path: Path, **overrides) -> Config:
    base = {"client_id": "test", "download_dir": str(tmp_path / "downloads")}
    base.update(overrides)
    return Config(**base)


# ── Download confinement ─────────────────────────────────────────────


class TestResolveDownloadPath:
    def test_relative_name_lands_inside_download_dir(self, tmp_path):
        cfg = _cfg(tmp_path)
        result = resolve_download_path(cfg, "report.pdf")
        assert result == tmp_path / "downloads" / "report.pdf"

    def test_creates_download_dir_if_missing(self, tmp_path):
        cfg = _cfg(tmp_path)
        resolve_download_path(cfg, "report.pdf")
        assert (tmp_path / "downloads").is_dir()

    def test_download_dir_is_private(self, tmp_path):
        cfg = _cfg(tmp_path)
        resolve_download_path(cfg, "report.pdf")
        mode = (tmp_path / "downloads").stat().st_mode & 0o777
        assert mode == 0o700

    def test_absolute_path_inside_download_dir_is_allowed(self, tmp_path):
        cfg = _cfg(tmp_path)
        target = tmp_path / "downloads" / "nested" / "report.pdf"
        assert resolve_download_path(cfg, str(target)) == target

    def test_nested_parent_dirs_are_created(self, tmp_path):
        cfg = _cfg(tmp_path)
        resolve_download_path(cfg, "a/b/report.pdf")
        assert (tmp_path / "downloads" / "a" / "b").is_dir()

    def test_rejects_absolute_path_outside_download_dir(self, tmp_path):
        cfg = _cfg(tmp_path)
        with pytest.raises(ValueError, match="outside"):
            resolve_download_path(cfg, str(tmp_path / "zshrc"))

    def test_rejects_parent_traversal(self, tmp_path):
        cfg = _cfg(tmp_path)
        with pytest.raises(ValueError, match="outside"):
            resolve_download_path(cfg, "../../.zshrc")

    def test_rejects_symlink_escaping_the_download_dir(self, tmp_path):
        """A symlink planted inside the sandbox must not redirect the write."""
        cfg = _cfg(tmp_path)
        download_dir = tmp_path / "downloads"
        download_dir.mkdir(parents=True)
        outside = tmp_path / "outside"
        outside.mkdir()
        (download_dir / "escape").symlink_to(outside)

        with pytest.raises(ValueError, match="outside"):
            resolve_download_path(cfg, str(download_dir / "escape" / "payload.sh"))

    def test_expands_tilde_in_configured_dir(self, tmp_path, monkeypatch):
        monkeypatch.setenv("HOME", str(tmp_path))
        cfg = Config(client_id="test", download_dir="~/dl")
        assert resolve_download_path(cfg, "x.pdf") == tmp_path / "dl" / "x.pdf"


# ── Source confinement ───────────────────────────────────────────────


class TestRequireAllowedSource:
    def test_denies_everything_when_no_dirs_configured(self, tmp_path):
        """Empty allowlist fails closed, and says how to open it."""
        cfg = _cfg(tmp_path)
        f = tmp_path / "doc.pdf"
        f.write_text("x")
        with pytest.raises(ValueError, match="attachment_source_dirs"):
            require_allowed_source(cfg, str(f))

    def test_allows_file_inside_configured_dir(self, tmp_path):
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        f = allowed / "doc.pdf"
        f.write_text("x")
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        assert require_allowed_source(cfg, str(f)) == f

    def test_allows_file_in_nested_subdir(self, tmp_path):
        allowed = tmp_path / "outbox"
        nested = allowed / "deep"
        nested.mkdir(parents=True)
        f = nested / "doc.pdf"
        f.write_text("x")
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        assert require_allowed_source(cfg, str(f)) == f

    def test_rejects_file_outside_configured_dir(self, tmp_path):
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        secret = tmp_path / "id_rsa"
        secret.write_text("PRIVATE KEY")
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        with pytest.raises(ValueError, match="not in an allowed"):
            require_allowed_source(cfg, str(secret))

    def test_rejects_symlink_pointing_outside(self, tmp_path):
        """The classic exfil trick: a symlink inside an allowed dir."""
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        secret = tmp_path / "id_rsa"
        secret.write_text("PRIVATE KEY")
        (allowed / "innocent.txt").symlink_to(secret)
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        with pytest.raises(ValueError, match="not in an allowed"):
            require_allowed_source(cfg, str(allowed / "innocent.txt"))

    def test_rejects_sibling_dir_with_shared_prefix(self, tmp_path):
        """`/a/outbox-evil` must not pass a check for `/a/outbox`."""
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        evil = tmp_path / "outbox-evil"
        evil.mkdir()
        f = evil / "doc.pdf"
        f.write_text("x")
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        with pytest.raises(ValueError, match="not in an allowed"):
            require_allowed_source(cfg, str(f))

    def test_honors_multiple_allowed_dirs(self, tmp_path):
        one = tmp_path / "one"
        two = tmp_path / "two"
        one.mkdir()
        two.mkdir()
        f = two / "doc.pdf"
        f.write_text("x")
        cfg = _cfg(cfg_dir := tmp_path, attachment_source_dirs=[str(one), str(two)])
        assert require_allowed_source(cfg, str(f)) == f
        assert cfg_dir.exists()

    def test_rejects_missing_file(self, tmp_path):
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        with pytest.raises(FileNotFoundError):
            require_allowed_source(cfg, str(allowed / "nope.pdf"))

    def test_does_not_leak_existence_of_files_outside_the_allowlist(self, tmp_path):
        """Probing an out-of-bounds path must not reveal whether it exists.

        Both a real and an imaginary out-of-bounds path must fail the same
        way, or the tool becomes a filesystem existence oracle.
        """
        allowed = tmp_path / "outbox"
        allowed.mkdir()
        real_secret = tmp_path / "id_rsa"
        real_secret.write_text("PRIVATE KEY")
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])

        with pytest.raises(ValueError, match="not in an allowed") as real_exc:
            require_allowed_source(cfg, str(real_secret))
        with pytest.raises(ValueError, match="not in an allowed") as fake_exc:
            require_allowed_source(cfg, str(tmp_path / "does_not_exist"))

        assert type(real_exc.value) is type(fake_exc.value)

    def test_rejects_directory_argument(self, tmp_path):
        allowed = tmp_path / "outbox"
        (allowed / "sub").mkdir(parents=True)
        cfg = _cfg(tmp_path, attachment_source_dirs=[str(allowed)])
        with pytest.raises(ValueError, match="not a regular file"):
            require_allowed_source(cfg, str(allowed / "sub"))

    def test_expands_tilde_in_configured_dirs(self, tmp_path, monkeypatch):
        monkeypatch.setenv("HOME", str(tmp_path))
        outbox = tmp_path / "outbox"
        outbox.mkdir()
        f = outbox / "doc.pdf"
        f.write_text("x")
        cfg = Config(client_id="test", attachment_source_dirs=["~/outbox"])
        assert require_allowed_source(cfg, str(f)) == f
        assert os.path.isfile(f)
