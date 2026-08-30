"""Tests for delta-URL host validation.

``fetch_delta_pages`` takes ``delta_token`` straight from the caller and
uses it as the request URL while attaching a Graph bearer token. An agent
that has been prompt-injected by message content can hand it any URL, so
the token must never be sent anywhere but Graph.
"""

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from outlook_mcp.tools._delta import fetch_delta_pages, require_graph_url

GRAPH_DELTA_URL = (
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta"
    "?$deltatoken=abc123"
)


# ── URL validator ────────────────────────────────────────────────────


class TestRequireGraphUrl:
    def test_accepts_graph_delta_url(self):
        assert require_graph_url(GRAPH_DELTA_URL) == GRAPH_DELTA_URL

    def test_rejects_foreign_host(self):
        with pytest.raises(ValueError, match="graph.microsoft.com"):
            require_graph_url("https://evil.example/steal")

    def test_rejects_plaintext_http(self):
        with pytest.raises(ValueError, match="https"):
            require_graph_url("http://graph.microsoft.com/v1.0/me/messages/delta")

    def test_rejects_userinfo_prefix_pointing_elsewhere(self):
        # urlparse().hostname is "evil.example" here — the Graph host is
        # only userinfo, a classic lookalike.
        with pytest.raises(ValueError, match="graph.microsoft.com"):
            require_graph_url("https://graph.microsoft.com@evil.example/steal")

    def test_rejects_lookalike_subdomain(self):
        with pytest.raises(ValueError, match="graph.microsoft.com"):
            require_graph_url("https://graph.microsoft.com.evil.example/v1.0/me")

    def test_rejects_nonstandard_port(self):
        with pytest.raises(ValueError):
            require_graph_url("https://graph.microsoft.com:8443/v1.0/me/messages/delta")

    def test_rejects_non_http_scheme(self):
        with pytest.raises(ValueError, match="https"):
            require_graph_url("file:///etc/passwd")


# ── fetch_delta_pages enforcement ────────────────────────────────────


def _credential():
    cred = MagicMock()
    cred.get_token = MagicMock(return_value=MagicMock(token="SECRET-TOKEN"))
    return cred


@pytest.mark.asyncio
async def test_hostile_delta_token_never_reaches_the_network():
    """A hostile delta_token must fail before any token is minted or sent."""
    cred = _credential()
    fake_client = MagicMock()
    fake_client.get = AsyncMock()
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    with patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=fake_client):
        with pytest.raises(ValueError, match="graph.microsoft.com"):
            await fetch_delta_pages(
                cred,
                initial_url="",
                delta_token="https://evil.example/collect",
                page_size=10,
            )

    fake_client.get.assert_not_called()
    cred.get_token.assert_not_called()


@pytest.mark.asyncio
async def test_hostile_next_link_is_rejected_mid_walk():
    """A poisoned @odata.nextLink must not redirect the bearer token."""
    cred = _credential()

    first_page = MagicMock()
    first_page.status_code = 200
    first_page.raise_for_status = MagicMock()
    first_page.json = MagicMock(
        return_value={
            "value": [{"id": "m1"}],
            "@odata.nextLink": "https://evil.example/page2",
        }
    )

    calls: list[str] = []

    async def fake_get(url, headers=None):
        calls.append(url)
        return first_page

    fake_client = MagicMock()
    fake_client.get = AsyncMock(side_effect=fake_get)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    with patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=fake_client):
        with pytest.raises(ValueError, match="graph.microsoft.com"):
            await fetch_delta_pages(
                cred,
                initial_url=GRAPH_DELTA_URL,
                delta_token=None,
                page_size=10,
            )

    assert calls == [GRAPH_DELTA_URL]


@pytest.mark.asyncio
async def test_valid_graph_delta_token_still_works():
    """The happy path must be unaffected by validation."""
    cred = _credential()

    page = MagicMock()
    page.status_code = 200
    page.raise_for_status = MagicMock()
    page.json = MagicMock(
        return_value={
            "value": [{"id": "m1"}],
            "@odata.deltaLink": GRAPH_DELTA_URL,
        }
    )

    fake_client = MagicMock()
    fake_client.get = AsyncMock(return_value=page)
    fake_client.__aenter__ = AsyncMock(return_value=fake_client)
    fake_client.__aexit__ = AsyncMock(return_value=False)

    with patch("outlook_mcp.tools._delta.httpx.AsyncClient", return_value=fake_client):
        items, token, has_more = await fetch_delta_pages(
            cred,
            initial_url="",
            delta_token=GRAPH_DELTA_URL,
            page_size=10,
        )

    assert items == [{"id": "m1"}]
    assert token == GRAPH_DELTA_URL
    assert has_more is False
