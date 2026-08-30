"""Tests for Graph client factory."""

from unittest.mock import MagicMock

import pytest

from outlook_mcp.errors import AuthRequiredError
from outlook_mcp.graph import GraphClient


def test_graph_client_requires_credential():
    """GraphClient raises without credential."""
    with pytest.raises(AuthRequiredError):
        GraphClient(credential=None)


def test_graph_client_init():
    """GraphClient initializes with a credential and creates sdk_client."""
    mock_credential = MagicMock()
    client = GraphClient(credential=mock_credential)
    assert client.sdk_client is not None


def test_graph_client_exposes_the_scopes_it_was_built_with():
    """Raw-httpx paths mint their own tokens and must reuse these scopes.

    Requesting a different scope set elsewhere causes an MSAL cache miss,
    which triggers a background interactive auth the server cannot answer.
    """
    scopes = ["https://graph.microsoft.com/Mail.Read"]
    client = GraphClient(credential=MagicMock(), scopes=scopes)
    assert client.scopes == scopes


def test_graph_client_passes_scopes_to_the_auth_provider():
    """The SDK defaults to `.default`, which 403s on personal accounts."""
    from unittest.mock import patch

    scopes = ["https://graph.microsoft.com/Mail.Read"]
    with patch("outlook_mcp.graph.AzureIdentityAuthenticationProvider") as provider:
        GraphClient(credential=MagicMock(), scopes=scopes)

    assert provider.call_args.kwargs["scopes"] == scopes


def test_graph_client_defaults_to_read_write_scopes():
    """Callers that don't specify scopes still must not get `.default`."""
    client = GraphClient(credential=MagicMock())
    assert client.scopes
    assert not any(s.endswith("/.default") for s in client.scopes)
