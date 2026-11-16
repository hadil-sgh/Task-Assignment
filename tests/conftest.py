"""Shared test setup: every test runs offline and without an API key."""

import socket

import pytest

from slidegen.template import TEMPLATE_DIR, Catalog, build_catalog

API_KEY_VAR = "OPENROUTER_API_KEY"


def _blocked(*args, **kwargs):
    raise RuntimeError("Network access is blocked in tests. Mock the model client instead.")


@pytest.fixture(autouse=True)
def offline(monkeypatch: pytest.MonkeyPatch) -> None:
    """Remove the API key and make any attempt to open a connection fail loudly."""
    monkeypatch.delenv(API_KEY_VAR, raising=False)
    monkeypatch.setattr(socket.socket, "connect", _blocked)
    monkeypatch.setattr(socket, "create_connection", _blocked)
    monkeypatch.setattr(socket, "getaddrinfo", _blocked)


@pytest.fixture(scope="session")
def catalog() -> Catalog:
    """Catalog of the base template (built directly, so no cache file is written)."""
    return build_catalog(TEMPLATE_DIR / "base.pptx")
