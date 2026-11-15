"""Prove the offline fixture works, so no test can ever reach the real model."""

import os
import socket

import pytest

from tests.conftest import API_KEY_VAR


def test_network_is_blocked() -> None:
    with pytest.raises(RuntimeError, match="Network access is blocked"):
        socket.create_connection(("openrouter.ai", 443))


def test_dns_is_blocked() -> None:
    with pytest.raises(RuntimeError, match="Network access is blocked"):
        socket.getaddrinfo("openrouter.ai", 443)


def test_api_key_is_removed() -> None:
    assert API_KEY_VAR not in os.environ
