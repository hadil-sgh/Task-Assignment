"""The model client, tested against a fake transport: no network, no real key."""

import json
from pathlib import Path

import httpx
import pytest

from slidegen.llm import (
    API_URL,
    KEY_VAR,
    LLMError,
    MissingKeyError,
    MockClient,
    OpenRouterClient,
    RejectedKeyError,
    UnavailableError,
    load_api_key,
)

FAKE_KEY = "sk-or-test-not-a-real-key"
MESSAGES = [{"role": "user", "content": "<notes>\nhello\n</notes>"}]
SCHEMA = {"name": "deck", "strict": True, "schema": {"type": "object"}}
REPLY = {
    "model": "test/model",
    "choices": [{"message": {"content": '{"slides": []}'}}],
    "usage": {"prompt_tokens": 900, "completion_tokens": 300, "cost": 0.0042},
}


def client_with(*responses, waits: list | None = None) -> tuple[OpenRouterClient, list]:
    """A client whose HTTP calls are answered, in order, by `responses`.

    Each response is a status code, a (status, json) pair, or an exception to raise.
    Returns the client and the list of requests it sent.
    """
    queue, sent = list(responses), []

    def handler(request: httpx.Request) -> httpx.Response:
        sent.append(request)
        item = queue.pop(0)
        if isinstance(item, Exception):
            raise item
        status, body = item if isinstance(item, tuple) else (item, {})
        return httpx.Response(status, json=body)

    record = waits.append if waits is not None else (lambda seconds: None)
    client = OpenRouterClient(FAKE_KEY, "test/model", httpx.MockTransport(handler), record)
    return client, sent


# ---------------------------------------------------------------- a normal call


def test_success_returns_text_and_usage() -> None:
    client, _ = client_with((200, REPLY))
    text, usage = client.complete(MESSAGES, SCHEMA)
    assert text == '{"slides": []}'
    assert (usage.model, usage.prompt_tokens, usage.completion_tokens) == ("test/model", 900, 300)
    assert usage.cost_usd == 0.0042
    assert usage.latency_s >= 0


def test_request_carries_key_model_and_strict_schema() -> None:
    client, sent = client_with((200, REPLY))
    client.complete(MESSAGES, SCHEMA)
    [request] = sent
    assert str(request.url) == API_URL
    assert request.headers["authorization"] == f"Bearer {FAKE_KEY}"
    body = json.loads(request.content)
    assert body["model"] == "test/model"
    assert body["messages"] == MESSAGES
    assert body["response_format"] == {"type": "json_schema", "json_schema": SCHEMA}
    assert body["usage"] == {"include": True}


def test_error_inside_a_200_reply_is_reported() -> None:
    client, _ = client_with((200, {"error": {"message": "provider overloaded"}}))
    with pytest.raises(LLMError, match="provider overloaded"):
        client.complete(MESSAGES, SCHEMA)


# ---------------------------------------------------------------- retries


def test_rate_limit_is_retried_with_backoff() -> None:
    waits: list[float] = []
    client, sent = client_with(429, 503, (200, REPLY), waits=waits)
    text, _ = client.complete(MESSAGES, SCHEMA)
    assert text == '{"slides": []}'
    assert len(sent) == 3
    assert waits == [1.0, 2.0]


def test_retry_after_header_is_honoured() -> None:
    waits: list[float] = []

    def handler(request: httpx.Request) -> httpx.Response:
        if not waits:
            return httpx.Response(429, headers={"Retry-After": "7"})
        return httpx.Response(200, json=REPLY)

    client = OpenRouterClient(FAKE_KEY, "test/model", httpx.MockTransport(handler), waits.append)
    client.complete(MESSAGES, SCHEMA)
    assert waits == [7.0]


def test_timeout_is_retried() -> None:
    client, sent = client_with(httpx.ReadTimeout("slow"), (200, REPLY))
    client.complete(MESSAGES, SCHEMA)
    assert len(sent) == 2


def test_gives_up_after_the_last_retry() -> None:
    waits: list[float] = []
    client, sent = client_with(500, 502, 503, 504, waits=waits)
    with pytest.raises(UnavailableError, match="after 4 tries"):
        client.complete(MESSAGES, SCHEMA)
    assert len(sent) == 4
    assert waits == [1.0, 2.0, 4.0]


# ---------------------------------------------------------------- errors that stop at once


@pytest.mark.parametrize("status", [401, 403])
def test_rejected_key_stops_immediately(status: int) -> None:
    client, sent = client_with(status, (200, REPLY))
    with pytest.raises(RejectedKeyError) as caught:
        client.complete(MESSAGES, SCHEMA)
    assert len(sent) == 1  # no retry
    assert FAKE_KEY not in str(caught.value)  # the key is never echoed


def test_other_client_errors_are_not_retried() -> None:
    client, sent = client_with((400, {"error": "bad schema"}), (200, REPLY))
    with pytest.raises(LLMError, match="HTTP 400"):
        client.complete(MESSAGES, SCHEMA)
    assert len(sent) == 1


# ---------------------------------------------------------------- the API key


def test_key_from_the_environment(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv(KEY_VAR, " env-key ")
    assert load_api_key(tmp_path / ".env") == "env-key"


def test_key_from_a_dotenv_file(tmp_path: Path) -> None:
    env_file = tmp_path / ".env"
    env_file.write_text(f'# local secrets\nOTHER=1\n{KEY_VAR}="file-key"\n', encoding="utf-8")
    assert load_api_key(env_file) == "file-key"


def test_missing_key_is_a_clear_error(tmp_path: Path) -> None:
    with pytest.raises(MissingKeyError, match=KEY_VAR):
        load_api_key(tmp_path / ".env")


# ---------------------------------------------------------------- the mock client


def test_mock_returns_answers_in_order_and_repeats_the_last() -> None:
    mock = MockClient(["first", "second"])
    answers = [mock.complete(MESSAGES, SCHEMA)[0] for _ in range(3)]
    assert answers == ["first", "second", "second"]
    assert len(mock.calls) == 3


def test_mock_can_simulate_a_failing_api() -> None:
    mock = MockClient([UnavailableError("down")])
    with pytest.raises(UnavailableError):
        mock.complete(MESSAGES, SCHEMA)
