"""Talk to the model: one OpenRouter call with retries, and a mock for tests and demos.

The client sends the messages from prompts.py and asks for JSON that follows the strict
schema from schema.py. Every call records tokens, cost and latency.

Errors, from the caller's point of view:
- MissingKeyError / RejectedKeyError: stop and tell the user (retrying cannot help);
- UnavailableError: timeouts, rate limits or server errors that survived the retries;
- LLMError: any other failure. The pipeline falls back to a plain deck for these.
"""

import os
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

import httpx

API_URL = "https://openrouter.ai/api/v1/chat/completions"
KEY_VAR = "OPENROUTER_API_KEY"
DEFAULT_MODEL = "anthropic/claude-sonnet-4.5"  # confirmed against OpenRouter in the notebook
TIMEOUT_S = 60
MAX_RETRIES = 3  # so at most 4 attempts, waiting 1 s, 2 s, 4 s in between
RETRY_STATUS = {408, 429, 500, 502, 503, 504}
MAX_WAIT_S = 30  # cap on a server's Retry-After


class LLMError(Exception):
    """The model call failed."""


class MissingKeyError(LLMError):
    pass


class RejectedKeyError(LLMError):
    pass


class UnavailableError(LLMError):
    pass


@dataclass
class Usage:
    model: str
    prompt_tokens: int = 0
    completion_tokens: int = 0
    cost_usd: float = 0.0
    latency_s: float = 0.0


def load_api_key(env_file: Path = Path(".env")) -> str:
    """The key from the environment, or else from a local .env file. Never printed."""
    key = os.environ.get(KEY_VAR, "").strip()
    if not key and env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            name, sep, value = line.partition("=")
            if sep and name.strip() == KEY_VAR:
                key = value.strip().strip("\"'")
    if not key:
        raise MissingKeyError(f"No API key found. Set {KEY_VAR} in the environment or in .env.")
    return key


class OpenRouterClient:
    def __init__(
        self,
        api_key: str,
        model: str = DEFAULT_MODEL,
        transport: httpx.BaseTransport | None = None,  # tests pass a fake transport
        sleep: Callable[[float], None] = time.sleep,  # tests skip the real waiting
    ):
        self.model = model
        self._sleep = sleep
        self._http = httpx.Client(
            timeout=TIMEOUT_S, transport=transport, headers={"Authorization": f"Bearer {api_key}"}
        )

    def complete(self, messages: list[dict], schema: dict) -> tuple[str, Usage]:
        """One structured call. Returns the reply text and what the call cost."""
        body = {
            "model": self.model,
            "messages": messages,
            "temperature": 0.2,
            "response_format": {"type": "json_schema", "json_schema": schema},
            "usage": {"include": True},  # OpenRouter then reports the real cost
        }
        start = time.perf_counter()
        data = self._post(body).json()
        latency = round(time.perf_counter() - start, 2)
        if "error" in data or not data.get("choices"):
            message = (data.get("error") or {}).get("message", "no choices in the reply")
            raise LLMError(f"OpenRouter returned an error: {message}")
        usage = data.get("usage") or {}
        return data["choices"][0]["message"].get("content") or "", Usage(
            model=data.get("model", self.model),
            prompt_tokens=usage.get("prompt_tokens", 0),
            completion_tokens=usage.get("completion_tokens", 0),
            cost_usd=float(usage.get("cost") or 0),
            latency_s=latency,
        )

    def _post(self, body: dict) -> httpx.Response:
        """POST with retries on timeouts, rate limits and server errors; stop on a bad key."""
        for attempt in range(MAX_RETRIES + 1):
            response = None
            try:
                response = self._http.post(API_URL, json=body)
            except httpx.TransportError:  # includes timeouts
                problem = "timeout or network error"
            else:
                status = response.status_code
                if status in (401, 403):
                    raise RejectedKeyError(f"OpenRouter rejected the API key (HTTP {status}).")
                if status < 400:
                    return response
                if status not in RETRY_STATUS:
                    raise LLMError(f"OpenRouter returned HTTP {status}: {response.text[:200]}")
                problem = f"HTTP {status}"
            if attempt < MAX_RETRIES:
                self._sleep(_wait_seconds(response, attempt))
        raise UnavailableError(f"OpenRouter unavailable after {MAX_RETRIES + 1} tries ({problem}).")


def _wait_seconds(response: httpx.Response | None, attempt: int) -> float:
    """The server's Retry-After when it sends one, otherwise 1, 2, 4 seconds."""
    retry_after = response.headers.get("retry-after", "") if response is not None else ""
    if retry_after.isdigit():
        return min(float(retry_after), MAX_WAIT_S)
    return 2.0**attempt


@dataclass
class MockClient:
    """Canned answers instead of a model: no key, no network.

    Each call returns the next answer; the last one repeats. An answer that is an exception
    is raised instead, to simulate a failing API.
    """

    answers: list[str | Exception]
    model: str = "mock"
    calls: list[list[dict]] = field(default_factory=list)

    def complete(self, messages: list[dict], schema: dict) -> tuple[str, Usage]:
        self.calls.append(messages)
        answer = self.answers[min(len(self.calls), len(self.answers)) - 1]
        if isinstance(answer, Exception):
            raise answer
        return answer, Usage(model=self.model)
