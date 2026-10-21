"""Minimal OpenRouter client: verified open-weights models, JSON output, bounded retries."""

import json
import os
from pathlib import Path
import time
from time import perf_counter
import urllib.error
import urllib.request

from .events import redact


MODEL = "qwen/qwen3-30b-a3b-instruct-2507"
VERIFIED_PROVIDERS = {MODEL: "nebius", "mistralai/mistral-small-3.2-24b-instruct": "venice"}
TIMEOUT = 180
RETRYABLE = {408, 429, 500, 502, 503, 504}
RETRIES = 2  # small fixed retry count for transient limits; an outage still stops the build.


def backoff(retry, retry_after):
    try:
        seconds = float(retry_after)
    except (TypeError, ValueError):
        seconds = -1
    return min(seconds, 30) if seconds >= 0 else 5 * 2 ** retry


def settings():
    # This accepts only the two documented simple KEY=value settings, not shell code.
    path = Path(".env")
    if path.is_file():
        for line in path.read_text().splitlines():
            key, sep, value = line.partition("=")
            if sep and key.strip() in {"OPENROUTER_API_KEY", "OPENROUTER_MODEL"}:
                os.environ.setdefault(key.strip(), value.strip().strip("\"'"))
    key = os.environ.get("OPENROUTER_API_KEY", "")
    model = os.environ.get("OPENROUTER_MODEL") or MODEL
    if not key:
        raise ValueError("Set OPENROUTER_API_KEY in the environment or local .env.")
    if model not in VERIFIED_PROVIDERS:
        raise ValueError("Choose a documented, verified open-weights model below 35B parameters.")
    return key, model


class OpenRouter:
    def __init__(self, key, model=MODEL, prompt_version="", max_tokens=1200, schema=None):
        if model not in VERIFIED_PROVIDERS:
            raise ValueError("Unverified model for this increment.")
        if type(max_tokens) is not int or not 1 <= max_tokens <= 8000:
            raise ValueError("Output token budget must be between one and 8000.")
        self.key, self.model = key, model
        self.prompt_version, self.max_tokens = prompt_version, max_tokens
        self.schema = schema

    def complete(self, messages, events, attempt):
        body = {"model": self.model, "messages": messages, "temperature": 0,
                "max_tokens": self.max_tokens, "response_format": {"type": "json_object"},
                # Preferred route first; fallbacks only to providers supporting every parameter (JSON schema).
                "provider": {"order": [VERIFIED_PROVIDERS[self.model]], "require_parameters": True, "allow_fallbacks": True}}
        if self.schema is not None:
            body["response_format"] = {"type": "json_schema", "json_schema": {
                "name": "workflow_proposal", "strict": True, "schema": self.schema}}
        encoded = json.dumps(body)
        if self.key in encoded or "sk-or-v1-" in encoded:
            raise ValueError("Credentials found in model context; request refused.")
        events.emit("model_request", attempt=attempt, call_id=str(attempt),
                    prompt_version=self.prompt_version, request=body)
        request = urllib.request.Request(
            "https://openrouter.ai/api/v1/chat/completions", data=encoded.encode(),
            headers={"Authorization": f"Bearer {self.key}", "Content-Type": "application/json"},
        )
        for retry in range(RETRIES + 1):
            started = perf_counter()
            try:
                with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
                    data = json.load(response)
                if "error" in data:
                    raise ValueError("Provider returned an error response.")
                content = data["choices"][0]["message"]["content"]
                if not isinstance(content, str):
                    raise ValueError("Provider did not return textual JSON content.")
                events.emit("model_response", attempt=attempt, call_id=str(attempt),
                            model=data.get("model"), provider=data.get("provider"), content=content,
                            finish_reason=data["choices"][0].get("finish_reason"),
                            usage=data.get("usage"),
                            cost=(data.get("usage") or {}).get("cost"),
                            elapsed_seconds=perf_counter() - started, retries=retry)
                return content
            except urllib.error.HTTPError as error:
                detail = "" if error.fp is None else redact(error.read(4096).decode("utf-8", errors="replace"), (self.key,))
                transient = error.code in RETRYABLE and retry < RETRIES
                events.emit("model_error", attempt=attempt, call_id=str(attempt), status=error.code,
                            provider_error=detail, retry_after=error.headers.get("Retry-After"),
                            elapsed_seconds=perf_counter() - started, retrying=transient)
                if not transient:
                    raise ValueError(f"OpenRouter HTTP {error.code} after {retry + 1} tries.") from None
                wait = backoff(retry, error.headers.get("Retry-After"))
            except (TimeoutError, urllib.error.URLError) as error:
                transient = retry < RETRIES
                events.emit("model_error", attempt=attempt, call_id=str(attempt), error=redact(str(error), (self.key,)),
                            elapsed_seconds=perf_counter() - started, retrying=transient)
                if not transient:
                    raise ValueError(f"OpenRouter unreachable after {retry + 1} tries: {error}") from None
                wait = backoff(retry, None)
            except Exception as error:
                message = redact(str(error), (self.key,))
                events.emit("model_error", attempt=attempt, call_id=str(attempt),
                            error=message, elapsed_seconds=perf_counter() - started)
                raise ValueError(message) from None
            time.sleep(wait)
