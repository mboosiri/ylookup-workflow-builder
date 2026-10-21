"""Local observable events; no credentials or calendar-date logging."""

import json
from pathlib import Path
import re


def redact(text, secrets=()):
    for secret in secrets:
        if secret:
            text = text.replace(secret, "[REDACTED]")
    return re.sub(r"sk-or-v1-[A-Za-z0-9]+", "[REDACTED]", text)


class Events:
    def __init__(self, directory, secrets=()):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, mode=0o700, exist_ok=False)
        self.path = self.directory / "events.jsonl"
        self.secrets = secrets
        self.sequence = 0

    def emit(self, event, **payload):
        item = {"run_id": self.directory.name, "sequence": self.sequence + 1,
                "event": event, **payload}
        line = redact(json.dumps(item, allow_nan=False), self.secrets)
        with self.path.open("a", encoding="utf-8") as stream:
            stream.write(line + "\n")
        self.sequence += 1
