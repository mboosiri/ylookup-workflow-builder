import io
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from urllib.error import HTTPError

from ylookup.events import Events
from ylookup.model import MODEL, OpenRouter
from ylookup.workflow import compare, execute


def workflow(operation="subtract"):
    return {"version": 1, "intent": "remaining amount", "operation": operation,
            "left": "due", "right": "paid", "output": "remaining"}


def ready(operation="subtract"):
    return json.dumps({"status": "ready", "workflow": workflow(operation)})


EXAMPLE = {"context": "Track the remaining amount for each record.",
           "input": [{"id": "001", "due": "105.10", "paid": "100.00"}],
           "target": [{"id": "001", "remaining": "5.10"}]}


class WorkflowTest(unittest.TestCase):
    def test_decimal_comparison_and_record_coverage(self):
        self.assertEqual(compare(execute(workflow(), EXAMPLE["input"]), EXAMPLE["target"]), [])
        self.assertTrue(compare([{"id": "002", "remaining": "5.10"}], EXAMPLE["target"]))
        self.assertEqual(compare([{"id": "001", "remaining": "5.099999999999999"}], EXAMPLE["target"]), [])
        self.assertTrue(compare([{"id": "001", "remaining": "5.11"}], EXAMPLE["target"]))

    def test_invalid_inputs_and_workflows_do_not_default_to_zero(self):
        for bad in [None, True, "NaN", "Infinity", "unknown"]:
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                execute(workflow(), [{"id": "001", "due": bad, "paid": "1"}])
        for rows in [[], [{"id": "001", "due": "2"}], EXAMPLE["input"] * 2]:
            with self.assertRaises(ValueError):
                execute(workflow(), rows)
        for bad in [{**workflow(), "operation": "eval"},
                    {**workflow(), "operation": []},
                    {**workflow(), "version": True},
                    {**workflow(), "output": "id"},
                    {**workflow(), "cached_answer": "5.10"}]:
            with self.assertRaises(ValueError):
                execute(bad, EXAMPLE["input"])
        with self.assertRaises(ValueError):
            execute(workflow("divide"), [{"id": "001", "due": "1", "paid": "0"}])

    def test_events_redact_and_refuse_existing_directory_or_write_failure(self):
        with TemporaryDirectory() as directory:
            events = Events(Path(directory) / "trace", ("private-value",))
            events.emit("example", content="private-value sk-or-v1-testcredential")
            item = json.loads(events.path.read_text())
            self.assertEqual(item["content"], "[REDACTED] [REDACTED]")
            self.assertEqual(item["sequence"], 1)
            with self.assertRaises(FileExistsError):
                Events(events.directory)
            with patch.object(Path, "open", side_effect=OSError("disk unavailable")):
                with self.assertRaises(OSError):
                    events.emit("must_not_claim_success")
            self.assertEqual(events.sequence, 1)

    def test_client_logs_usage_but_never_authorization_or_reasoning(self):
        response = {"model": MODEL, "choices": [{"finish_reason": "stop", "message": {
            "content": ready(), "reasoning": "private internal field"}}],
            "usage": {"prompt_tokens": 20, "completion_tokens": 10}}
        with TemporaryDirectory() as directory:
            events = Events(Path(directory) / "api", ("private-value",))
            with patch("ylookup.model.urllib.request.urlopen", return_value=io.BytesIO(json.dumps(response).encode())) as request:
                content = OpenRouter("private-value").complete(
                    [{"role": "user", "content": "example data"}], events, 1)
            self.assertEqual(json.loads(content)["status"], "ready")
            self.assertEqual(request.call_args.kwargs["timeout"], 90)
            body = json.loads(request.call_args.args[0].data)
            self.assertEqual(body["provider"]["order"], ["nebius"])
            self.assertTrue(body["provider"]["allow_fallbacks"])
            self.assertTrue(body["provider"]["require_parameters"])
            text = events.path.read_text()
            self.assertNotIn("private-value", text)
            self.assertNotIn("private internal field", text)
            self.assertNotIn("Authorization", text)
            self.assertIsNone(json.loads(text.splitlines()[-1])["cost"])

    def test_verified_alternative_model_uses_its_declared_provider(self):
        response = {"choices": [{"message": {"content": "{}"}}]}
        with TemporaryDirectory() as directory:
            with patch("ylookup.model.urllib.request.urlopen", return_value=io.BytesIO(json.dumps(response).encode())) as request:
                OpenRouter("private-value", "mistralai/mistral-small-3.2-24b-instruct", schema={"type": "object"}).complete([], Events(Path(directory) / "api"), 1)
            body = json.loads(request.call_args.args[0].data)
            self.assertEqual(body["provider"]["order"], ["venice"])
            self.assertEqual(body["response_format"]["type"], "json_schema")
            self.assertTrue(body["response_format"]["json_schema"]["strict"])
            self.assertEqual(body["model"], "mistralai/mistral-small-3.2-24b-instruct")
        with self.assertRaises(ValueError):
            OpenRouter("private-value", "unverified-model")

    def test_transient_errors_retry_with_backoff_then_stop(self):
        def limited():
            return HTTPError("url", 429, "Rate limited", {"Retry-After": "10"},
                             io.BytesIO(b'{"error":{"message":"Provider limit private-value"}}'))
        ok = {"choices": [{"message": {"content": "{}"}}]}
        with TemporaryDirectory() as directory, patch("ylookup.model.time.sleep") as sleep:
            events = Events(Path(directory) / "api", ("private-value",))
            with patch("ylookup.model.urllib.request.urlopen",
                       side_effect=[limited(), TimeoutError("read timed out"), io.BytesIO(json.dumps(ok).encode())]):
                self.assertEqual(OpenRouter("private-value").complete([], events, 1), "{}")
            self.assertEqual([c.args[0] for c in sleep.call_args_list], [10, 10])
            with patch("ylookup.model.urllib.request.urlopen", side_effect=[limited(), limited(), limited()]) as request:
                with self.assertRaisesRegex(ValueError, "HTTP 429 after 3 tries"):
                    OpenRouter("private-value").complete([], events, 2)
            self.assertEqual(request.call_count, 3)
            recorded = json.loads(events.path.read_text().splitlines()[-1])
            self.assertIn("Provider limit [REDACTED]", recorded["provider_error"])
            self.assertFalse(recorded["retrying"])

    def test_http_failure_and_context_credentials_stop_without_secret_leak(self):
        with TemporaryDirectory() as directory:
            events = Events(Path(directory) / "api", ("private-value",))
            with patch("ylookup.model.urllib.request.urlopen", side_effect=HTTPError("url", 401, "private-value", {}, None)) as request:
                with self.assertRaisesRegex(ValueError, "HTTP 401 after 1 tries"):
                    OpenRouter("private-value").complete([], events, 1)
            with self.assertRaisesRegex(ValueError, "Credentials"):
                OpenRouter("private-value").complete([{"role": "user", "content": "private-value"}], events, 2)
            self.assertNotIn("private-value", events.path.read_text())


if __name__ == "__main__":
    unittest.main()
