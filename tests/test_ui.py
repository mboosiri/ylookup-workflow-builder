"""Chat routing keeps every source when a learned workflow is run."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

try:
    from streamlit.testing.v1 import AppTest
except ImportError:
    AppTest = None


@unittest.skipIf(AppTest is None, "install the ui extra to test the chat")
class ChatTest(unittest.TestCase):
    def test_run_keeps_workbook_sources_with_or_without_a_pdf(self):
        ui = Path(__file__).resolve().parents[1] / "src/ylookup/ui.py"
        for with_pdf in (False, True):
            with self.subTest(with_pdf=with_pdf), TemporaryDirectory() as folder:
                root = Path(folder)
                source = root / "lookup.xlsx"
                pdf = root / "statement.pdf"
                original, target, workbook = [root / name for name in ("original.xlsx", "target.xlsx", "new.xlsx")]
                sources = [source, pdf] if with_pdf else [source]
                found = {"input": original, "target": target, "sources": sources,
                         "context": [], "workflows": [], "questions": []}
                script = ui.read_text().replace('RUNS = Path("runs")', f'RUNS = Path({str(root)!r})')
                with patch("ylookup.builder.guess_roles", return_value=found), \
                     patch("ylookup.adapt.adapt_workflow", return_value={"status": "success"}) as adapt:
                    app = AppTest.from_string(script).run()
                    app.session_state.files = [original, target, *sources, workbook]
                    app.session_state.built_from = {original, target, *sources}
                    app.session_state.workflow = root / "workflow.json"
                    app.session_state.folder = root
                    app.chat_input[0].set_value("run").run()
                    self.assertFalse(list(app.exception))
                    adapt.assert_called_once()
                    self.assertEqual(adapt.call_args.args[1], workbook)
                    self.assertEqual(adapt.call_args.args[2], sources)

    def chat(self, root, found, message, **session):
        """The chat after one message, with file roles fixed and no model or workflow code run."""
        ui = Path(__file__).resolve().parents[1] / "src/ylookup/ui.py"
        script = ui.read_text().replace('RUNS = Path("runs")', f'RUNS = Path({str(root)!r})')
        with patch("ylookup.builder.guess_roles", return_value=found), \
             patch("ylookup.model.settings", return_value=("key", "model")), patch("ylookup.model.OpenRouter"), \
             patch("ylookup.builder.build_workflow", return_value={"status": "stopped"}) as build, \
             patch("ylookup.adapt.adapt_workflow", return_value={"status": "success"}) as adapt:
            app = AppTest.from_string(script).run()
            app.session_state.folder = root
            for key, value in session.items():
                setattr(app.session_state, key, value)
            app.chat_input[0].set_value(message).run()
            self.assertFalse(list(app.exception))
            return app, build, adapt

    def roles(self, root, **found):
        return {"input": None, "target": None, "sources": [], "context": [], "workflows": [], "questions": [], **found}

    def test_a_pasted_description_that_says_build_is_kept_and_starts_the_build(self):
        brief = (Path(__file__).resolve().parents[1] / "examples/reconciliation-context.txt").read_text().strip()
        with TemporaryDirectory() as folder:
            root = Path(folder)
            found = self.roles(root, input=root / "input.xlsx", target=root / "target.xlsx")
            app, _, adapt = self.chat(root, found, brief, files=[found["input"], found["target"]])
            self.assertEqual(app.session_state.description, [brief])
            self.assertEqual(next(root.glob("chat-build-*/context.txt")).read_text(), brief)
            adapt.assert_not_called()

    def test_a_description_mentioning_truncated_names_does_not_start_a_run(self):
        text = "Bank references are a mess; sometimes the name is truncated, so match on partial names when needed."
        with TemporaryDirectory() as folder:
            root = Path(folder)
            original, workbook = root / "original.xlsx", root / "new.xlsx"
            found = self.roles(root, input=original)
            app, _, adapt = self.chat(root, found, text, files=[original, workbook], built_from={original},
                                      workflow=root / "workflow.json")
            adapt.assert_not_called()
            self.assertEqual(app.session_state.description, [text])

    def test_an_out_of_scope_request_is_declined_and_not_kept(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            found = self.roles(root, input=root / "input.xlsx", target=root / "target.xlsx")
            app, _, _ = self.chat(root, found, "make an image or illustration out of this, please, using every row in it",
                                  files=[found["input"], found["target"]])
            self.assertEqual(app.session_state.description, [])
            self.assertIn("I can only learn and run spreadsheet workflows", app.session_state.messages[-1]["text"])
            self.assertFalse(list(root.glob("chat-build-*")))

    def test_a_lone_workbook_or_two_unrelated_ones_are_described_honestly(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            lone, pdf = root / "book.xlsx", root / "statement.pdf"
            app, _, _ = self.chat(root, self.roles(root, input=lone, sources=[pdf]), "hello", files=[lone, pdf])
            self.assertIn("Attach the other workbook", app.session_state.messages[-1]["text"])
            self.assertIn("the input or the completed example", "".join(m.value for m in app.sidebar.markdown))
            first, second = root / "a.xlsx", root / "b.xlsx"
            app, _, _ = self.chat(root, self.roles(root, sources=[first, second]), "hello", files=[first, second])
            self.assertIn("cannot tell which is the completed example", app.session_state.messages[-1]["text"])

    def test_short_commands_still_work(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            found = self.roles(root, input=root / "input.xlsx", target=root / "target.xlsx")
            app, _, _ = self.chat(root, found, "Build it now", files=[found["input"], found["target"]])
            self.assertTrue(list(root.glob("chat-build-*")))
            self.assertEqual(app.session_state.description, [])


if __name__ == "__main__":
    unittest.main()
