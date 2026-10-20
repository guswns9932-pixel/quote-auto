import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "no_com"))

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication
import app_settings
from page_quote import QuoteBuilderPage


class LocalUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_output_choice_is_persisted_and_cancel_does_not_generate(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(
            app_settings, "_st", side_effect=lambda: QSettings(str(Path(tmp) / "settings.ini"), QSettings.IniFormat)
        ):
            page = QuoteBuilderPage()
            page.state.template_path = "selected-template.xlsx"
            with patch("page_quote.QFileDialog.getExistingDirectory", return_value=""), patch("page_quote._GenQuoteThread") as worker:
                page._generate_quote()
                worker.assert_not_called()
                self.assertIsNone(page.state.output_root)
            with patch("page_quote.QFileDialog.getExistingDirectory", return_value=tmp), patch("page_quote._GenQuoteThread") as worker:
                page._generate_quote()
                worker.return_value.start.assert_called_once()
                self.assertEqual(page.state.output_root, tmp)
            page._gen_qt_thread = None
            page.close()
            reopened = QuoteBuilderPage()
            self.assertEqual(reopened.state.output_root, tmp)
            reopened.close()


if __name__ == "__main__":
    unittest.main()
