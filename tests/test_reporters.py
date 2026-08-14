import unittest
import tempfile
from pathlib import Path
import pandas as pd
from cites_ops.reporters.excel_reporter import ExcelReporter
from cites_ops.reporters.pptx_reporter import PPTXReporter
from cites_ops.reporters.html_reporter import HTMLReporter
from cites_ops.reporters.docx_reporter import DocxReporter

class TestReporters(unittest.TestCase):
    def test_all_reporters(self):
        df = pd.DataFrame([
            {
                "Id": "001",
                "Category": "Form-13",
                "Status": "open",
                "Assigned To": "team_epfo_form_13",
                "Summary": "Form 13 issue",
                "Description": "Details",
                "major_topic_label": "Claim/task is not visible or routed",
                "topic_label": "At DA level",
                "rule_id": "C01_VISIBILITY_DA",
                "workflow_level_label": "DA",
                "category_description": "Claim missing at DA login",
                "age_days": 10,
            }
        ])

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)

            # Excel
            xl = tmp_path / "test.xlsx"
            ExcelReporter.generate_report(df, xl)
            self.assertTrue(xl.exists() and xl.stat().st_size > 0)

            # PPTX
            pptx = tmp_path / "test.pptx"
            PPTXReporter.generate_presentation(df, pptx)
            self.assertTrue(pptx.exists() and pptx.stat().st_size > 0)

            # HTML
            html = tmp_path / "test.html"
            HTMLReporter.generate_html(df, html)
            self.assertTrue(html.exists() and html.stat().st_size > 0)

            # DOCX
            docx = tmp_path / "test.docx"
            DocxReporter.generate_govt_note([{"title": "Test Category", "description": "Test Desc"}], docx)
            self.assertTrue(docx.exists() and docx.stat().st_size > 0)

if __name__ == "__main__":
    unittest.main()
