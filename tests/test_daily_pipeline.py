import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from cites_ops.pipeline import (
    DEFAULT_REPORTS,
    parse_report_date,
    parse_report_selection,
    run_daily_pipeline,
)


class TestDailyPipeline(unittest.TestCase):
    def test_strict_report_date_and_selection(self):
        self.assertEqual(parse_report_date("2026-08-28"), date(2026, 8, 28))
        self.assertEqual(parse_report_selection("default"), set(DEFAULT_REPORTS))
        self.assertEqual(parse_report_selection("status,dashboard"), {"status", "dashboard"})
        with self.assertRaises(ValueError):
            parse_report_date("28-08-2026")
        with self.assertRaises(ValueError):
            parse_report_selection("unknown")

    @patch("cites_ops.pipeline.MantisClient.fetch_issues")
    @patch("cites_ops.pipeline.MantisClient.test_connection")
    def test_pipeline_publishes_dated_standardized_output(self, test_connection, fetch_issues):
        test_connection.return_value = {"id": 1, "name": "reporter"}
        fetch_issues.return_value = [
            {
                "id": 101,
                "project": {"name": "CITES"},
                "reporter": {"name": "ro.agra"},
                "handler": {"name": "team_epfo_form_13"},
                "category": {"name": "Form-13"},
                "created_at": "2026-08-20T09:00:00",
                "updated_at": "2026-08-27T10:00:00",
                "summary": "Older pagination copy",
                "description": "Older version of the same issue",
                "status": {"name": "assigned"},
                "resolution": {"name": "open"},
            },
            {
                "id": 101,
                "project": {"name": "CITES"},
                "reporter": {"name": "ro.agra"},
                "handler": {"name": "team_epfo_form_13"},
                "category": {"name": "Form-13"},
                "created_at": "2026-08-20T09:00:00",
                "updated_at": "2026-08-28T10:00:00",
                "summary": "Claim not visible at DA level",
                "description": "Claim is not showing at dealing assistant login",
                "status": {"name": "assigned"},
                "resolution": {"name": "open"},
            }
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_dir = root / "input"
            output_root = root / "output"
            input_dir.mkdir()
            (input_dir / "issue_teams.csv").write_text(
                "Team,Account handled by,DD(IS),JD(IS)\n"
                "Form-13,Handler A,Deputy A,Joint A\n",
                encoding="utf-8",
            )

            result = run_daily_pipeline(
                report_date=date(2026, 8, 28),
                input_dir=input_dir,
                output_root=output_root,
                base_url="https://mantis.example/api/rest",
                token="secret",
                reports={"issues_csv"},
            )

            self.assertEqual(result.output_dir.name, "2026-08-28")
            self.assertTrue((result.output_dir / "cites_issues_2026-08-28.csv").is_file())
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["report_date"], "2026-08-28")
            self.assertEqual(manifest["source"]["issue_count"], 1)
            self.assertEqual(manifest["quality"]["duplicate_issue_ids_after_normalization"], 0)
            self.assertEqual(manifest["quality"]["duplicate_issue_ids_returned"], 1)
            self.assertEqual(manifest["quality"]["duplicate_rows_removed"], 1)
            self.assertEqual(
                manifest["artifacts"]["issues_csv"]["file"],
                "cites_issues_2026-08-28.csv",
            )


if __name__ == "__main__":
    unittest.main()
