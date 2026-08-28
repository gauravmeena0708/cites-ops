import unittest
from unittest.mock import patch, MagicMock
from io import BytesIO
import json
import pandas as pd
from cites_ops.core.mantis_client import MantisClient


class TestMantisClient(unittest.TestCase):

    def setUp(self):
        self.client = MantisClient(
            base_url="http://localhost:8080/mantisbt/api/rest",
            token="test_token_12345"
        )

    def test_to_dataframe_formatting(self):
        sample_issues = [
            {
                "id": 101,
                "project": {"id": 1, "name": "CITES-2.0"},
                "reporter": {"id": 2, "name": "ro.agra"},
                "handler": {"id": 4, "name": "team_unified_portal"},
                "category": {"id": 1, "name": "Form-31"},
                "created_at": "2026-08-28T09:15:00+02:00",
                "updated_at": "2026-08-28T10:00:00+02:00",
                "summary": "Advance claim settlement delay",
                "description": "Details about settlement delay in form 31",
                "status": {"id": 50, "name": "assigned"},
                "resolution": {"id": 10, "name": "open"},
                "target_version": "v2.1"
            },
            {
                "id": 102,
                "project": "CITES-2.0",
                "reporter": "ro.krpuram",
                "handler": None,
                "category": "Annual_Accounts",
                "created_at": "2026-08-27T08:00:00",
                "updated_at": "2026-08-28T08:00:00",
                "summary": "Annual account generation error",
                "description": "Error in generating accounts for 2025",
                "status": "resolved",
                "resolution": "fixed",
                "target_version": ""
            }
        ]

        df = MantisClient.to_dataframe(sample_issues)

        self.assertEqual(len(df), 2)
        expected_cols = [
            "Id", "Project", "Reporter", "Assigned To", "Category",
            "Date Submitted", "Updated", "Summary", "Status",
            "Resolution", "Fixed in Version", "Description"
        ]
        self.assertListEqual(list(df.columns), expected_cols)

        # Check row 0
        self.assertEqual(df.iloc[0]["Id"], "0000101")
        self.assertEqual(df.iloc[0]["Project"], "CITES-2.0")
        self.assertEqual(df.iloc[0]["Reporter"], "ro.agra")
        self.assertEqual(df.iloc[0]["Assigned To"], "team_unified_portal")
        self.assertEqual(df.iloc[0]["Category"], "Form-31")
        self.assertEqual(df.iloc[0]["Date Submitted"], "2026-08-28")
        self.assertEqual(df.iloc[0]["Status"], "assigned")

        # Check row 1
        self.assertEqual(df.iloc[1]["Id"], "0000102")
        self.assertEqual(df.iloc[1]["Reporter"], "ro.krpuram")
        self.assertEqual(df.iloc[1]["Category"], "Annual_Accounts")
        self.assertEqual(df.iloc[1]["Date Submitted"], "2026-08-27")
        self.assertEqual(df.iloc[1]["Status"], "resolved")

    @patch("urllib.request.urlopen")
    def test_test_connection(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps({
            "id": 1,
            "name": "administrator",
            "email": "root@localhost"
        }).encode("utf-8")
        mock_resp.__enter__.return_value = mock_resp
        mock_urlopen.return_value = mock_resp

        user_info = self.client.test_connection()
        self.assertEqual(user_info["id"], 1)
        self.assertEqual(user_info["name"], "administrator")

    @patch("urllib.request.urlopen")
    def test_fetch_issues_pagination(self, mock_urlopen):
        # Page 1
        page1_resp = MagicMock()
        page1_resp.read.return_value = json.dumps({
            "issues": [{"id": 1, "summary": "Issue 1"}, {"id": 2, "summary": "Issue 2"}]
        }).encode("utf-8")
        page1_resp.__enter__.return_value = page1_resp

        # Page 2
        page2_resp = MagicMock()
        page2_resp.read.return_value = json.dumps({
            "issues": [{"id": 3, "summary": "Issue 3"}]
        }).encode("utf-8")
        page2_resp.__enter__.return_value = page2_resp

        mock_urlopen.side_effect = [page1_resp, page2_resp]

        issues = self.client.fetch_issues(page_size=2)
        self.assertEqual(len(issues), 3)
        self.assertEqual(issues[0]["id"], 1)
        self.assertEqual(issues[2]["id"], 3)


if __name__ == "__main__":
    unittest.main()
