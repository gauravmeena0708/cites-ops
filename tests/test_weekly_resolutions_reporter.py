import unittest
from datetime import date

import pandas as pd

from cites_ops.reporters.interactive_topics_reporter import InteractiveTopicsReporter
from cites_ops.reporters.weekly_resolutions_reporter import WeeklyResolutionsReporter


class TestWeeklyResolutionsReporter(unittest.TestCase):
    def test_weekday_values_come_from_snapshot_deltas(self):
        issues = pd.DataFrame(
            [
                {
                    "Id": "1",
                    "Category": "Form-13",
                    "Status": "resolved",
                    "Assigned To": "team_epfo_form_13",
                }
            ]
        )
        teams = pd.DataFrame(
            [
                {
                    "Team": "Form-13",
                    "Account handled by": "Handler A",
                    "DD(IS)": "Deputy A",
                    "JD(IS)": "Joint A",
                }
            ]
        )
        sources = [
            {
                "source": {"data_date": "2026-08-27"},
                "totals": {"resolved": 5, "closed": 1},
                "categories": [{"module_key": "form13", "resolved": 5, "closed": 1}],
            },
            {
                "source": {"data_date": "2026-08-28"},
                "totals": {"resolved": 7, "closed": 2},
                "categories": [{"module_key": "form13", "resolved": 7, "closed": 2}],
            },
        ]

        week = WeeklyResolutionsReporter._build_week_data(
            date(2026, 8, 24),
            issues,
            teams,
            stats_sources=sources,
            report_date="2026-08-28",
        )

        self.assertEqual(week["totals"]["2026-08-28"], 3)
        self.assertEqual(week["status"]["resolved_since_previous"], 3)
        team_node = week["hierarchy"][0]["children"][0]["children"][0]["children"][0]
        self.assertEqual(team_node["days"]["2026-08-28"], 3)
        self.assertEqual(week["dates"][4]["from_date"], "2026-08-27")

    def test_expandable_layout_is_width_constrained(self):
        weekly_css = WeeklyResolutionsReporter.HTML_TEMPLATE
        topics_css = InteractiveTopicsReporter.HTML_TEMPLATE
        self.assertIn(".children{margin-left:0", weekly_css)
        self.assertIn("max-width:1480px", weekly_css)
        self.assertIn(".tree-children{margin-left:0", topics_css)
        self.assertIn("overflow-wrap:anywhere", topics_css)


if __name__ == "__main__":
    unittest.main()
