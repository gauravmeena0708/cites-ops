import unittest
import pandas as pd
from cites_ops.core.workforce import WorkforceMapper

class TestWorkforce(unittest.TestCase):
    def test_workforce_hierarchy_mapping(self):
        df_issues = pd.DataFrame([
            {"Id": "1", "Category": "Form-13", "Status": "open", "Assigned To": "team_epfo_form_13"},
            {"Id": "2", "Category": "Form-13", "Status": "resolved", "Assigned To": "team_epfo_form_13"},
            {"Id": "3", "Category": "Form-10D", "Status": "open", "Assigned To": "team_cdac_pension"},
        ])
        df_teams = pd.DataFrame([
            {"Team": "Form-13", "Account handled by": "Officer A", "DD(IS)": "DD 1", "JD(IS)": "JD 1"},
            {"Team": "Form-10D", "Account handled by": "Officer B", "DD(IS)": "DD 2", "JD(IS)": "JD 2"},
        ])

        mapper = WorkforceMapper()
        res = mapper.process_workload(df_issues, df_teams)

        self.assertEqual(res["kpis"]["total_issues"], 3)
        self.assertEqual(res["kpis"]["open_issues"], 2)
        self.assertEqual(res["kpis"]["resolved_issues"], 1)
        self.assertEqual(len(res["tree"]), 2)

if __name__ == "__main__":
    unittest.main()
