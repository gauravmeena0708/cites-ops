import unittest
import pandas as pd
from cites_ops.core.entity_matcher import EntityMatcher

class TestEntityMatcher(unittest.TestCase):
    def setUp(self):
        self.matcher = EntityMatcher()

    def test_entity_extraction(self):
        text = "Please check UAN 100123456789 and Grievance MOLBR/E/2026/00123 for member DL/CPM/12345/678"
        entities = self.matcher.extract_from_text(text)
        types = {e["entity_type"] for e in entities}
        self.assertIn("uan", types)
        self.assertIn("grievance_id", types)
        self.assertIn("member_id", types)

    def test_pii_masking(self):
        text = "Member UAN is 100123456789 and mobile is 9876543210"
        masked = self.matcher.mask_pii(text)
        self.assertNotIn("100123456789", masked)
        self.assertNotIn("9876543210", masked)
        self.assertIn("6789", masked)

    def test_issue_index_construction(self):
        df = pd.DataFrame([
            {"Id": "T-101", "Summary": "Form 10D error for UAN 100123456789", "Description": "Details here"},
            {"Id": "T-102", "Summary": "Grievance MOLBR/E/2026/00999", "Description": "Member grievance"},
        ])
        idx = self.matcher.build_issue_index(df)
        self.assertIn(("uan", "100123456789"), idx)
        self.assertEqual(idx[("uan", "100123456789")], ["T-101"])
        self.assertIn(("grievance_id", "MOLBR/E/2026/00999"), idx)
        self.assertEqual(idx[("grievance_id", "MOLBR/E/2026/00999")], ["T-102"])

if __name__ == "__main__":
    unittest.main()
