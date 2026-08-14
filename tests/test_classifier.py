import unittest
from cites_ops.core.classifier import IssueClassifier

class TestClassifier(unittest.TestCase):
    def setUp(self):
        self.classifier = IssueClassifier()

    def test_classifier_initialization(self):
        self.assertEqual(self.classifier.version, "issue-categories-v3.2.0")
        self.assertGreater(len(self.classifier.compiled_rules), 0)

    def test_visibility_classification(self):
        res = self.classifier.classify_issue(
            summary="Form-13 Claim not visible at DA level",
            description="Claim inwarded but not appearing in Dealing Assistant login inbox."
        )
        self.assertEqual(res["rule_id"], "C01_VISIBILITY_DA")
        self.assertEqual(res["topic_key"], "claim_not_visible_da")
        self.assertEqual(res["workflow_level_key"], "da")

    def test_identity_kyc_classification(self):
        res = self.classifier.classify_issue(
            summary="Unable to link UAN with Aadhaar",
            description="Aadhaar verification failed due to demographic mismatch."
        )
        self.assertEqual(res["rule_id"], "C25_IDENTITY_KYC")
        self.assertEqual(res["major_topic_key"], "identity_kyc_correction")

    def test_deterministic_checksum(self):
        res1 = self.classifier.classify_issue("Gateway timeout 504 error", "Unable to login to portal")
        res2 = self.classifier.classify_issue("Gateway timeout 504 error", "Unable to login to portal")
        self.assertEqual(res1["text_sha256"], res2["text_sha256"])
        self.assertEqual(res1["rule_id"], res2["rule_id"])

if __name__ == "__main__":
    unittest.main()
