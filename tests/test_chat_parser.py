import unittest
import pandas as pd
from cites_ops.core.chat_parser import ChatParser, ChatKnowledgeExtractor

class TestChatParser(unittest.TestCase):
    def test_chat_line_parsing(self):
        lines = [
            "[04/08/2026, 10:15:30] RO Delhi: Unable to process Form 10D for UAN 100123456789.",
            "[04/08/2026, 10:22:15] Tech Team: Please check now, script executed and ledger updated.",
        ]
        df = ChatParser.parse_lines(lines)
        self.assertEqual(len(df), 2)
        self.assertEqual(df.iloc[0]["sender"], "RO Delhi")
        self.assertIn("100123456789", df.iloc[0]["message"])

    def test_knowledge_extraction_and_masking(self):
        lines = [
            "[04/08/2026, 10:15:30] RO Delhi: Unable to process Form 10D for UAN 100123456789.",
            "[04/08/2026, 10:22:15] Tech Team: Resolved. Mapping updated for UAN 100123456789.",
        ]
        df = ChatParser.parse_lines(lines)
        extractor = ChatKnowledgeExtractor()
        items = extractor.extract_knowledge_items(df)
        self.assertGreaterEqual(len(items), 1)
        self.assertNotIn("100123456789", items[0]["sanitized_message"])

if __name__ == "__main__":
    unittest.main()
