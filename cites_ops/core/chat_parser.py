import io
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd
from .entity_matcher import EntityMatcher

class ChatParser:
    """
    Parses WhatsApp chat transcripts (.txt or .zip) into structured messages.
    Extracts timestamps, senders, bodies, and detects problem-solution discussions.
    """

    # Common WhatsApp timestamp regex patterns (12h, 24h, brackets or no brackets)
    _PATTERNS = [
        re.compile(r"^\[?(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[APap][Mm])?)\]?\s*[-:]?\s*([^:]+?):\s*(.*)$"),
        re.compile(r"^(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[APap][Mm])?)\s*-\s*([^:]+?):\s*(.*)$"),
    ]

    @classmethod
    def parse_file(cls, file_path: Union[str, Path]) -> pd.DataFrame:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Chat file not found: {path}")

        lines: List[str] = []
        if path.suffix.lower() == ".zip":
            with zipfile.ZipFile(path, "r") as zf:
                for filename in zf.namelist():
                    if filename.endswith(".txt") and not filename.startswith("__MACOSX"):
                        with zf.open(filename) as f:
                            text = f.read().decode("utf-8", errors="replace")
                            lines.extend(text.splitlines())
                        break
        else:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                lines = f.read().splitlines()

        return cls.parse_lines(lines)

    @classmethod
    def parse_lines(cls, lines: List[str]) -> pd.DataFrame:
        records: List[Dict[str, Any]] = []
        current: Optional[Dict[str, Any]] = None

        for line in lines:
            line_str = line.strip("\u200e\u200f").strip()
            if not line_str:
                continue

            matched = False
            for pat in cls._PATTERNS:
                m = pat.match(line_str)
                if m:
                    if current:
                        records.append(current)
                    date_str, time_str, sender, msg = m.groups()
                    current = {
                        "date_str": date_str.strip(),
                        "time_str": time_str.strip(),
                        "sender": sender.strip(),
                        "message": msg.strip(),
                    }
                    matched = True
                    break

            if not matched and current is not None:
                # Multiline message continuation
                current["message"] += "\n" + line_str

        if current:
            records.append(current)

        df = pd.DataFrame(records)
        return df


class ChatKnowledgeExtractor:
    """
    Scans parsed chat messages, extracts referenced tracking IDs/UANs,
    identifies technical guidance/resolutions, and builds sanitized Knowledge Notes.
    """

    def __init__(self, entity_matcher: Optional[EntityMatcher] = None):
        self.matcher = entity_matcher or EntityMatcher()

        # Keywords indicating a solution or technical guidance
        self.solution_pattern = re.compile(
            r"(?i)\b(?:please check|kindly check|resolved|done|cleared|updated|corrected|solved|"
            r"issue is resolved|solution|try now|please verify|resettled|processed|cad generated|"
            r"role provided|mapping updated|restarted|now available|script executed)\b"
        )
        self.problem_pattern = re.compile(
            r"(?i)\b(?:error|unable to|not working|failed|problem|issue|not visible|exception|mismatch)\b"
        )

    def extract_knowledge_items(
        self,
        df_chats: pd.DataFrame,
        issue_index: Optional[Dict[tuple, List[str]]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Identify problem-solution pairs from chat logs and mask PII.
        """
        knowledge_items: List[Dict[str, Any]] = []
        if df_chats.empty:
            return knowledge_items

        for idx, row in df_chats.iterrows():
            msg = str(row.get("message", ""))
            sender = str(row.get("sender", ""))
            date_str = str(row.get("date_str", ""))

            # Extract referenced entities (UAN, Member ID, etc.)
            entities = self.matcher.extract_from_text(msg)
            is_solution = bool(self.solution_pattern.search(msg))
            is_problem = bool(self.problem_pattern.search(msg))

            # Cross-reference with tracker issues if index provided
            linked_issues = set()
            if issue_index:
                for ent in entities:
                    key = (ent["entity_type"], ent["value"])
                    if key in issue_index:
                        linked_issues.update(issue_index[key])

            if is_solution or linked_issues:
                sanitized_msg = self.matcher.mask_pii(msg)
                knowledge_items.append({
                    "chat_index": idx,
                    "date": date_str,
                    "sender": sender,
                    "is_solution": is_solution,
                    "is_problem": is_problem,
                    "entities": entities,
                    "linked_issues": sorted(list(linked_issues)),
                    "original_message": msg,
                    "sanitized_message": sanitized_msg,
                })

        return knowledge_items
