import re
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd
from ..utils.config_loader import ConfigLoader

class EntityMatcher:
    """
    Extracts entities (UAN, Member ID, Grievance ID, Tracking ID) and
    manages cross-referencing between chats and issue tracker tickets.
    Includes automated PII masking.
    """

    def __init__(self, entities_path: Optional[str] = None):
        self.config = ConfigLoader.get_entities(entities_path).get("entities", {})
        self.compiled_entities: Dict[str, Dict[str, Any]] = {}
        for entity_key, meta in self.config.items():
            pattern_str = meta.get("pattern", "")
            if pattern_str:
                self.compiled_entities[entity_key] = {
                    "label": meta.get("label", entity_key),
                    "pattern": re.compile(pattern_str, re.IGNORECASE),
                    "mask_template": meta.get("mask", "XXXX"),
                    "normalize": meta.get("normalize", "none"),
                }

    def extract_from_text(self, text: str) -> List[Dict[str, str]]:
        """
        Extract all configured entities found in a given text block.
        Returns a list of {entity_type, value, raw_value}.
        """
        if not text:
            return []
        found = []
        seen = set()

        for entity_key, meta in self.compiled_entities.items():
            for match in meta["pattern"].finditer(str(text)):
                raw_val = match.group(0).strip()
                norm_val = raw_val.upper()
                if (entity_key, norm_val) in seen:
                    continue
                seen.add((entity_key, norm_val))
                found.append({
                    "entity_type": entity_key,
                    "label": meta["label"],
                    "value": norm_val,
                    "raw_value": raw_val,
                })
        return found

    def mask_pii(self, text: str) -> str:
        """
        Mask sensitive identifiers (UANs, Mobile numbers, Member accounts) in text.
        """
        if not text:
            return ""
        masked_text = str(text)

        for entity_key, meta in self.compiled_entities.items():
            mask_template = meta["mask_template"]

            def _repl(match: re.Match) -> str:
                raw = match.group(0)
                digits = re.sub(r"\D", "", raw)
                last4 = digits[-4:] if len(digits) >= 4 else "****"
                
                # Check for template placeholders
                if "{last4}" in mask_template:
                    return mask_template.replace("{last4}", last4)
                if "{prefix}" in mask_template:
                    parts = raw.split("/")
                    prefix = parts[0] if parts else "ID"
                    return mask_template.replace("{prefix}", prefix)
                if "{region}" in mask_template:
                    parts = raw.split("/")
                    region = parts[0] if parts else "REG"
                    return mask_template.replace("{region}", region).replace("{last4}", last4)
                return "[REDACTED]"

            masked_text = meta["pattern"].sub(_repl, masked_text)

        return masked_text

    def build_issue_index(
        self,
        df_issues: pd.DataFrame,
        id_col: str = "Id",
        summary_col: str = "Summary",
        desc_col: str = "Description",
    ) -> Dict[Tuple[str, str], List[str]]:
        """
        Build an inverted index mapping (entity_type, entity_value) -> [issue_ids].
        """
        index: Dict[Tuple[str, str], List[str]] = {}

        for _, row in df_issues.iterrows():
            issue_id = str(row.get(id_col, "")).strip()
            if not issue_id:
                continue

            combined = f"{row.get(summary_col, '')} {row.get(desc_col, '')}"
            entities = self.extract_from_text(combined)

            for ent in entities:
                key = (ent["entity_type"], ent["value"])
                if key not in index:
                    index[key] = []
                if issue_id not in index[key]:
                    index[key].append(issue_id)

        return index
