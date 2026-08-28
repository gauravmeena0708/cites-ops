import re
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
from ..utils.config_loader import ConfigLoader
from ..utils.helpers import normalise_text, text_sha256, extract_workflow_level

class IssueClassifier:
    """
    Deterministic rule-based classifier for issue tracker tickets.
    Driven by external rules.yaml without hardcoded patterns.
    """

    def __init__(self, rules_path: Optional[str] = None):
        self.rules_config = ConfigLoader.get_rules(rules_path)
        self.version = self.rules_config.get("version", "1.0.0")
        self.major_catalog = self.rules_config.get("major_categories", {})
        self.compiled_rules = self._compile_rules(self.rules_config.get("rules", []))

    def _compile_rules(self, raw_rules: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        compiled = []
        for r in raw_rules:
            patterns = [re.compile(p, re.IGNORECASE | re.DOTALL) for p in r.get("patterns", [])]
            compiled.append({
                "rule_id": r.get("rule_id", "C99_OTHER"),
                "topic_key": r.get("topic_key", "other"),
                "major_topic_key": r.get("major_topic_key", "other"),
                "label": r.get("label", "Other"),
                "description": r.get("description", ""),
                "priority": int(r.get("priority", 0)),
                "patterns": patterns,
            })
        # Sort by priority descending
        compiled.sort(key=lambda x: x["priority"], reverse=True)
        return compiled

    def classify_issue(self, summary: str, description: str) -> Dict[str, Any]:
        """
        Classify a single issue by its summary and description.
        Returns a rich classification dictionary.
        """
        summary_clean = str(summary or "").strip()
        desc_clean = str(description or "").strip()
        combined_text = f"{summary_clean}\n{desc_clean}".strip()
        hash_val = text_sha256(f"{summary_clean} {desc_clean}")

        wf_key, wf_label = extract_workflow_level(combined_text)

        best_rule: Optional[Dict[str, Any]] = None
        best_score = -1
        best_match = ""

        for rule in self.compiled_rules:
            # Early break: remaining rules cannot beat best_score
            if rule["priority"] + 20 <= best_score:
                break

            for pattern in rule["patterns"]:
                sum_match = pattern.search(summary_clean)
                if sum_match:
                    score = rule["priority"] + 20
                    if score > best_score:
                        best_score = score
                        best_rule = rule
                        best_match = sum_match.group(0)
                        break

            # Only search description if rule can still beat best_score
            if (rule["priority"] + 2 > best_score) and (not best_rule or best_rule != rule):
                for pattern in rule["patterns"]:
                    desc_match = pattern.search(desc_clean)
                    if desc_match:
                        score = rule["priority"] + 2
                        if score > best_score:
                            best_score = score
                            best_rule = rule
                            best_match = desc_match.group(0)
                            break

        if best_rule is None:
            # Fallback to last rule (C99_OTHER)
            best_rule = self.compiled_rules[-1] if self.compiled_rules else {
                "rule_id": "C99_OTHER",
                "topic_key": "other",
                "major_topic_key": "other",
                "label": "Other / Insufficient Detail",
                "description": "Issue does not match predefined rules.",
            }

        major_meta = self.major_catalog.get(best_rule["major_topic_key"], {})
        major_label = major_meta.get("label", best_rule["major_topic_key"].replace("_", " ").title())

        return {
            "rule_id": best_rule["rule_id"],
            "topic_key": best_rule["topic_key"],
            "topic_label": best_rule["label"],
            "major_topic_key": best_rule["major_topic_key"],
            "major_topic_label": major_label,
            "category_description": best_rule["description"],
            "workflow_level_key": wf_key,
            "workflow_level_label": wf_label,
            "classifier_version": self.version,
            "text_sha256": hash_val,
            "matched_text": best_match,
        }

    def classify_dataframe(
        self,
        df: pd.DataFrame,
        summary_col: str = "Summary",
        desc_col: str = "Description",
    ) -> pd.DataFrame:
        """
        Classify all rows in a DataFrame and append classification columns.
        """
        results = []
        for _, row in df.iterrows():
            summary = row.get(summary_col, "")
            desc = row.get(desc_col, "")
            results.append(self.classify_issue(summary, desc))

        res_df = pd.DataFrame(results)
        return pd.concat([df.reset_index(drop=True), res_df.reset_index(drop=True)], axis=1)
