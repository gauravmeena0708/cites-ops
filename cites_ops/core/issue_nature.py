"""Local, auditable issue-nature classification; no model or network dependency."""

import hashlib
import json
import re
import unicodedata
from pathlib import Path

import yaml


OTHER = "Other / needs review"


def normalize_key(value):
    return re.sub(r"[^a-z0-9]", "", unicodedata.normalize("NFKC", str(value)).casefold())


class IssueNatureClassifier:
    """One primary nature per ticket, with optional functionality scoping/exclusions.

    The existing rules catalogue supplies symptom labels. Broad major categories
    are deliberately absent from the result. Cache keys include rule content,
    functionality and exact normalized text, never only the ticket ID.
    """

    ENGINE_VERSION = "nature-3"

    def __init__(self, rules_path=None, cache=None):
        path = Path(rules_path) if rules_path else Path(__file__).parents[1] / "config/rules.yaml"
        raw = path.read_bytes()  # Explicit missing custom paths must fail.
        config = yaml.safe_load(raw)
        self.fingerprint = hashlib.sha256(self.ENGINE_VERSION.encode() + raw).hexdigest()
        self.version = config.get("version", "unversioned")
        self.rules = []
        for rule in config["rules"]:
            if not rule.get("patterns"):
                continue
            compiled = dict(rule)
            for field in ("patterns", "exclude_patterns"):
                compiled[field] = [re.compile(p, re.I | re.S) for p in rule.get(field, [])]
            compiled["functionalities"] = {normalize_key(v) for v in rule.get("functionalities", [])}
            self.rules.append(compiled)
        self.rules.sort(key=lambda r: -int(r.get("priority", 0)))
        self.cache = cache if cache is not None else {}
        self.hits = 0
        self.misses = 0

    @staticmethod
    def clean(value):
        return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(value or ""))).strip().casefold()

    def classify(self, functionality, summary, description):
        summary, description = self.clean(summary), self.clean(description)
        functionality = normalize_key(functionality)
        key = hashlib.sha256(json.dumps(
            [self.fingerprint, functionality, summary, description], ensure_ascii=False
        ).encode()).hexdigest()
        if key in self.cache:
            self.hits += 1
            return dict(self.cache[key])
        self.misses += 1
        best, score, match_field = None, -1, ""
        for rule in self.rules:
            priority = int(rule.get("priority", 0))
            if priority + 20 <= score:
                break
            if rule["functionalities"] and functionality not in rule["functionalities"]:
                continue
            if any(p.search(summary + "\n" + description) for p in rule["exclude_patterns"]):
                continue
            for text, bonus in ((summary, 20), (description, 2)):
                if priority + bonus > score and any(p.search(text) for p in rule["patterns"]):
                    best, score = rule, priority + bonus
                    match_field = "summary" if bonus == 20 else "description"
        result = {
            "nature": best.get("nature_label", best["label"]) if best else OTHER,
            "rule_id": best["rule_id"] if best else "C99_OTHER",
            "match_field": match_field,
        }
        self.cache[key] = result
        return dict(result)
