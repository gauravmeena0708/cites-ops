"""Shared distinct-ticket rankings and deterministic, masked sample selection."""

import re
from collections import Counter

from .entity_matcher import EntityMatcher
from .issue_nature import OTHER


def clean_excerpt(value, matcher, limit):
    value = re.sub(r"_x000[Dd]_", " ", str(value or ""))
    value = re.sub(r"\b(password|passwd|pwd|pw|passcode|user\s*name)\s*(?::|=|\bis\b)\s*[^\s;,)]+",
                   r"\1: [masked]", value, flags=re.I)
    value = matcher.mask_pii(value)
    value = re.sub(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", "[email masked]", value)
    value = re.sub(r"\b[A-Z]{5}\d{4}[A-Z]\b", "[PAN masked]", value, flags=re.I)
    value = re.sub(r"\b[A-Z]{2,5}\d{12,}\b", "[member ID masked]", value, flags=re.I)
    value = re.sub(r"(?<!\w)(?:\+\d{1,3}[ -]?)?\d(?:[ -]?\d){9,17}(?!\w)", "[identifier masked]", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value if len(value) <= limit else value[:limit - 1].rsplit(" ", 1)[0] + "…"


def sample_tickets(rows, matcher, limit=3):
    # Prefer different functionalities/statuses and distinct symptom descriptions.
    ordered = sorted(rows, key=lambda r: (r.get("nature_match") == "summary", r["updated"], r["submitted"], r["id"]), reverse=True)
    selected, seen_text, seen_groups = [], set(), set()
    for diverse in (True, False):
        for row in ordered:
            text_key = re.sub(r"\s+", " ", row["summary"]).strip().casefold()
            group = (row["functionality"], row["status"])
            if not text_key:
                text_key = row["description"].strip().casefold() or row["id"]
            if text_key in seen_text or (diverse and group in seen_groups):
                continue
            seen_text.add(text_key)
            seen_groups.add(group)
            selected.append({
                "id": row["id"], "functionality": row["functionality"], "status": row["status"],
                "submitted": row["submitted"], "summary": clean_excerpt(row["summary"], matcher, 155),
                "description": clean_excerpt(row["description"], matcher, 245),
            })
            if len(selected) == limit:
                return selected
    return selected


def build_rankings(data, top_n=15, period="all"):
    if top_n not in (10, 15):
        raise ValueError("Top issue-nature count must be 10 or 15")
    if period not in ("all", "week"):
        raise ValueError("Ranking period must be all or week")
    matcher = EntityMatcher()
    selected = [i for i in data["issues"] if period == "all" or i["last_week"]]
    results = {}
    for scope, statuses in (("open", {"open"}), ("resolved_closed", {"resolved", "closed"})):
        rows = [i for i in selected if i["status"] in statuses]
        groups = {}
        for row in rows:
            groups.setdefault(row["nature"], []).append(row)
        ordered = sorted(((name, members) for name, members in groups.items() if name != OTHER),
                         key=lambda pair: (-len(pair[1]), pair[0]))
        ranking = []
        for rank, (name, members) in enumerate(ordered[:top_n], 1):
            funcs = Counter(i["functionality"] for i in members)
            ranking.append({"rank": rank, "nature": name, "count": len(members),
                            "resolved": sum(i["status"] == "resolved" for i in members),
                            "closed": sum(i["status"] == "closed" for i in members),
                            "functionalities": [{"name": name, "count": count} for name, count in sorted(funcs.items(), key=lambda pair: (-pair[1], pair[0]))],
                            "samples": sample_tickets(members, matcher)})
        results[scope] = {
            "rows": ranking, "total": len(rows), "review_count": len(groups.get(OTHER, [])),
            "ranked_total": sum(r["count"] for r in ranking), "defined_natures": len(ordered),
            "available": not (scope == "resolved_closed" and data["scope"] == "open"),
            "resolved": sum(i["status"] == "resolved" for i in rows),
            "closed": sum(i["status"] == "closed" for i in rows),
        }
    return {"top_n": top_n, "period": period, **results}
