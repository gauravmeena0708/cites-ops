"""Prepare one snapshot and all summaries consumed by the simplified reports."""

from datetime import date, timedelta

import pandas as pd

from .issue_nature import OTHER, normalize_key
from ..utils.helpers import parse_date


STATUS_GROUPS = {
    "new": "open", "assigned": "open", "feedback": "open", "open": "open",
    "acknowledged": "open", "confirmed": "open", "reopened": "open",
    "resolved": "resolved", "fixed": "resolved", "closed": "closed",
}
COUNTS = ("total", "open", "resolved", "closed")


def counts(rows):
    result = dict.fromkeys(COUNTS, 0)
    for row in rows:
        result["total"] += 1
        result[row["status"]] += 1
    return result


def ownership_map(teams):
    required = {"Team", "Account handled by", "DD(IS)", "JD(IS)"}
    if not required.issubset(teams.columns):
        raise ValueError("Teams CSV must contain: " + ", ".join(sorted(required)))
    owners = {}
    for row in teams.fillna("").to_dict("records"):
        owner = {name: str(row[col]).strip() or "Not mapped" for name, col in (
            ("officer", "Account handled by"), ("dd", "DD(IS)"), ("jd", "JD(IS)"))}
        for functionality in str(row["Team"]).split(","):
            key = normalize_key(functionality)
            if not key:
                continue
            if key in owners and owners[key] != owner:
                raise ValueError(f"Conflicting ownership mappings for {functionality.strip()}")
            owners[key] = owner
    return owners


def _validate_stats(stats, report_date):
    if stats["source"]["data_date"] != report_date.isoformat():
        raise ValueError("Status document date must match the report date")
    for row in [stats["totals"], *stats["categories"]]:
        if any(row[k] < 0 for k in COUNTS) or row["total"] != sum(row[k] for k in COUNTS[1:]):
            raise ValueError("Status document counts do not reconcile")
    if any(sum(row[k] for row in stats["categories"]) != stats["totals"][k] for k in COUNTS):
        raise ValueError("Status document functionality counts do not reconcile with its totals")


def prepare_briefing(df, teams, classifier, report_date, scope="auto", stats=None):
    """Counts from complete CSV or reconciled status document; detail scope stays explicit."""
    required = {"Id", "Category", "Status", "Summary", "Description", "Date Submitted"}
    if not required.issubset(df.columns):
        raise ValueError("Issue CSV is missing: " + ", ".join(sorted(required - set(df.columns))))
    if df.empty:
        raise ValueError("Issue CSV contains no tickets")
    df = df.fillna("").copy()
    warnings = []
    df["Id"] = df["Id"].astype(str).str.strip().map(lambda x: str(int(x)) if x.isdigit() else x)
    if df["Id"].eq("").any():
        raise ValueError("Issue IDs must not be blank")
    # Only exact repeated rows are safely removable without guessing which copy is current.
    before = len(df)
    df = df.drop_duplicates()
    if df["Id"].duplicated().any():
        raise ValueError("Conflicting duplicate issue IDs; export a consistent snapshot")
    if before != len(df):
        warnings.append(f"Removed {before - len(df)} identical duplicate rows.")
    normalized = df["Status"].astype(str).str.strip().str.casefold()
    unknown = sorted(set(normalized) - set(STATUS_GROUPS))
    if unknown:
        raise ValueError("Unrecognized ticket statuses: " + ", ".join(unknown))
    df["_status"] = normalized.map(STATUS_GROUPS)
    if scope == "auto":
        if set(df["_status"]) == {"open"}:
            scope = "open"
        else:
            raise ValueError("CSV contains completed tickets. Use --scope all for a full export or --scope partial for a filtered export.")
    if scope == "open" and set(df["_status"]) != {"open"}:
        raise ValueError("--scope open cannot contain resolved or closed tickets")
    owners = ownership_map(teams)
    monday = report_date - timedelta(days=report_date.weekday())
    week_start, week_end = monday - timedelta(days=7), monday - timedelta(days=1)
    prior_start = week_start - timedelta(days=7)
    issues, invalid_dates = [], 0
    labels = {}
    for row in df.to_dict("records"):
        functionality = str(row["Category"]).strip() or "Unassigned"
        key = normalize_key(functionality)
        if key in labels and labels[key] != functionality:
            raise ValueError(f"Ambiguous functionality names: {labels[key]!r} and {functionality!r}")
        labels[key] = functionality
        created = parse_date(row["Date Submitted"])
        updated = parse_date(row.get("Updated", row.get("Last Update", "")))
        if created and created > report_date:
            raise ValueError(f"Issue {row['Id']} was submitted after the report date; use the matching snapshot")
        if updated and updated > report_date:
            raise ValueError(f"Issue {row['Id']} was updated after the report date; current states cannot represent that historical snapshot")
        if not created:
            invalid_dates += 1
        nature = classifier.classify(functionality, row["Summary"], row["Description"])
        issues.append({
            "id": row["Id"], "functionality": functionality, "nature": nature["nature"],
            "nature_match": nature.get("match_field", ""),
            "status": row["_status"], "tracker_status": str(row["Status"]).strip(),
            "summary": str(row["Summary"]), "description": str(row["Description"]),
            "assigned_to": str(row.get("Assigned To", "")),
            "submitted": created.isoformat() if created else "",
            "updated": updated.isoformat() if updated else "",
            "age_days": (report_date - created).days if created else None,
            "last_week": bool(created and week_start <= created <= week_end),
            "prior_week": bool(created and prior_start <= created < week_start),
            **owners.get(key, dict.fromkeys(("officer", "dd", "jd"), "Not mapped")),
        })
    if invalid_dates:
        warnings.append(f"{invalid_dates} tickets have missing/unreadable submission dates and are excluded from weekly analysis.")
    detail_counts = counts(issues)
    grouped = {}
    for issue in issues:
        grouped.setdefault(issue["functionality"], []).append(issue)
    supplied = {}
    if stats:
        _validate_stats(stats, report_date)
        for row in stats["categories"]:
            key = normalize_key(row.get("module_label", row.get("category", "")))
            if key in supplied:
                raise ValueError("Duplicate functionality in status document")
            supplied[key] = row
        for functionality, rows in grouped.items():
            actual = counts(rows)
            expected = supplied.get(normalize_key(functionality))
            if expected is None or any(actual[k] > expected[k] for k in COUNTS):
                raise ValueError(f"CSV and status document disagree for {functionality}")
            if scope in ("all", "open"):
                keys = COUNTS if scope == "all" else ("open",)
                if any(actual[k] != expected[k] for k in keys):
                    raise ValueError(f"CSV and status document disagree for {functionality}")
        keys = COUNTS if scope == "all" else ("open",) if scope == "open" else ()
        if any(detail_counts[k] != stats["totals"][k] for k in keys):
            raise ValueError("CSV coverage does not match status document totals")
    portfolio_known = scope == "all" or stats is not None
    totals = dict(stats["totals"]) if stats else detail_counts if scope == "all" else {
        "total": None, "open": detail_counts["open"] if scope == "open" else None,
        "resolved": None, "closed": None,
    }
    if scope != "all":
        warnings.append("Issue natures, ticket details and weekly intake cover only the supplied " +
                        ("currently open tickets." if scope == "open" else "filtered tickets."))
    if scope == "partial":
        absent = [name for name in ("open", "resolved", "closed") if not detail_counts[name]]
        if absent:
            warnings.append("No " + "/".join(absent) + " tickets were supplied. Zero in the detail analysis does not establish a complete portfolio count.")
    functionality_labels = set(grouped)
    functionality_labels.update(row.get("module_label", row.get("category", ""))
                                for key, row in supplied.items() if key not in labels)
    functions, natures = [], []
    for functionality in sorted(functionality_labels, key=str.casefold):
        rows = grouped.get(functionality, [])
        local = counts(rows)
        official = supplied.get(normalize_key(functionality))
        metrics = {k: official[k] for k in COUNTS} if official else local if scope == "all" else {
            "total": None, "open": local["open"] if scope == "open" else None,
            "resolved": None, "closed": None,
        }
        owner = owners.get(normalize_key(functionality), dict.fromkeys(("officer", "dd", "jd"), "Not mapped"))
        functions.append({"functionality": functionality, **metrics, **owner,
                          "detail_total": len(rows), "last_week": sum(r["last_week"] for r in rows),
                          "prior_week": sum(r["prior_week"] for r in rows)})
        nature_groups = {}
        for row in rows:
            nature_groups.setdefault(row["nature"], []).append(row)
        for nature, members in nature_groups.items():
            natures.append({"functionality": functionality, "nature": nature, **counts(members),
                            "last_week": sum(r["last_week"] for r in members),
                            "prior_week": sum(r["prior_week"] for r in members)})
    for functionality in functionality_labels:
        members = [r for r in natures if r["functionality"] == functionality]
        for period, ordering in (("rank", lambda r: (-r["open"], -r["total"], r["nature"])),
                                 ("week_rank", lambda r: (-r["last_week"], r["nature"]))):
            ranked = sorted((r for r in members if r["nature"] != OTHER and
                             (period == "rank" or r["last_week"] > 0)), key=ordering)
            for index, row in enumerate(ranked, 1):
                row[period] = index
    unmapped = sum(i["officer"] == "Not mapped" for i in issues)
    if unmapped:
        warnings.append(f"{unmapped} tickets have no mapped responsible officer.")
    return {
        "date": report_date.isoformat(), "scope": scope, "portfolio_known": portfolio_known,
        "count_source": stats["source"]["name"] if stats else "Issue CSV" if scope == "all" else "Open/filtered issue CSV",
        "totals": totals, "detail_counts": detail_counts,
        "week": {"start": week_start.isoformat(), "end": week_end.isoformat(),
                 "prior_start": prior_start.isoformat(), "prior_end": (week_start - timedelta(days=1)).isoformat()},
        "functions": functions, "natures": natures, "issues": issues, "warnings": warnings,
        "review_count": sum(i["nature"] == OTHER for i in issues),
        "rules_version": classifier.version, "rules_fingerprint": classifier.fingerprint,
        "projects": sorted(set(df["Project"].astype(str))) if "Project" in df else [],
    }
