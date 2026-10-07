"""Offline folder-to-brief workflow. Existing legacy commands remain available."""

import hashlib
import json
import shutil
import uuid
from time import perf_counter
from datetime import date
from pathlib import Path

import pandas as pd

from .core.briefing import prepare_briefing
from .core.ingest import IngestValidator
from .core.issue_nature import IssueNatureClassifier
from .core.knowledge import extract_candidates
from .core.top_natures import build_rankings
from .core.stats_parser import StatsDocxParser
from .reporters.briefing_reporter import write_html, write_workbook


def _one(paths, description):
    paths = sorted(set(paths))
    if len(paths) != 1:
        raise ValueError(f"Expected one {description}; found {len(paths)}. Supply an explicit file option.")
    return paths[0]


def _load_json(path, default):
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return default


def _remove_child(root, path):
    # Every recursive cleanup is constrained to a direct child we created.
    root, path = root.resolve(), path.resolve()
    if path.parent != root or not path.name.startswith(".brief-"):
        raise ValueError("Refusing cleanup outside the briefing output directory")
    shutil.rmtree(path)


def run_brief(input_dir, output_root, report_date, scope="auto", issues_file=None,
              teams_file=None, stats_file=None, no_stats=False, rules_file=None,
              chats=None, chat_since=None, chat_date_order="dmy", overwrite=False,
              top_n=15, ranking_period="all", no_pptx=False, pptx_runtime=None):
    input_dir, output_root = Path(input_dir).resolve(), Path(output_root).resolve()
    if not input_dir.is_dir():
        raise ValueError(f"Input directory does not exist: {input_dir}")
    if scope not in ("auto", "all", "open", "partial"):
        raise ValueError("Scope must be auto, all, open or partial")
    if stats_file and no_stats:
        raise ValueError("--stats and --no-stats cannot be combined")
    teams_file = Path(teams_file).resolve() if teams_file else _one(
        [p for p in input_dir.glob("*.csv") if p.stem.casefold() in ("teams", "issue_teams")], "teams CSV")
    issues_file = Path(issues_file).resolve() if issues_file else _one(
        [p for p in input_dir.glob("*.csv") if p.resolve() != teams_file.resolve()
         and p.stem.casefold() not in ("teams", "issue_teams")], "issue CSV")
    valid, errors, df = IngestValidator.validate_issue_csv(str(issues_file))
    if not valid:
        raise ValueError("; ".join(errors))
    teams = pd.read_csv(teams_file, dtype=str, encoding="utf-8-sig").fillna("")
    stats = None
    if not no_stats:
        candidates = [Path(stats_file)] if stats_file else [p for p in input_dir.glob("*.docx") if "stat" in p.stem.casefold() and not p.name.startswith("~$")]
        if candidates:
            stats_file = _one(candidates, "status DOCX")
            stats = StatsDocxParser.parse_file(stats_file)
            if not stats:
                raise ValueError(f"Cannot read status tables from {stats_file}")
    final = output_root / report_date.isoformat()
    if final == input_dir or final in input_dir.parents or any(
        final == path.resolve() or final in path.resolve().parents
        for path in (issues_file, teams_file, *([Path(stats_file)] if stats else []))
    ):
        raise ValueError("Output date directory must not contain the input files")
    if final.exists() and not overwrite:
        raise FileExistsError(f"Output exists: {final}. Use --overwrite to replace this briefing.")
    previous = []
    for path in output_root.glob("????-??-??/_internal/manifest.json"):
        payload = _load_json(path, {})
        if payload.get("schema") == "brief-1" and payload.get("date", "") <= report_date.isoformat():
            previous.append((payload["date"], path.parent, payload))
    previous.sort(key=lambda x: x[0])
    cache = _load_json(previous[-1][1] / "nature_cache.json", {}) if previous else {}
    classifier = IssueNatureClassifier(rules_file, cache if isinstance(cache, dict) else {})
    timings = {}
    phase = perf_counter()
    data = prepare_briefing(df, teams, classifier, report_date, scope, stats)
    if no_stats:
        data["warnings"].append("Status Word document not used; counts and analysis use the declared CSV scope.")
    data["rankings"] = build_rankings(data, top_n, ranking_period)
    if not data["rankings"]["resolved_closed"]["available"]:
        data["warnings"].append("Resolved/closed nature rankings and examples are unavailable: this snapshot supplies only open-ticket details.")
    timings["prepare_seconds"] = round(perf_counter() - phase, 3)
    data["comparison"] = None
    for previous_date, _, payload in reversed(previous):
        if previous_date < data["date"] and payload.get("portfolio_known") and data["portfolio_known"] and payload.get("projects") == data["projects"]:
            data["comparison"] = {"date": previous_date,
                                  "delta": {k: data["totals"][k] - payload["totals"][k] for k in data["totals"]}}
            break
    knowledge = {}
    if previous:
        prior_entries = _load_json(previous[-1][1] / "knowledge.json", [])
        knowledge = {e["id"]: e for e in prior_entries if e["date"] <= data["date"]}
    chat_paths = []
    for value in chats or []:
        path = Path(value)
        if path.is_dir():
            chat_paths.extend(p for p in path.iterdir() if p.suffix.lower() in (".zip", ".txt"))
        elif path.is_file() and path.suffix.lower() in (".zip", ".txt"):
            chat_paths.append(path)
        else:
            raise ValueError(f"Chat source must be a .txt/.zip file or directory: {path}")
    chat_paths = sorted(set(p.resolve() for p in chat_paths))
    if any(final in path.parents for path in chat_paths):
        raise ValueError("Output date directory must not contain chat input files")
    if chats and not chat_paths:
        raise ValueError("No .zip or .txt chat exports found in the supplied paths")
    if chat_paths:
        phase = perf_counter()
        start = chat_since or date.fromisoformat(data["week"]["start"])
        if start > report_date:
            raise ValueError("--chat-since must be on or before the report date")
        # Reprocessing an export replaces its candidates in the selected window,
        # so refined extraction rules cannot leave stale rejected entries behind.
        sources = {p.name for p in chat_paths}
        knowledge = {key: entry for key, entry in knowledge.items()
                     if not (start.isoformat() <= entry["date"] <= data["date"] and
                             entry["evidence"][-1]["source"].split(" / ")[0] in sources)}
        entries, warnings = extract_candidates(chat_paths, data["issues"], start, report_date, chat_date_order)
        knowledge.update({e["id"]: e for e in entries})
        data["warnings"].extend(warnings)
        timings["knowledge_seconds"] = round(perf_counter() - phase, 3)
    knowledge_entries = sorted(knowledge.values(), key=lambda e: (e["date"], e["id"]), reverse=True)
    data["knowledge_file"] = "knowledge_base.html" if chat_paths or knowledge_entries else None
    output_root.mkdir(parents=True, exist_ok=True)
    staging = output_root / (".brief-staging-" + uuid.uuid4().hex)
    backup = output_root / (".brief-backup-" + uuid.uuid4().hex)
    staging.mkdir()
    try:
        internal = staging / "_internal"
        internal.mkdir()
        phase = perf_counter()
        write_workbook(data, staging / f"issues_{data['date']}.xlsx")
        if not no_pptx:
            from .reporters.top_natures_pptx import write_presentations
            write_presentations(data, staging, pptx_runtime)
        write_html(data, staging / f"dashboard_{data['date']}.html")
        if data["knowledge_file"]:
            write_html({"date": data["date"], "entries": knowledge_entries}, staging / "knowledge_base.html", "knowledge.html")
        timings["render_seconds"] = round(perf_counter() - phase, 3)
        manifest = {k: data[k] for k in ("date", "scope", "portfolio_known", "totals", "detail_counts", "week", "count_source", "projects", "rules_version", "rules_fingerprint", "review_count", "warnings")}
        manifest.update({"schema": "brief-1", "functions": data["functions"],
                         "ranking_period": ranking_period, "top_n": top_n,
                         "cache_hits": classifier.hits, "cache_misses": classifier.misses,
                         "knowledge_candidates": len(knowledge_entries), "inputs": {}, "artifacts": {}, "timings": timings})
        input_paths = [issues_file, teams_file, *chat_paths]
        if stats:
            input_paths.append(Path(stats_file))
        for path in input_paths:
            manifest["inputs"][path.name] = hashlib.sha256(path.read_bytes()).hexdigest()
        for path in staging.iterdir():
            if path.is_file():
                manifest["artifacts"][path.name] = {"bytes": path.stat().st_size,
                    "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
        for name, payload in (("manifest.json", manifest), ("nature_cache.json", classifier.cache), ("knowledge.json", knowledge_entries)):
            (internal / name).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        (internal / "run.log").write_text(
            f"Prepared {len(data['issues'])} tickets; scope={data['scope']}\n"
            f"Cache hits={classifier.hits}; classifications={classifier.misses}\n" + "\n".join(data["warnings"]), encoding="utf-8")
        if final.exists():
            # Never replace legacy packs or unrelated folders through this command.
            if _load_json(final / "_internal/manifest.json", {}).get("schema") != "brief-1":
                raise ValueError("Existing directory is not a simplified briefing; choose another --output-root")
            final.rename(backup)
        try:
            staging.rename(final)
        except OSError:
            if backup.exists():
                backup.rename(final)
            raise
        if backup.exists():
            _remove_child(output_root, backup)
    except Exception:
        if staging.exists():
            _remove_child(output_root, staging)
        raise
    return final, manifest
