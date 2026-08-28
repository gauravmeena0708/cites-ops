"""End-to-end daily MantisBT reporting pipeline."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Sequence, Set

import pandas as pd

from .core.classifier import IssueClassifier
from .core.ingest import IngestValidator
from .core.mantis_client import MantisClient
from .core.workforce import WorkforceMapper
from .reporters.defect_drilldown_reporter import DefectDrilldownReporter
from .reporters.excel_reporter import ExcelReporter
from .reporters.html_reporter import HTMLReporter
from .reporters.interactive_topics_reporter import InteractiveTopicsReporter
from .reporters.status_docx_reporter import StatusDocxReporter
from .reporters.weekly_resolutions_reporter import WeeklyResolutionsReporter


DEFAULT_REPORTS = frozenset(
    {
        "status",
        "issues_csv",
        "issues_xlsx",
        "dashboard",
        "issue_topics",
        "defect_drilldown",
        "weekly_resolutions",
    }
)

REPORT_FILENAMES = {
    "status": "cites_status_{date}.docx",
    "issues_csv": "cites_issues_{date}.csv",
    "issues_xlsx": "cites_issues_{date}.xlsx",
    "dashboard": "cites_dashboard_{date}.html",
    "issue_topics": "cites_issue_topics_{date}.html",
    "defect_drilldown": "cites_defect_drilldown_{date}.html",
    "weekly_resolutions": "cites_weekly_resolutions_{date}.html",
}


@dataclass(frozen=True)
class DailyRunResult:
    output_dir: Path
    manifest_path: Path
    artifacts: Dict[str, Path]


def parse_report_date(value: Optional[str]) -> date:
    """Parse an ISO report date and fail instead of silently changing the date."""
    if value is None:
        return date.today()
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Invalid report date {value!r}; expected YYYY-MM-DD") from exc


def parse_report_selection(value: Optional[str]) -> Set[str]:
    """Parse `default`, `all`, or a comma-separated list of report keys."""
    if not value or value.strip().lower() in {"default", "all"}:
        return set(DEFAULT_REPORTS)
    selected = {part.strip().lower() for part in value.split(",") if part.strip()}
    unknown = selected - set(REPORT_FILENAMES)
    if unknown:
        valid = ", ".join(sorted(REPORT_FILENAMES))
        raise ValueError(f"Unknown report(s): {', '.join(sorted(unknown))}. Valid values: {valid}")
    if not selected:
        raise ValueError("At least one report must be selected")
    return selected


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _resolve_token(input_dir: Path, explicit_token: Optional[str], token_file: Optional[Path]) -> str:
    token = (explicit_token or os.getenv("MANTIS_API_TOKEN") or "").strip()
    source = token_file or (input_dir / "token.txt")
    if not token and source.is_file():
        token = source.read_text(encoding="utf-8-sig").strip()
    if not token:
        raise ValueError(
            "MantisBT API token is missing. Set MANTIS_API_TOKEN or provide --token-file."
        )
    return token


def _resolve_teams_file(input_dir: Path, explicit_path: Optional[Path]) -> Path:
    candidates = [
        explicit_path,
        input_dir / "issue_teams.csv",
        input_dir / "Issue_teams.csv",
        input_dir / "teams.csv",
    ]
    for candidate in candidates:
        if candidate and candidate.is_file():
            return candidate.resolve()
    raise FileNotFoundError(
        f"No issue-team mapping found in {input_dir}. Expected issue_teams.csv."
    )


def _load_prior_manifests(output_root: Path, report_date: date) -> Sequence[Dict[str, Any]]:
    manifests = []
    if not output_root.is_dir():
        return manifests
    for path in output_root.glob("????-??-??/manifest.json"):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            payload_date = parse_report_date(payload.get("report_date"))
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        if payload_date < report_date and payload.get("metrics"):
            manifests.append(payload)
    manifests.sort(key=lambda item: item["report_date"])
    return manifests


def _snapshot_source(report_date: date, totals: Dict[str, int], categories: pd.DataFrame) -> Dict[str, Any]:
    category_rows = []
    for row in categories.to_dict(orient="records"):
        category = str(row.get("Category", "Unassigned"))
        category_rows.append(
            {
                "module_key": WorkforceMapper._normalize_key(category),
                "category": category,
                "open": int(row.get("open", 0)),
                "resolved": int(row.get("resolved", 0)),
                "closed": int(row.get("closed", 0)),
                "total": int(row.get("total", 0)),
            }
        )
    return {
        "source": {"data_date": report_date.isoformat()},
        "totals": {key: int(value) for key, value in totals.items()},
        "categories": category_rows,
    }


def _manifest_to_snapshot(manifest: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    metrics = manifest.get("metrics", {})
    if not metrics.get("totals") or not isinstance(metrics.get("categories"), list):
        return None
    return {
        "source": {"data_date": manifest["report_date"]},
        "totals": metrics["totals"],
        "categories": metrics["categories"],
    }


def _deduplicate_snapshot(df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """Collapse pagination overlap by ID, retaining the most recently updated row."""
    if "Id" not in df.columns or df.empty:
        return df, 0, 0
    working = df.copy()
    working["_id_key"] = working["Id"].astype(str).str.strip()
    duplicate_mask = working["_id_key"].ne("") & working["_id_key"].duplicated(keep=False)
    duplicate_id_count = int(working.loc[duplicate_mask, "_id_key"].nunique())
    if duplicate_id_count == 0:
        return df, 0, 0

    working["_updated_sort"] = pd.to_datetime(working.get("Updated"), errors="coerce")
    working["_source_order"] = range(len(working))
    working.sort_values(
        ["_id_key", "_updated_sort", "_source_order"],
        kind="stable",
        na_position="first",
        inplace=True,
    )
    before = len(working)
    working = working.drop_duplicates(subset=["_id_key"], keep="last")
    working.sort_values("_source_order", kind="stable", inplace=True)
    working.drop(columns=["_id_key", "_updated_sort", "_source_order"], inplace=True)
    working.reset_index(drop=True, inplace=True)
    return working, duplicate_id_count, before - len(working)


def _validate_snapshot(df: pd.DataFrame) -> None:
    required = {"Id", "Category", "Status", "Assigned To", "Date Submitted"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"MantisBT response is missing required column(s): {', '.join(missing)}")
    if df.empty:
        raise ValueError("MantisBT returned no issues; no reports were published")
    normalized_ids = df["Id"].astype(str).str.strip()
    duplicates = normalized_ids[normalized_ids.ne("") & normalized_ids.duplicated(keep=False)]
    if not duplicates.empty:
        raise ValueError("Duplicate issue IDs remain after pagination-overlap normalization")


def run_daily_pipeline(
    *,
    report_date: date,
    input_dir: Path,
    output_root: Path,
    base_url: str,
    token: Optional[str] = None,
    token_file: Optional[Path] = None,
    teams_file: Optional[Path] = None,
    project: Optional[str] = None,
    reports: Optional[Iterable[str]] = None,
    overwrite: bool = False,
    verify_ssl: bool = True,
) -> DailyRunResult:
    """Fetch once, derive once, verify once, and publish a dated report directory."""
    input_dir = input_dir.resolve()
    output_root = output_root.resolve()
    selected = set(reports or DEFAULT_REPORTS)
    unknown = selected - set(REPORT_FILENAMES)
    if unknown:
        raise ValueError(f"Unknown report(s): {', '.join(sorted(unknown))}")

    resolved_token = _resolve_token(input_dir, token, token_file)
    resolved_teams = _resolve_teams_file(input_dir, teams_file)
    df_teams = pd.read_csv(resolved_teams, encoding="utf-8-sig", dtype=str)
    if df_teams.empty:
        raise ValueError(f"Issue-team mapping is empty: {resolved_teams}")

    final_dir = output_root / report_date.isoformat()
    if final_dir.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {final_dir}. Use --overwrite to replace it.")

    output_root.mkdir(parents=True, exist_ok=True)
    staging_dir = output_root / f".{report_date.isoformat()}.staging-{uuid.uuid4().hex[:8]}"
    staging_dir.mkdir(parents=False, exist_ok=False)
    log_lines = []

    def log(message: str) -> None:
        log_lines.append(message)
        print(message)

    try:
        client = MantisClient(base_url=base_url, token=resolved_token, verify_ssl=verify_ssl)
        user_info = client.test_connection()
        log(f"Authenticated to MantisBT as {user_info.get('name', 'unknown user')}")
        raw_issues = client.fetch_issues(project_id=project)
        df_raw = MantisClient.to_dataframe(raw_issues)
        df_raw, duplicate_id_count, duplicate_rows_removed = _deduplicate_snapshot(df_raw)
        _validate_snapshot(df_raw)
        log(f"Fetched {len(df_raw):,} unique issues")
        if duplicate_rows_removed:
            log(
                "Normalized MantisBT pagination overlap: "
                f"removed {duplicate_rows_removed:,} repeated rows across "
                f"{duplicate_id_count:,} issue IDs"
            )

        classifier = IssueClassifier()
        df_classified = classifier.classify_dataframe(df_raw)
        df_enriched = IngestValidator.compute_aging(df_classified, as_of_date=report_date)

        mapper = WorkforceMapper()
        workload_data = mapper.process_workload(df_enriched, df_teams)
        coverage = workload_data["kpis"]["coverage_pct"]
        log(f"Applied issue-team mapping; ownership coverage is {coverage}")

        totals, category_df = StatusDocxReporter.compute_metrics(df_enriched)
        current_snapshot = _snapshot_source(report_date, totals, category_df)
        prior_manifests = _load_prior_manifests(output_root, report_date)
        prior_snapshots = [snapshot for item in prior_manifests if (snapshot := _manifest_to_snapshot(item))]
        latest_prior = prior_snapshots[-1] if prior_snapshots else None
        artifacts: Dict[str, Path] = {}

        def target(key: str) -> Path:
            path = staging_dir / REPORT_FILENAMES[key].format(date=report_date.isoformat())
            artifacts[key] = path
            return path

        if "issues_csv" in selected:
            df_enriched.to_csv(target("issues_csv"), index=False, encoding="utf-8-sig")
        if "issues_xlsx" in selected:
            ExcelReporter.generate_report(
                df_enriched,
                target("issues_xlsx"),
                workload_data=workload_data,
                title=f"CITES Issues Report - {report_date.isoformat()}",
            )
        if "status" in selected:
            StatusDocxReporter.generate_status_docx(
                df_enriched,
                target("status"),
                report_date=report_date,
                prev_metrics=latest_prior.get("totals") if latest_prior else None,
            )
        if "dashboard" in selected:
            HTMLReporter.generate_html(
                df_enriched,
                target("dashboard"),
                workload_data=workload_data,
                report_date=report_date.isoformat(),
            )
        if "issue_topics" in selected:
            InteractiveTopicsReporter.generate_html(
                df_enriched,
                target("issue_topics"),
                df_teams=df_teams,
                report_date=report_date.isoformat(),
            )
        if "defect_drilldown" in selected:
            DefectDrilldownReporter.generate_html(
                df_enriched,
                target("defect_drilldown"),
                report_date=report_date.isoformat(),
            )
        if "weekly_resolutions" in selected:
            WeeklyResolutionsReporter.generate_html(
                df_enriched,
                target("weekly_resolutions"),
                df_teams=df_teams,
                stats_sources=[*prior_snapshots, current_snapshot],
                report_date=report_date.isoformat(),
            )

        missing = [key for key, path in artifacts.items() if not path.is_file() or path.stat().st_size == 0]
        if missing:
            raise RuntimeError(f"Report generation produced missing/empty artifact(s): {', '.join(missing)}")

        manifest = {
            "schema_version": 1,
            "report_date": report_date.isoformat(),
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "source": {
                "mantis_url": client.base_url,
                "project": project,
                "issue_count": len(df_enriched),
            },
            "inputs": {
                "issue_teams_file": resolved_teams.name,
                "issue_teams_sha256": _sha256(resolved_teams),
            },
            "metrics": current_snapshot | {"ownership_coverage": coverage},
            "artifacts": {
                key: {
                    "file": path.name,
                    "bytes": path.stat().st_size,
                    "sha256": _sha256(path),
                }
                for key, path in sorted(artifacts.items())
            },
            "quality": {
                "duplicate_issue_ids_returned": duplicate_id_count,
                "duplicate_rows_removed": duplicate_rows_removed,
                "duplicate_issue_ids_after_normalization": 0,
                "artifact_count": len(artifacts),
                "reports_requested": sorted(selected),
                "weekly_values_are_snapshot_deltas": True,
            },
        }
        manifest_path = staging_dir / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
        (staging_dir / "run.log").write_text("\n".join(log_lines) + "\n", encoding="utf-8")

        if final_dir.exists():
            shutil.rmtree(final_dir)
        staging_dir.replace(final_dir)
        published = {key: final_dir / path.name for key, path in artifacts.items()}
        return DailyRunResult(final_dir, final_dir / "manifest.json", published)
    except Exception:
        shutil.rmtree(staging_dir, ignore_errors=True)
        raise
