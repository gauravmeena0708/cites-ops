import argparse
import sys
from datetime import date
from pathlib import Path
from typing import Optional, List
import pandas as pd

from .core.classifier import IssueClassifier
from .core.chat_parser import ChatParser, ChatKnowledgeExtractor
from .core.entity_matcher import EntityMatcher
from .core.workforce import WorkforceMapper
from .core.ingest import IngestValidator
from .core.stats_parser import StatsDocxParser
from .reporters.excel_reporter import ExcelReporter
from .reporters.pptx_reporter import PPTXReporter
from .reporters.regional_pptx_reporter import RegionalPPTXReporter
from .reporters.docx_reporter import DocxReporter
from .reporters.html_reporter import HTMLReporter
from .reporters.defect_drilldown_reporter import DefectDrilldownReporter
from .reporters.interactive_topics_reporter import InteractiveTopicsReporter
from .reporters.weekly_resolutions_reporter import WeeklyResolutionsReporter

def find_default_teams_file() -> Optional[Path]:
    """Look for standard teams mapping files in current working directory and common locations."""
    candidates = [
        Path("teams.csv"),
        Path("Issue_teams.csv"),
        Path("issues/Issue_teams.csv"),
        Path("../issues/Issue_teams.csv"),
        Path(r"C:\Users\IT\Downloads\CITES\teams.csv"),
        Path(r"C:\Users\IT\Downloads\CITES\issues\Issue_teams.csv"),
    ]
    for cand in candidates:
        if cand.is_file():
            return cand.resolve()
    return None

def find_default_stats_files() -> List[Path]:
    """Look for standard stats DOCX files in issues/ or current working directory."""
    candidates = []
    bases = [Path("."), Path("issues"), Path("archive"), Path(r"C:\Users\IT\Downloads\CITES"), Path(r"C:\Users\IT\Downloads\CITES\issues"), Path(r"C:\Users\IT\Downloads\CITES\archive")]
    for base in bases:
        if base.is_dir():
            for f in base.rglob("*.docx"):
                if not f.name.startswith("~$") and "stat" in f.name.lower():
                    cand_resolved = f.resolve()
                    if cand_resolved not in candidates:
                        candidates.append(cand_resolved)
    return candidates

def main():
    parser = argparse.ArgumentParser(
        prog="cites-ops",
        description="CITES Operations Intelligence & Issue Triage Framework",
    )
    subparsers = parser.add_subparsers(dest="command", help="Available subcommands")

    # Command: classify
    cmd_classify = subparsers.add_parser("classify", help="Classify an issue tracker CSV into Major/Minor categories.")
    cmd_classify.add_argument("input_csv", type=str, help="Path to tracker CSV file.")
    cmd_classify.add_argument("--output", "-o", type=str, default="categorized_issues.csv", help="Path to save categorized CSV.")
    cmd_classify.add_argument("--excel", "-x", type=str, help="Path to save formatted Excel workbook (.xlsx).")
    cmd_classify.add_argument("--rules", type=str, help="Optional custom rules.yaml path.")

    # Command: chat-kb
    cmd_chat = subparsers.add_parser("chat-kb", help="Extract technical knowledge & mask PII from WhatsApp chat exports.")
    cmd_chat.add_argument("chat_files", nargs="+", type=str, help="One or more WhatsApp chat export paths (.txt or .zip).")
    cmd_chat.add_argument("--issues", type=str, help="Optional tracker CSV to cross-reference ticket IDs.")
    cmd_chat.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report date (YYYY-MM-DD).")
    cmd_chat.add_argument("--output", "-o", type=str, default="Knowledge_Note.docx", help="Path to save Word Knowledge Note (.docx).")

    # Command: report
    cmd_report = subparsers.add_parser("report", help="Generate the full reporting pack (Excel, PPTX decks, HTML Dashboards, Defect Drilldown, Word).")
    cmd_report.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_report.add_argument("--teams", type=str, help="Optional path to teams.csv for workforce hierarchy mapping.")
    cmd_report.add_argument("--stats", nargs="*", type=str, help="Optional Samadhan Setu / CITES daily statistics DOCX files.")
    cmd_report.add_argument("--chats", nargs="*", type=str, help="Optional list of WhatsApp chat exports (.txt or .zip).")
    cmd_report.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")
    cmd_report.add_argument("--out-dir", "-o", type=str, default="reports", help="Output directory for generated reports.")

    # Command: workforce
    cmd_workforce = subparsers.add_parser("workforce", help="Map issues across 5-tier organizational & topical hierarchy (JD -> DD -> Handler -> Category -> Defect Topic).")
    cmd_workforce.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_workforce.add_argument("teams_csv", nargs="?", type=str, help="Optional path to teams.csv mapping hierarchy (auto-detected if omitted).")
    cmd_workforce.add_argument("--html", type=str, help="Path to export interactive HTML workforce dashboard.")
    cmd_workforce.add_argument("--excel", "-x", type=str, help="Path to export Excel workload workbook.")
    cmd_workforce.add_argument("--output", "-o", type=str, help="Optional path to export workload summary CSV.")

    # Command: defects
    cmd_defects = subparsers.add_parser("defects", help="Generate interactive root-cause defect drilldown HTML with issue pop-up modals.")
    cmd_defects.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_defects.add_argument("--output", "-o", type=str, default="Defect_Drilldown.html", help="Path to save HTML dashboard.")
    cmd_defects.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")
    cmd_defects.add_argument("--rules", type=str, help="Optional custom rules.yaml path.")

    # Command: topics
    cmd_topics = subparsers.add_parser("topics", help="Generate interactive issue topics & ownership hierarchy HTML dashboard (CITES_Interactive_Issue_Topics).")
    cmd_topics.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_topics.add_argument("--teams", type=str, help="Optional path to teams.csv.")
    cmd_topics.add_argument("--output", "-o", type=str, default="CITES_Interactive_Issue_Topics.html", help="Path to save HTML.")
    cmd_topics.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")

    # Command: weekly
    cmd_weekly = subparsers.add_parser("weekly", help="Generate weekly resolutions & transition calendar grid HTML dashboard (CITES_Resolved_Closed_Weekly).")
    cmd_weekly.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_weekly.add_argument("--teams", type=str, help="Optional path to teams.csv.")
    cmd_weekly.add_argument("--stats", nargs="*", type=str, help="Optional stats.docx files.")
    cmd_weekly.add_argument("--output", "-o", type=str, default="CITES_Resolved_Closed_Weekly.html", help="Path to save HTML.")
    cmd_weekly.add_argument("--week", type=str, help="Optional start date of week (YYYY-MM-DD).")
    cmd_weekly.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")

    # Command: regional-pptx
    cmd_reg = subparsers.add_parser("regional-pptx", help="Generate executive light-themed PPTX focused on major issue types affecting major regional offices (last 7 days).")
    cmd_reg.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_reg.add_argument("--output", "-o", type=str, default="Regional_Offices_Review.pptx", help="Path to save PowerPoint presentation (.pptx).")
    cmd_reg.add_argument("--days", type=int, default=7, help="Days window for recent intake focus (default: 7).")
    cmd_reg.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")
    cmd_reg.add_argument("--rules", type=str, help="Optional custom rules.yaml path.")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command == "classify":
        run_classify(args)
    elif args.command == "chat-kb":
        run_chat_kb(args)
    elif args.command == "report":
        run_report(args)
    elif args.command == "workforce":
        run_workforce(args)
    elif args.command == "defects":
        run_defects(args)
    elif args.command == "topics":
        run_topics(args)
    elif args.command == "weekly":
        run_weekly(args)
    elif args.command == "regional-pptx":
        run_regional_pptx(args)

def run_classify(args):
    print(f"Reading {args.input_csv}...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.input_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    classifier = IssueClassifier(args.rules)
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    out_csv = Path(args.output)
    df_enriched.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[OK] Categorized CSV written to {out_csv}")

    if args.excel:
        out_xl = ExcelReporter.generate_report(df_enriched, args.excel)
        print(f"[OK] Formatted Excel workbook written to {out_xl}")

def run_chat_kb(args):
    all_dfs = []
    for cf in args.chat_files:
        print(f"Parsing chat file: {cf}...")
        try:
            df_part = ChatParser.parse_file(cf)
            print(f"  Extracted {len(df_part)} messages.")
            all_dfs.append(df_part)
        except Exception as e:
            print(f"  Warning: Failed to parse {cf}: {e}")

    if not all_dfs:
        print("No chat messages were parsed.")
        sys.exit(1)

    df_chat = pd.concat(all_dfs, ignore_index=True)
    print(f"Total chat messages across sources: {len(df_chat):,}")

    issue_index = None
    if args.issues:
        _, _, df_issues = IngestValidator.validate_issue_csv(args.issues)
        matcher = EntityMatcher()
        issue_index = matcher.build_issue_index(df_issues)
        print(f"Built cross-reference entity index from {len(df_issues):,} tracker issues.")

    extractor = ChatKnowledgeExtractor()
    items = extractor.extract_knowledge_items(df_chat, issue_index)
    print(f"Extracted {len(items):,} technical resolution and problem guidance items.")

    out_docx = DocxReporter.generate_knowledge_note(items, args.output, note_date=args.date)
    print(f"[OK] Field Office Knowledge Note written to {out_docx}")

def run_report(args):
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Processing daily intake for {args.date}...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.issues_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    classifier = IssueClassifier()
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    # Resolve teams mapping
    teams_path = Path(args.teams) if args.teams else find_default_teams_file()
    df_teams = None
    workload_data = None
    if teams_path and teams_path.is_file():
        print(f"Loading workforce hierarchy mapping from: {teams_path}...")
        df_teams = pd.read_csv(teams_path, encoding="utf-8-sig", dtype=str)
        mapper = WorkforceMapper()
        workload_data = mapper.process_workload(df_enriched, df_teams)
        print(f"  Mapped {workload_data['kpis']['total_issues']:,} issues across {len(workload_data['hierarchy_levels'])} hierarchy tiers.")

    # Parse any stats.docx files
    stats_sources = []
    stats_paths = [Path(p) for p in args.stats] if args.stats else find_default_stats_files()
    for sp in stats_paths:
        if sp.is_file():
            parsed = StatsDocxParser.parse_file(sp)
            if parsed:
                stats_sources.append(parsed)

    # 1. Excel
    xl_path = out_dir / f"CITES_Report_{args.date}.xlsx"
    ExcelReporter.generate_report(df_enriched, xl_path, workload_data=workload_data, title=f"CITES Operations Report - {args.date}")
    print(f"[OK] Excel Report: {xl_path}")

    # 2. PowerPoint (Executive Overview Deck)
    pptx_path = out_dir / f"CITES_Brief_{args.date}.pptx"
    PPTXReporter.generate_presentation(df_enriched, pptx_path, workload_data=workload_data, report_date=args.date)
    print(f"[OK] PowerPoint Presentation: {pptx_path}")

    # 3. PowerPoint (Regional Offices & 7-Day Defect Focus)
    reg_pptx_path = out_dir / f"Regional_Offices_Review_{args.date}.pptx"
    RegionalPPTXReporter.generate_presentation(df_enriched, reg_pptx_path, report_date=args.date, days_window=7)
    print(f"[OK] Regional Offices PPTX (7-Day Focus): {reg_pptx_path}")

    # 4. HTML Dashboard (Interactive Hierarchy + Cross-Module Matrix)
    html_path = out_dir / f"Dashboard_{args.date}.html"
    HTMLReporter.generate_html(df_enriched, html_path, workload_data=workload_data, report_date=args.date)
    print(f"[OK] HTML Dashboard: {html_path}")

    # 5. Defect Drilldown Pop-up Dashboard
    defect_html_path = out_dir / f"Defect_Drilldown_{args.date}.html"
    DefectDrilldownReporter.generate_html(df_enriched, defect_html_path, report_date=args.date)
    print(f"[OK] Defect Drilldown Pop-up Dashboard: {defect_html_path}")

    # 6. Interactive Issue Topics HTML (CITES_Interactive_Issue_Topics)
    topics_html_path = out_dir / f"CITES_Interactive_Issue_Topics_{args.date}.html"
    InteractiveTopicsReporter.generate_html(df_enriched, topics_html_path, df_teams=df_teams, report_date=args.date)
    print(f"[OK] Interactive Issue Topics Dashboard: {topics_html_path}")

    # 7. Weekly Resolutions HTML (CITES_Resolved_Closed_Weekly)
    if df_teams is not None:
        weekly_html_path = out_dir / f"CITES_Resolved_Closed_Weekly_{args.date}.html"
        WeeklyResolutionsReporter.generate_html(df_enriched, weekly_html_path, df_teams=df_teams, stats_sources=stats_sources, report_date=args.date)
        print(f"[OK] Weekly Resolutions HTML: {weekly_html_path}")

    # 8. Government Note
    top_cats = df_enriched["major_topic_label"].value_counts().head(6).index.tolist()
    major_items = []
    for cat in top_cats:
        desc_matches = df_enriched[df_enriched["major_topic_label"] == cat]["category_description"]
        desc = desc_matches.iloc[0] if not desc_matches.empty else ""
        major_items.append({"title": cat, "description": desc})

    docx_path = out_dir / f"Administrative_Note_{args.date}.docx"
    DocxReporter.generate_govt_note(major_items, docx_path, note_date=args.date)
    print(f"[OK] Administrative Note: {docx_path}")

    # 9. Knowledge Note (if chats provided)
    if args.chats:
        all_dfs = []
        for cf in args.chats:
            try:
                df_part = ChatParser.parse_file(cf)
                all_dfs.append(df_part)
            except Exception as e:
                print(f"  Warning: Could not parse chat {cf}: {e}")
        if all_dfs:
            df_chat = pd.concat(all_dfs, ignore_index=True)
            matcher = EntityMatcher()
            issue_index = matcher.build_issue_index(df_enriched)
            extractor = ChatKnowledgeExtractor(matcher)
            items = extractor.extract_knowledge_items(df_chat, issue_index)
            kb_path = out_dir / f"Field_Office_Knowledge_Note_{args.date}.docx"
            DocxReporter.generate_knowledge_note(items, kb_path, note_date=args.date)
            print(f"[OK] Knowledge Note: {kb_path}")

    print(f"\nAll substantive reports successfully generated in: {out_dir.resolve()}")

def run_workforce(args):
    teams_path = Path(args.teams_csv) if args.teams_csv else find_default_teams_file()
    if not teams_path or not teams_path.is_file():
        print("Error: No teams.csv / Issue_teams.csv found or provided.")
        sys.exit(1)

    print(f"Mapping {args.issues_csv} against {teams_path}...")
    _, _, df_issues = IngestValidator.validate_issue_csv(args.issues_csv)
    df_teams = pd.read_csv(teams_path, encoding="utf-8-sig", dtype=str)

    classifier = IssueClassifier()
    df_classified = classifier.classify_dataframe(df_issues)
    df_enriched = IngestValidator.compute_aging(df_classified)

    mapper = WorkforceMapper()
    res = mapper.process_workload(df_enriched, df_teams)
    print(f"\n--- Workforce Accountability & Workload Summary ---")
    print(f"Total Issues Ingested : {res['kpis']['total_issues']:,}")
    print(f"Open Backlog          : {res['kpis']['open_issues']:,}")
    print(f"Resolved / Closed     : {res['kpis']['resolved_issues']:,}")
    print(f"Resolution Rate       : {res['kpis']['resolution_rate']}")
    print(f"Ownership Coverage    : {res['kpis']['coverage_pct']}")
    print(f"Routing Breakdown     : {res['kpis']['routing_breakdown']}")

    print("\n--- Top 10 Major Problem Categories ---")
    for cat in res["top_10_categories"]:
        print(f" #{cat['rank']:<2} {cat['category']:<32} | Total: {cat['total']:<5} | Open: {cat['open']:<5} | Share: {cat['share_of_backlog']:<6} | Handler: {cat['handler']}")

    if res["unmapped_categories"]:
        print(f"\nWarning: {len(res['unmapped_categories'])} categories not mapped in teams.csv:")
        for cat in res["unmapped_categories"]:
            print(f" - {cat}")
    else:
        print("\n[OK] 100% of categories successfully mapped (0 unmapped).")

    if args.output:
        df_summary = pd.DataFrame(res["category_summary"])
        df_summary.to_csv(args.output, index=False, encoding="utf-8-sig")
        print(f"[OK] Category summary exported to: {args.output}")

    if args.html:
        HTMLReporter.generate_html(df_enriched, args.html, workload_data=res, title="CITES Workforce & Operations Dashboard")
        print(f"[OK] Interactive HTML Workforce Dashboard written to: {args.html}")

    if args.excel:
        ExcelReporter.generate_report(df_enriched, args.excel, workload_data=res, title="CITES Workforce & Operations Report")
        print(f"[OK] Excel Workload Workbook written to: {args.excel}")

def run_defects(args):
    print(f"Classifying {args.issues_csv} for Root-Cause Defect Drilldown...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.issues_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    classifier = IssueClassifier(args.rules)
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    out_file = DefectDrilldownReporter.generate_html(
        df_enriched,
        args.output,
        report_date=args.date,
        title="CITES System-Wide Root-Cause Defect Drivers & Issue Drilldown"
    )
    print(f"[OK] Interactive Defect Drilldown Pop-up Dashboard written to: {out_file}")

def run_topics(args):
    print(f"Generating Interactive Issue Topics Dashboard from {args.issues_csv}...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.issues_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    teams_path = Path(args.teams) if args.teams else find_default_teams_file()
    df_teams = None
    if teams_path and teams_path.is_file():
        df_teams = pd.read_csv(teams_path, encoding="utf-8-sig", dtype=str)

    classifier = IssueClassifier()
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    out_file = InteractiveTopicsReporter.generate_html(
        df_enriched,
        args.output,
        df_teams=df_teams,
        report_date=args.date,
    )
    print(f"[OK] CITES Interactive Issue Topics HTML written to: {out_file}")

def run_weekly(args):
    print(f"Generating Weekly Resolutions HTML from {args.issues_csv}...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.issues_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    teams_path = Path(args.teams) if args.teams else find_default_teams_file()
    if not teams_path or not teams_path.is_file():
        print("Error: No teams.csv found. Weekly report requires workforce mapping.")
        sys.exit(1)

    df_teams = pd.read_csv(teams_path, encoding="utf-8-sig", dtype=str)

    stats_sources = []
    stats_paths = [Path(p) for p in args.stats] if args.stats else find_default_stats_files()
    for sp in stats_paths:
        if sp.is_file():
            parsed = StatsDocxParser.parse_file(sp)
            if parsed:
                stats_sources.append(parsed)

    classifier = IssueClassifier()
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    out_file = WeeklyResolutionsReporter.generate_html(
        df_enriched,
        args.output,
        df_teams=df_teams,
        stats_sources=stats_sources,
        report_date=args.date,
        week_date=args.week,
    )
    print(f"[OK] CITES Resolved & Closed Weekly HTML written to: {out_file}")

def run_regional_pptx(args):
    print(f"Generating Regional Offices PPTX from {args.issues_csv} (Last {args.days} Days Focus)...")
    valid, errors, df = IngestValidator.validate_issue_csv(args.issues_csv)
    if not valid:
        print(f"Validation Error: {', '.join(errors)}")
        sys.exit(1)

    classifier = IssueClassifier(args.rules)
    df_classified = classifier.classify_dataframe(df)
    df_enriched = IngestValidator.compute_aging(df_classified)

    out_file = RegionalPPTXReporter.generate_presentation(
        df_enriched,
        args.output,
        report_date=args.date,
        days_window=args.days,
        title="Regional Defect Diagnostics & Major Offices Review"
    )
    print(f"[OK] Light-Themed Regional Offices Presentation (.pptx) written to: {out_file}")

if __name__ == "__main__":
    main()
