import argparse
import sys
from datetime import date
from pathlib import Path
import pandas as pd

from .core.classifier import IssueClassifier
from .core.chat_parser import ChatParser, ChatKnowledgeExtractor
from .core.entity_matcher import EntityMatcher
from .core.workforce import WorkforceMapper
from .core.ingest import IngestValidator
from .reporters.excel_reporter import ExcelReporter
from .reporters.pptx_reporter import PPTXReporter
from .reporters.docx_reporter import DocxReporter
from .reporters.html_reporter import HTMLReporter

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
    cmd_report = subparsers.add_parser("report", help="Generate the full reporting pack (Excel, PPTX, HTML, Word).")
    cmd_report.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_report.add_argument("--teams", type=str, help="Optional path to teams.csv for workforce hierarchy mapping.")
    cmd_report.add_argument("--chats", nargs="*", type=str, help="Optional list of WhatsApp chat exports (.txt or .zip).")
    cmd_report.add_argument("--date", "-d", type=str, default=str(date.today()), help="Report snapshot date.")
    cmd_report.add_argument("--out-dir", "-o", type=str, default="reports", help="Output directory for generated reports.")

    # Command: workforce
    cmd_workforce = subparsers.add_parser("workforce", help="Map issues across organizational management hierarchy.")
    cmd_workforce.add_argument("issues_csv", type=str, help="Path to issue tracker CSV.")
    cmd_workforce.add_argument("teams_csv", type=str, help="Path to teams.csv mapping hierarchy.")
    cmd_workforce.add_argument("--output", "-o", type=str, help="Optional path to export workload summary CSV.")

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

    # 1. Excel
    xl_path = out_dir / f"CITES_Report_{args.date}.xlsx"
    ExcelReporter.generate_report(df_enriched, xl_path, title=f"CITES Operations Report - {args.date}")
    print(f"[OK] Excel Report: {xl_path}")

    # 2. PowerPoint
    pptx_path = out_dir / f"CITES_Brief_{args.date}.pptx"
    PPTXReporter.generate_presentation(df_enriched, pptx_path, report_date=args.date)
    print(f"[OK] PowerPoint Presentation: {pptx_path}")

    # 3. HTML Dashboard
    html_path = out_dir / f"Dashboard_{args.date}.html"
    HTMLReporter.generate_html(df_enriched, html_path, report_date=args.date)
    print(f"[OK] HTML Dashboard: {html_path}")

    # 4. Government Note
    top_cats = df_enriched["major_topic_label"].value_counts().head(6).index.tolist()
    major_items = []
    for cat in top_cats:
        desc_matches = df_enriched[df_enriched["major_topic_label"] == cat]["category_description"]
        desc = desc_matches.iloc[0] if not desc_matches.empty else ""
        major_items.append({"title": cat, "description": desc})

    docx_path = out_dir / f"Administrative_Note_{args.date}.docx"
    DocxReporter.generate_govt_note(major_items, docx_path, note_date=args.date)
    print(f"[OK] Administrative Note: {docx_path}")

    # 5. Knowledge Note (if chats provided)
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
    print(f"Mapping {args.issues_csv} against {args.teams_csv}...")
    _, _, df_issues = IngestValidator.validate_issue_csv(args.issues_csv)
    df_teams = pd.read_csv(args.teams_csv, encoding="utf-8-sig", dtype=str)

    mapper = WorkforceMapper()
    res = mapper.process_workload(df_issues, df_teams)
    print(f"Total Issues: {res['kpis']['total_issues']}")
    print(f"Open Issues: {res['kpis']['open_issues']}")
    print(f"Resolved Issues: {res['kpis']['resolved_issues']}")
    print("Routing Breakdown:", res["kpis"]["routing_breakdown"])

    if res["unmapped_categories"]:
        print("\nWarning: The following categories are not mapped in teams.csv:")
        for cat in res["unmapped_categories"]:
            print(f" - {cat}")

if __name__ == "__main__":
    main()
