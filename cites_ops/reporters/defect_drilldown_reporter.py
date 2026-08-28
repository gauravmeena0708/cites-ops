import html
import json
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd

class DefectDrilldownReporter:
    """
    Generates a dedicated standalone HTML dashboard focused on System-Wide Top 10 Root-Cause
    Defect Drivers with an interactive pop-up modal/drawer containing full issue data tables
    with Summary & Description as the 2nd column, sorted agewise (latest at top).
    """

    @classmethod
    def generate_html(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        report_date: Optional[Union[str, date]] = None,
        title: str = "CITES System-Wide Root-Cause Defect Drivers & Issue Drilldown",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        rep_date_str = str(report_date or date.today())
        total_cnt = len(df_classified)

        status_col = "Status" if "Status" in df_classified.columns else ""
        id_col = "Id" if "Id" in df_classified.columns else "Issue_Id"
        cat_col = "Category" if "Category" in df_classified.columns else ""
        assigned_col = "Assigned To" if "Assigned To" in df_classified.columns else ""
        summary_col = "Summary" if "Summary" in df_classified.columns else ""
        desc_col = "Description" if "Description" in df_classified.columns else ""
        rule_col = "rule_id" if "rule_id" in df_classified.columns else ""
        topic_col = "topic_label" if "topic_label" in df_classified.columns else ""
        major_col = "major_topic_label" if "major_topic_label" in df_classified.columns else ""
        desc_rule_col = "category_description" if "category_description" in df_classified.columns else ""
        date_sub_col = "Date Submitted" if "Date Submitted" in df_classified.columns else ""

        # Status counts
        status_series = df_classified[status_col].astype(str).str.lower() if status_col else pd.Series([])
        resolved_mask = status_series.isin(["resolved", "fixed", "closed"])
        resolved_cnt = int(resolved_mask.sum())
        open_cnt = total_cnt - resolved_cnt
        res_rate = f"{round(100 * resolved_cnt / (total_cnt or 1), 1)}%"

        # Group issues by rule_id / defect topic
        topic_groups = {}
        issues_by_rule: Dict[str, List[Dict[str, Any]]] = {}

        for _, row in df_classified.iterrows():
            r_id = str(row.get(rule_col, "C99_OTHER") or "C99_OTHER")
            r_topic = str(row.get(topic_col, "Other / Insufficient Detail") or "Other / Insufficient Detail")
            r_major = str(row.get(major_col, "Other or insufficient detail") or "Other or insufficient detail")
            r_desc = str(row.get(desc_rule_col, "") or "")

            if r_id not in topic_groups:
                topic_groups[r_id] = {
                    "rule_id": r_id,
                    "topic_label": r_topic,
                    "major_topic_label": r_major,
                    "description": r_desc,
                    "total": 0,
                    "open": 0,
                    "resolved": 0,
                }
                issues_by_rule[r_id] = []

            is_res = str(row.get(status_col, "")).lower() in ["resolved", "fixed", "closed"]
            topic_groups[r_id]["total"] += 1
            if is_res:
                topic_groups[r_id]["resolved"] += 1
            else:
                topic_groups[r_id]["open"] += 1

            age_val = int(row.get("age_days", 0)) if "age_days" in row and pd.notna(row.get("age_days")) else 0
            date_sub_val = str(row.get(date_sub_col, "")) if date_sub_col else ""

            # Store issue payload
            issues_by_rule[r_id].append({
                "id": str(row.get(id_col, "")),
                "summary": str(row.get(summary_col, "")),
                "description": str(row.get(desc_col, "")),
                "category": str(row.get(cat_col, "")),
                "assigned_to": str(row.get(assigned_col, "")),
                "status": str(row.get(status_col, "Open")),
                "is_resolved": is_res,
                "age_days": age_val,
                "date_submitted": date_sub_val,
            })

        # Sort issues inside each topic agewise (latest at top: age_days ascending)
        for r_id in issues_by_rule:
            issues_by_rule[r_id].sort(key=lambda x: (x.get("age_days", 0), x.get("id", "")))

        # Rank defect topics by open backlog descending, then total descending
        sorted_topics = list(topic_groups.values())
        sorted_topics.sort(key=lambda x: (x["open"], x["total"]), reverse=True)

        top_10_topics = sorted_topics[:10]

        # Top 10 Table Rows
        top_10_rows_html = []
        for rank, item in enumerate(top_10_topics, 1):
            r_code = item["rule_id"]
            tot = item["total"]
            op = item["open"]
            re = item["resolved"]
            r_rate = f"{round(100 * re / (tot or 1), 1)}%"
            share_backlog = f"{round(100 * op / (open_cnt or 1), 1)}%"
            share_val = float(share_backlog.replace("%", ""))

            top_10_rows_html.append(f"""
            <tr class="clickable-row" onclick="openDefectModal('{r_code}')" title="Click to view all {tot:,} issues for {html.escape(item['topic_label'])}">
                <td style="text-align: center; font-weight: 700; color: #1F4E79;">#{rank}</td>
                <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <span class="code-badge">{html.escape(r_code)}</span>
                        <strong style="color: #0F172A; font-size: 0.95rem;">{html.escape(item['topic_label'])}</strong>
                    </div>
                    <div style="font-size: 0.8rem; color: #64748B; margin-top: 4px;">{html.escape(item['description'])}</div>
                </td>
                <td style="color: #475569; font-weight: 500;">{html.escape(item['major_topic_label'])}</td>
                <td style="text-align: right; font-weight: 700;">{tot:,}</td>
                <td style="text-align: right; font-weight: 800; color: #DC2626;">{op:,}</td>
                <td style="text-align: right; font-weight: 600; color: #16A34A;">{re:,}</td>
                <td style="text-align: right; font-weight: 600;">{r_rate}</td>
                <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <div style="flex: 1; background: #E2E8F0; height: 8px; border-radius: 4px; overflow: hidden;">
                            <div style="background: #1F4E79; width: {min(share_val * 3, 100)}%; height: 100%;"></div>
                        </div>
                        <span style="font-size: 0.85rem; font-weight: 700; width: 45px; text-align: right;">{share_backlog}</span>
                    </div>
                </td>
                <td style="text-align: center;">
                    <button class="view-btn" onclick="event.stopPropagation(); openDefectModal('{r_code}')">
                        View Issues ({tot:,}) &rarr;
                    </button>
                </td>
            </tr>
            """)

        # Serialize topics metadata & issues payload to JSON
        topics_meta_json = json.dumps({t["rule_id"]: t for t in sorted_topics})
        issues_data_json = json.dumps(issues_by_rule)

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} — {rep_date_str}</title>
    <style>
        :root {{
            --navy: #0B1F33;
            --primary: #1F4E79;
            --primary-light: #F0F4F8;
            --accent-blue: #2563EB;
            --accent-teal: #0D9488;
            --accent-red: #DC2626;
            --accent-green: #16A34A;
            --accent-amber: #D97706;
            --bg: #F8FAFC;
            --card-bg: #FFFFFF;
            --text-main: #0F172A;
            --text-muted: #64748B;
            --border: #E2E8F0;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Inter, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            line-height: 1.5;
            padding: 24px;
        }}
        .container {{ max-width: 1480px; margin: 0 auto; }}

        /* Header */
        header {{
            background: linear-gradient(135deg, var(--navy) 0%, var(--primary) 100%);
            color: #FFFFFF;
            padding: 28px 36px;
            border-radius: 16px;
            margin-bottom: 24px;
            box-shadow: 0 4px 12px rgba(11, 31, 51, 0.15);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}
        .header-title h1 {{ font-size: 1.75rem; font-weight: 700; letter-spacing: -0.02em; margin-bottom: 6px; }}
        .header-title p {{ color: #E2E8F0; font-size: 0.95rem; }}
        .badge-pill {{
            background: rgba(255, 255, 255, 0.18);
            color: #FFFFFF;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            border: 1px solid rgba(255, 255, 255, 0.25);
        }}

        /* KPI Cards */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}
        .kpi-card {{
            background: var(--card-bg);
            padding: 20px;
            border-radius: 12px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .kpi-title {{ font-size: 0.8rem; color: var(--text-muted); text-transform: uppercase; font-weight: 700; letter-spacing: 0.05em; }}
        .kpi-value {{ font-size: 2.2rem; font-weight: 800; color: var(--primary); margin: 6px 0; }}
        .kpi-card.alert .kpi-value {{ color: var(--accent-red); }}
        .kpi-card.success .kpi-value {{ color: var(--accent-green); }}
        .kpi-subtext {{ font-size: 0.8rem; color: var(--text-muted); }}

        /* Panel */
        .panel {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        }}
        .panel-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 18px;
            flex-wrap: wrap;
            gap: 12px;
        }}
        .panel-title {{
            font-size: 1.3rem;
            font-weight: 700;
            color: var(--primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .panel-subtitle {{ font-size: 0.9rem; color: var(--text-muted); margin-top: 4px; }}

        /* Table */
        .table-wrap {{ overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; }}
        th, td {{ padding: 14px 16px; border-bottom: 1px solid var(--border); font-size: 0.92rem; vertical-align: middle; }}
        th {{ background-color: var(--primary-light); color: var(--primary); font-weight: 700; font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.03em; }}
        
        tr.clickable-row {{ cursor: pointer; transition: background-color 0.15s; }}
        tr.clickable-row:hover {{ background-color: #F1F5F9; }}

        .code-badge {{
            display: inline-block;
            background: #EDE9FE;
            color: #5B21B6;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 700;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            border: 1px solid #DDD6FE;
        }}

        .view-btn {{
            background: var(--primary-light);
            color: var(--primary);
            border: 1px solid #CBD5E1;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }}
        .view-btn:hover {{
            background: var(--primary);
            color: #FFFFFF;
            border-color: var(--primary);
        }}

        /* Modal / Drawer Overlay */
        .modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(15, 23, 42, 0.65);
            backdrop-filter: blur(4px);
            z-index: 9999;
            display: none;
            justify-content: center;
            align-items: center;
            opacity: 0;
            transition: opacity 0.2s ease-in-out;
        }}
        .modal-overlay.active {{
            display: flex;
            opacity: 1;
        }}

        .modal-container {{
            background: #FFFFFF;
            width: 96%;
            max-width: 1440px;
            height: 92vh;
            border-radius: 16px;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.2), 0 10px 10px -5px rgba(0, 0, 0, 0.1);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            animation: modalSlideUp 0.25s ease-out;
        }}
        @keyframes modalSlideUp {{
            from {{ transform: translateY(30px); opacity: 0; }}
            to {{ transform: translateY(0); opacity: 1; }}
        }}

        /* Modal Header */
        .modal-header {{
            background: #0B1F33;
            color: #FFFFFF;
            padding: 20px 28px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .modal-title-wrap h2 {{ font-size: 1.35rem; font-weight: 700; display: flex; align-items: center; gap: 10px; }}
        .modal-title-wrap p {{ color: #94A3B8; font-size: 0.85rem; margin-top: 4px; }}
        .modal-close-btn {{
            background: rgba(255,255,255,0.15);
            color: #FFFFFF;
            border: none;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            font-size: 1.2rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.2s;
        }}
        .modal-close-btn:hover {{ background: rgba(255,255,255,0.3); }}

        /* Modal Toolbar */
        .modal-toolbar {{
            background: #F8FAFC;
            padding: 14px 28px;
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }}
        .modal-search-input {{
            padding: 8px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 0.9rem;
            min-width: 340px;
            outline: none;
        }}
        .modal-search-input:focus {{ border-color: var(--accent-blue); }}
        .filter-btn-group {{ display: flex; gap: 6px; }}
        .filter-chip {{
            padding: 6px 12px;
            border: 1px solid var(--border);
            background: #FFFFFF;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .filter-chip.active {{
            background: var(--primary);
            color: #FFFFFF;
            border-color: var(--primary);
        }}
        .sort-toggle-btn {{
            padding: 6px 12px;
            border: 1px solid var(--border);
            background: #FFFFFF;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .sort-toggle-btn:hover {{ background: #F1F5F9; }}
        .export-btn {{
            background: #16A34A;
            color: #FFFFFF;
            border: none;
            padding: 7px 14px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 700;
            cursor: pointer;
            transition: background 0.2s;
        }}
        .export-btn:hover {{ background: #15803D; }}

        /* Modal Table Body */
        .modal-body {{
            flex: 1;
            overflow-y: auto;
            padding: 0 28px 20px;
        }}
        .issue-table th {{ position: sticky; top: 0; z-index: 10; background: #F1F5F9; font-size: 0.82rem; }}
        .issue-table td {{ font-size: 0.88rem; vertical-align: top; padding: 12px 14px; }}
        .badge-status {{
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
        }}
        .badge-open {{ background: #FEE2E2; color: #991B1B; }}
        .badge-resolved {{ background: #DCFCE7; color: #166534; }}

        .desc-text {{
            font-size: 0.82rem;
            color: #475569;
            margin-top: 6px;
            line-height: 1.4;
            max-height: 110px;
            overflow-y: auto;
            background: #F8FAFC;
            padding: 8px 10px;
            border-radius: 6px;
            border: 1px solid #E2E8F0;
            font-family: inherit;
            white-space: pre-wrap;
        }}

        footer {{
            text-align: center;
            color: var(--text-muted);
            font-size: 0.85rem;
            margin-top: 36px;
            padding-top: 16px;
            border-top: 1px solid var(--border);
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <header>
            <div class="header-title">
                <h1>{title}</h1>
                <p>Interactive Root-Cause Defect Diagnostics &bull; As of {rep_date_str}</p>
            </div>
            <div>
                <span class="badge-pill">Click any category to pop up its issues table</span>
            </div>
        </header>

        <!-- KPI Metric Ribbon -->
        <section class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Total Ingested</div>
                <div class="kpi-value">{total_cnt:,}</div>
                <div class="kpi-subtext">Issues across all modules</div>
            </div>
            <div class="kpi-card alert">
                <div class="kpi-title">Open Backlog</div>
                <div class="kpi-value">{open_cnt:,}</div>
                <div class="kpi-subtext">Requires resolution action</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-title">Resolved / Closed</div>
                <div class="kpi-value">{resolved_cnt:,}</div>
                <div class="kpi-subtext">Resolution Rate: <strong>{res_rate}</strong></div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Distinct Defect Topics</div>
                <div class="kpi-value" style="color: #2563EB;">{len(topic_groups)}</div>
                <div class="kpi-subtext">Calibrated in rules.yaml</div>
            </div>
        </section>

        <!-- Main Panel: Top 10 Root-Cause Defect Drivers -->
        <section class="panel">
            <div class="panel-header">
                <div>
                    <div class="panel-title">System-Wide Top 10 Root-Cause Defect Drivers</div>
                    <div class="panel-subtitle">Ranked by open backlog pendency and total issue volume across all 5,086 tickets.</div>
                </div>
            </div>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 50px; text-align: center;">Rank</th>
                            <th style="width: 380px;">Defect Topic &amp; Rule Code</th>
                            <th>Major Problem Group</th>
                            <th style="text-align: right; width: 90px;">Total</th>
                            <th style="text-align: right; width: 100px;">Open Backlog</th>
                            <th style="text-align: right; width: 90px;">Resolved</th>
                            <th style="text-align: right; width: 100px;">Resolution %</th>
                            <th style="width: 160px;">Share of Backlog</th>
                            <th style="width: 140px; text-align: center;">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(top_10_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <footer>
            CITES Operations Intelligence Platform &bull; Interactive Root-Cause Defect Drilldown &bull; Generated for {rep_date_str}
        </footer>
    </div>

    <!-- Pop-up Modal Container -->
    <div id="defectModalOverlay" class="modal-overlay" onclick="closeDefectModalOnBackdrop(event)">
        <div class="modal-container" onclick="event.stopPropagation()">
            <!-- Modal Header -->
            <div class="modal-header">
                <div class="modal-title-wrap">
                    <h2 id="modalDefectTitle">
                        <span id="modalRuleBadge" class="code-badge">C01_VISIBILITY_DA</span>
                        <span id="modalTopicLabel">At DA level</span>
                    </h2>
                    <p id="modalMajorSubtext">Major Problem Group: Claim/task is not visible or routed</p>
                </div>
                <button class="modal-close-btn" onclick="closeDefectModal()">&times;</button>
            </div>

            <!-- Modal Toolbar -->
            <div class="modal-toolbar">
                <div style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap;">
                    <input type="text" id="modalSearchInput" class="modal-search-input" placeholder="Search by Issue ID, summary, description, category, queue..." onkeyup="filterModalIssues()">
                    <div class="filter-btn-group">
                        <button class="filter-chip active" id="chipAll" onclick="setModalFilter('all')">All (<span id="countAll">0</span>)</button>
                        <button class="filter-chip" id="chipOpen" onclick="setModalFilter('open')">Open (<span id="countOpen">0</span>)</button>
                        <button class="filter-chip" id="chipResolved" onclick="setModalFilter('resolved')">Resolved (<span id="countResolved">0</span>)</button>
                    </div>
                    <button class="sort-toggle-btn" id="sortToggleBtn" onclick="toggleAgeSort()">
                        Sort: <strong id="sortOrderLabel">Latest First (Age &uarr;)</strong>
                    </button>
                </div>
                <div>
                    <button class="export-btn" onclick="exportModalIssuesCSV()">Download CSV</button>
                </div>
            </div>

            <!-- Modal Issue Data Table (Summary & Description is 2nd Column) -->
            <div class="modal-body">
                <table class="issue-table">
                    <thead>
                        <tr>
                            <th style="width: 95px;">Issue ID</th>
                            <th>Summary &amp; Description</th>
                            <th style="width: 150px;">Category</th>
                            <th style="width: 180px;">Assigned Queue</th>
                            <th style="width: 100px;">Status</th>
                            <th style="width: 80px; text-align: center; cursor: pointer;" onclick="toggleAgeSort()" title="Click to toggle Age sort">
                                Age &#x21C5;
                            </th>
                        </tr>
                    </thead>
                    <tbody id="modalIssuesTbody">
                        <!-- Populated dynamically via JS -->
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Embedded Data & Scripts -->
    <script>
        const TOPICS_META = {topics_meta_json};
        const ISSUES_DATA = {issues_data_json};

        let currentRuleId = null;
        let currentStatusFilter = 'all';
        let sortAscendingAge = true; // Latest (youngest age) first by default

        function openDefectModal(ruleId) {{
            currentRuleId = ruleId;
            currentStatusFilter = 'all';
            sortAscendingAge = true;

            const meta = TOPICS_META[ruleId] || {{
                rule_id: ruleId,
                topic_label: ruleId,
                major_topic_label: 'Uncategorized',
                total: 0,
                open: 0,
                resolved: 0
            }};

            document.getElementById('modalRuleBadge').innerText = meta.rule_id;
            document.getElementById('modalTopicLabel').innerText = meta.topic_label;
            document.getElementById('modalMajorSubtext').innerText = `Major Problem Group: ${{meta.major_topic_label}} • Total Issues: ${{meta.total.toLocaleString()}} (Open: ${{meta.open.toLocaleString()}}, Resolved: ${{meta.resolved.toLocaleString()}})`;

            const issues = ISSUES_DATA[ruleId] || [];
            const openIssues = issues.filter(i => !i.is_resolved);
            const resolvedIssues = issues.filter(i => i.is_resolved);

            document.getElementById('countAll').innerText = issues.length.toLocaleString();
            document.getElementById('countOpen').innerText = openIssues.length.toLocaleString();
            document.getElementById('countResolved').innerText = resolvedIssues.length.toLocaleString();
            document.getElementById('modalSearchInput').value = '';
            document.getElementById('sortOrderLabel').innerHTML = 'Latest First (Age &uarr;)';

            document.getElementById('chipAll').classList.add('active');
            document.getElementById('chipOpen').classList.remove('active');
            document.getElementById('chipResolved').classList.remove('active');

            renderModalRows(issues);

            const overlay = document.getElementById('defectModalOverlay');
            overlay.classList.add('active');
            document.body.style.overflow = 'hidden';
        }}

        function closeDefectModal() {{
            const overlay = document.getElementById('defectModalOverlay');
            overlay.classList.remove('active');
            document.body.style.overflow = '';
        }}

        function closeDefectModalOnBackdrop(e) {{
            if (e.target.id === 'defectModalOverlay') {{
                closeDefectModal();
            }}
        }}

        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape') {{
                closeDefectModal();
            }}
        }});

        function setModalFilter(filterType) {{
            currentStatusFilter = filterType;
            document.getElementById('chipAll').classList.toggle('active', filterType === 'all');
            document.getElementById('chipOpen').classList.toggle('active', filterType === 'open');
            document.getElementById('chipResolved').classList.toggle('active', filterType === 'resolved');
            filterModalIssues();
        }}

        function toggleAgeSort() {{
            sortAscendingAge = !sortAscendingAge;
            document.getElementById('sortOrderLabel').innerHTML = sortAscendingAge 
                ? 'Latest First (Age &uarr;)' 
                : 'Oldest First (Age &darr;)';
            filterModalIssues();
        }}

        function filterModalIssues() {{
            if (!currentRuleId) return;
            const query = document.getElementById('modalSearchInput').value.toLowerCase().trim();
            let issues = ISSUES_DATA[currentRuleId] || [];

            let filtered = issues.filter(iss => {{
                if (currentStatusFilter === 'open' && iss.is_resolved) return false;
                if (currentStatusFilter === 'resolved' && !iss.is_resolved) return false;

                if (!query) return true;
                return (
                    iss.id.toLowerCase().includes(query) ||
                    iss.category.toLowerCase().includes(query) ||
                    iss.assigned_to.toLowerCase().includes(query) ||
                    iss.status.toLowerCase().includes(query) ||
                    iss.summary.toLowerCase().includes(query) ||
                    iss.description.toLowerCase().includes(query)
                );
            }});

            // Sort agewise
            filtered.sort((a, b) => {{
                return sortAscendingAge 
                    ? (a.age_days - b.age_days) 
                    : (b.age_days - a.age_days);
            }});

            renderModalRows(filtered);
        }}

        function renderModalRows(issues) {{
            const tbody = document.getElementById('modalIssuesTbody');
            if (!issues || issues.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px; color: #94A3B8;">No matching issues found.</td></tr>';
                return;
            }}

            const rowsHtml = issues.map(iss => {{
                const statusBadge = iss.is_resolved 
                    ? '<span class="badge-status badge-resolved">Resolved</span>' 
                    : '<span class="badge-status badge-open">Open</span>';
                
                const ageHtml = iss.age_days >= 7 
                    ? `<span style="color: #991B1B; font-weight: 700;">${{iss.age_days}}d</span>` 
                    : `<span style="color: #64748B; font-weight: 600;">${{iss.age_days}}d</span>`;

                const descSnippet = iss.description 
                    ? `<div class="desc-text">${{escapeHtml(iss.description)}}</div>` 
                    : '';

                return `
                <tr>
                    <td style="font-weight: 700; color: #1F4E79;">#${{escapeHtml(iss.id)}}</td>
                    <td>
                        <strong style="color: #0F172A; font-size: 0.92rem;">${{escapeHtml(iss.summary)}}</strong>
                        ${{descSnippet}}
                    </td>
                    <td><strong>${{escapeHtml(iss.category)}}</strong></td>
                    <td><code style="font-size: 0.82rem; color: #334155;">${{escapeHtml(iss.assigned_to)}}</code></td>
                    <td>${{statusBadge}} <div style="font-size: 0.75rem; color: #64748B; margin-top: 2px;">${{escapeHtml(iss.status)}}</div></td>
                    <td style="text-align: center;">${{ageHtml}}</td>
                </tr>
                `;
            }}).join('');

            tbody.innerHTML = rowsHtml;
        }}

        function escapeHtml(str) {{
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }}

        function exportModalIssuesCSV() {{
            if (!currentRuleId) return;
            const issues = ISSUES_DATA[currentRuleId] || [];
            if (!issues.length) return;

            let csvContent = 'data:text/csv;charset=utf-8,';
            csvContent += 'Id,Summary,Description,Category,Assigned To,Status,Age Days\\n';

            issues.forEach(iss => {{
                const cleanSummary = (iss.summary || '').replace(/"/g, '""');
                const cleanDesc = (iss.description || '').replace(/"/g, '""').replace(/\\r?\\n/g, ' ');
                const cleanCat = (iss.category || '').replace(/"/g, '""');
                const cleanAssigned = (iss.assigned_to || '').replace(/"/g, '""');
                csvContent += `"${{iss.id}}","${{cleanSummary}}","${{cleanDesc}}","${{cleanCat}}","${{cleanAssigned}}","${{iss.status}}",${{iss.age_days}}\\n`;
            }});

            const encodedUri = encodeURI(csvContent);
            const link = document.createElement('a');
            link.setAttribute('href', encodedUri);
            link.setAttribute('download', `CITES_Issues_${{currentRuleId}}_${{new Date().toISOString().slice(0,10)}}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}
    </script>
</body>
</html>
"""

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        return str(out_file)
