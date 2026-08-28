import html
import json
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd

class HTMLReporter:
    """
    Generates self-contained, offline interactive HTML management dashboards
    with responsive styling, 5-tier topical workload hierarchy drilldown
    (JD -> DD -> Handler -> Category -> Specific Problem Topic),
    flat Top 10 category analysis, cross-module defect matrix, and aging exception registers.
    """

    @classmethod
    def generate_html(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        workload_data: Optional[Dict[str, Any]] = None,
        report_date: Optional[Union[str, date]] = None,
        title: str = "CITES Operations Intelligence & Topical Dashboard",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        rep_date_str = str(report_date or date.today())
        total = len(df_classified)
        
        status_series = df_classified["Status"].astype(str).str.lower() if "Status" in df_classified.columns else pd.Series([])
        resolved = int(status_series.isin(["resolved", "fixed", "closed"]).sum())
        open_cnt = total - resolved
        rate = f"{round(100 * resolved / total, 1)}%" if total else "0%"

        # Workload KPIs
        wk_kpis = (workload_data or {}).get("kpis", {})
        routing_breakdown = wk_kpis.get("routing_breakdown", {})
        epfo_cnt = routing_breakdown.get("internal_tech", total)
        cdac_cnt = routing_breakdown.get("vendor_tech", 0)
        field_cnt = routing_breakdown.get("field_office", 0)
        cov_pct = wk_kpis.get("coverage_pct", "100%")

        # 1. Flat Top 10 Categories HTML Rows (with Top Problem Topics tags)
        top_10_cats = (workload_data or {}).get("top_10_categories", [])
        top_10_rows_html = []
        for cat in top_10_cats:
            share_val = float(str(cat.get("share_of_backlog", "0")).replace("%", ""))
            
            # Render top 2 topical driver tags for this category
            topic_tags = []
            for t_item in cat.get("top_topics", [])[:2]:
                topic_tags.append(
                    f"<span class='mini-topic-pill' title='{html.escape(t_item['topic_label'])} ({t_item['total']} issues)'>"
                    f"<code>{html.escape(t_item['rule_id'])}</code> {html.escape(t_item['topic_label'][:28])}.. "
                    f"<strong>({t_item['share_of_module']})</strong></span>"
                )
            topic_tags_html = "".join(topic_tags) if topic_tags else "<span style='color:#94A3B8; font-size:0.8rem;'>No specific topics</span>"

            top_10_rows_html.append(f"""
            <tr>
                <td style="text-align: center; font-weight: 700; color: #1F4E79;">#{cat.get('rank', '-')}</td>
                <td>
                    <span style="font-weight: 700; color: #0F172A; font-size: 0.95rem;">{html.escape(str(cat.get('category', '')))}</span>
                    <div style="margin-top: 4px; display: flex; gap: 4px; flex-wrap: wrap;">
                        {topic_tags_html}
                    </div>
                </td>
                <td style="font-weight: 600; text-align: right;">{cat.get('total', 0):,}</td>
                <td style="font-weight: 700; color: #DC2626; text-align: right;">{cat.get('open', 0):,}</td>
                <td style="font-weight: 600; color: #16A34A; text-align: right;">{cat.get('resolved', 0):,}</td>
                <td style="text-align: right;">{cat.get('resolution_rate', '0%')}</td>
                <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <div style="flex: 1; background: #E2E8F0; height: 8px; border-radius: 4px; overflow: hidden;">
                            <div style="background: #1F4E79; width: {min(share_val * 2.5, 100)}%; height: 100%;"></div>
                        </div>
                        <span style="font-size: 0.85rem; font-weight: 600; width: 45px; text-align: right;">{cat.get('share_of_backlog', '0%')}</span>
                    </div>
                </td>
                <td>
                    <div style="font-size: 0.85rem; line-height: 1.3;">
                        <strong>{html.escape(str(cat.get('handler', '')))}</strong><br>
                        <span style="color: #64748B; font-size: 0.8rem;">DD: {html.escape(str(cat.get('dd', '')))} &bull; JD: {html.escape(str(cat.get('jd', '')))}</span>
                    </div>
                </td>
            </tr>
            """)

        # 2. 5-Tier Hierarchy Table Rows (JD -> DD -> Handler -> Category -> Specific Topic)
        hierarchy_rows = (workload_data or {}).get("hierarchy_rows", [])
        hier_rows_html = []
        for r in hierarchy_rows:
            depth = r.get("depth", 0)
            level_lbl = r.get("level_label", "")
            name_val = r.get("name", "")
            r_id = r.get("id", "")
            p_id = r.get("parent_id", "root")
            has_child = r.get("has_children", False)

            badge_class = f"badge-level-{min(depth, 4)}"
            indent_px = depth * 20
            toggle_icon = '<span class="tree-icon">&#9660;</span>' if has_child else '<span class="tree-icon-leaf">&bull;</span>'

            hier_rows_html.append(f"""
            <tr class="tree-row depth-{min(depth, 4)}" data-id="{r_id}" data-parent="{p_id}" data-depth="{depth}">
                <td>
                    <span class="level-badge {badge_class}">{html.escape(level_lbl)}</span>
                </td>
                <td style="padding-left: {indent_px + 12}px;">
                    <div class="tree-node-cell" onclick="toggleTreeNode('{r_id}')">
                        {toggle_icon}
                        <span class="node-name depth-font-{min(depth, 4)}">{html.escape(name_val)}</span>
                    </div>
                </td>
                <td style="text-align: right; font-weight: 600;">{r.get('total', 0):,}</td>
                <td style="text-align: right; font-weight: 700; color: #DC2626;">{r.get('open', 0):,}</td>
                <td style="text-align: right; font-weight: 600; color: #16A34A;">{r.get('resolved', 0):,}</td>
                <td style="text-align: right;">{r.get('resolution_rate', '0%')}</td>
                <td style="text-align: right; font-size: 0.85rem; color: #475569;">
                    <span title="CDAC Vendor Cases">{r.get('cdac_count', 0):,}</span> / 
                    <span title="Field Office RO Cases">{r.get('ro_count', 0):,}</span>
                </td>
            </tr>
            """)

        # 3. Cross-Tab Matrix: Top 10 Functionalities x Major Problem Groups
        cross_tab = (workload_data or {}).get("cross_tab_matrix", {})
        major_cols = (workload_data or {}).get("major_columns", [])
        
        matrix_header_html = ["<th style='position: sticky; left: 0; z-index: 2; background: #EBF1F5;'>Functionality / Module</th>"]
        for m_col in major_cols:
            short_col = m_col.replace("Claim/task is not visible or routed", "Visibility / Routing") \
                             .replace("Record, service or data availability", "Data / Record Availability") \
                             .replace("Workflow actions and claim processing", "Workflow / Processing") \
                             .replace("Eligibility, validation and status conflicts", "Eligibility / Status Conflict") \
                             .replace("Financial, benefit and ledger discrepancies", "Financial / Ledger") \
                             .replace("Document generation and digital signing", "Document / DSC / eSign") \
                             .replace("Identity, KYC and data correction", "Identity / KYC / Correction") \
                             .replace("Login, portal and system availability", "Login / System Access") \
                             .replace("Legacy migration and synchronisation", "Legacy Migration / Sync") \
                             .replace("Other or insufficient detail", "Other / Insufficient Detail")
            matrix_header_html.append(f"<th style='text-align: center; font-size: 0.78rem;' title='{html.escape(m_col)}'>{html.escape(short_col)}</th>")
        matrix_header_html.append("<th style='text-align: right;'>Module Total</th>")

        matrix_rows_html = []
        for cat_name, col_dict in cross_tab.items():
            row_sum = sum(col_dict.values())
            cells = [f"<td style='font-weight: 700; position: sticky; left: 0; background: #FFFFFF; border-right: 2px solid #E2E8F0;'>{html.escape(str(cat_name))}</td>"]
            for m_col in major_cols:
                v = col_dict.get(m_col, 0)
                # Heatmap background intensity based on count
                bg_color = "transparent"
                txt_color = "#0F172A"
                if v > 100:
                    bg_color = "#FEE2E2"
                    txt_color = "#991B1B"
                elif v > 40:
                    bg_color = "#FEF3C7"
                    txt_color = "#92400E"
                elif v > 10:
                    bg_color = "#EFF6FF"
                    txt_color = "#1E40AF"
                elif v == 0:
                    txt_color = "#94A3B8"

                cells.append(f"<td style='text-align: center; background: {bg_color}; color: {txt_color}; font-weight: {700 if v > 20 else 500};'>{v if v > 0 else '-'}</td>")
            cells.append(f"<td style='text-align: right; font-weight: 700; background: #F8FAFC;'>{row_sum:,}</td>")
            matrix_rows_html.append(f"<tr>{''.join(cells)}</tr>")

        # 4. System-Wide Top Systemic Root-Cause Defect Drivers
        top_systemic = (workload_data or {}).get("top_systemic_defects", [])
        systemic_rows_html = []
        for idx, item in enumerate(top_systemic, 1):
            systemic_rows_html.append(f"""
            <tr>
                <td style="text-align: center; font-weight: 700; color: #1F4E79;">#{idx}</td>
                <td>
                    <span class="level-badge badge-level-4" style="margin-right: 6px;">{html.escape(item['rule_id'])}</span>
                    <strong style="color: #0F172A;">{html.escape(item['topic_label'])}</strong>
                    <div style="font-size: 0.8rem; color: #64748B; margin-top: 2px;">{html.escape(item['description'])}</div>
                </td>
                <td style="font-weight: 600; color: #475569;">{html.escape(item['major_topic_label'])}</td>
                <td style="text-align: right; font-weight: 600;">{item['total']:,}</td>
                <td style="text-align: right; font-weight: 700; color: #DC2626;">{item['open']:,}</td>
                <td style="text-align: right; font-weight: 600; color: #16A34A;">{item['resolved']:,}</td>
                <td style="text-align: right; font-weight: 600;">{item['share_of_total']}</td>
            </tr>
            """)

        # 5. Aging Exceptions Rows (> 7 days)
        aging_rows_html = []
        if "age_days" in df_classified.columns:
            aging_df = df_classified[df_classified["age_days"] >= 7].sort_values(by="age_days", ascending=False).head(20)
            for _, row in aging_df.iterrows():
                iss_id = row.get("Id", "")
                cat = row.get("Category", "")
                assigned = row.get("Assigned To", "")
                age = row.get("age_days", 0)
                st = row.get("Status", "")
                summ = row.get("Summary", "")
                badge_style = "background: #FEE2E2; color: #991B1B;" if age >= 15 else "background: #FEF3C7; color: #92400E;"

                aging_rows_html.append(f"""
                <tr>
                    <td style="font-weight: 700; color: #1F4E79;">{html.escape(str(iss_id))}</td>
                    <td><span class="badge" style="{badge_style}">{age} days</span></td>
                    <td><strong>{html.escape(str(cat))}</strong></td>
                    <td><code style="font-size: 0.85rem;">{html.escape(str(assigned))}</code></td>
                    <td><span class="badge">{html.escape(str(st))}</span></td>
                    <td style="font-size: 0.85rem; color: #334155;">{html.escape(str(summ)[:120])}...</td>
                </tr>
                """)

        html_template = f"""<!DOCTYPE html>
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
            --border-subtle: #F1F5F9;
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
        .header-nav {{ display: flex; gap: 8px; flex-wrap: wrap; }}
        .nav-btn {{
            background: rgba(255, 255, 255, 0.15);
            color: #FFFFFF;
            text-decoration: none;
            padding: 8px 14px;
            border-radius: 8px;
            font-size: 0.85rem;
            font-weight: 600;
            transition: all 0.2s;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }}
        .nav-btn:hover {{ background: rgba(255, 255, 255, 0.3); }}

        /* KPI Cards */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
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

        /* Panels */
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
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .panel-subtitle {{ font-size: 0.9rem; color: var(--text-muted); margin-top: 4px; }}

        /* Controls */
        .table-controls {{
            display: flex;
            gap: 10px;
            align-items: center;
            flex-wrap: wrap;
        }}
        .search-input {{
            padding: 8px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 0.9rem;
            min-width: 260px;
            outline: none;
            transition: border 0.2s;
        }}
        .search-input:focus {{ border-color: var(--accent-blue); }}
        .action-btn {{
            padding: 8px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            background: #FFFFFF;
            color: var(--text-main);
            font-size: 0.85rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .action-btn:hover {{ background: #F1F5F9; border-color: #CBD5E1; }}

        /* Tables */
        .table-wrap {{ overflow-x: auto; max-width: 100%; overscroll-behavior-x: contain; }}
        table {{ width: 100%; min-width: 980px; border-collapse: collapse; text-align: left; }}
        th, td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: 0.90rem; vertical-align: middle; overflow-wrap: anywhere; }}
        th {{ background-color: var(--primary-light); color: var(--primary); font-weight: 700; font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.03em; position: sticky; top: 0; }}
        tr:hover {{ background-color: #F8FAFC; }}

        /* Hierarchy Level Badges */
        .level-badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.72rem;
            font-weight: 700;
            letter-spacing: 0.02em;
            white-space: nowrap;
        }}
        .badge-level-0 {{ background: #EDE9FE; color: #5B21B6; border: 1px solid #DDD6FE; }} /* JD: Purple */
        .badge-level-1 {{ background: #CCFBF1; color: #115E59; border: 1px solid #99F6E4; }} /* DD: Teal */
        .badge-level-2 {{ background: #E0E7FF; color: #3730A3; border: 1px solid #C7D2FE; }} /* Handler: Indigo */
        .badge-level-3 {{ background: #E2E8F0; color: #1E293B; border: 1px solid #CBD5E1; }} /* Category: Slate */
        .badge-level-4 {{ background: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; }} /* Topic: Amber */

        /* Tree Structure Styling */
        .tree-node-cell {{ display: flex; align-items: center; gap: 8px; cursor: pointer; user-select: none; min-width: 0; max-width: 100%; }}
        .tree-node-cell .node-name {{ min-width: 0; overflow-wrap: anywhere; word-break: break-word; }}
        .tree-icon {{ font-size: 0.75rem; color: var(--text-muted); width: 14px; text-align: center; transition: transform 0.2s; }}
        .tree-icon-leaf {{ font-size: 0.9rem; color: #94A3B8; width: 14px; text-align: center; }}
        .depth-font-0 {{ font-size: 1.02rem; font-weight: 700; color: #1E1B4B; }}
        .depth-font-1 {{ font-size: 0.96rem; font-weight: 600; color: #042F2E; }}
        .depth-font-2 {{ font-size: 0.92rem; font-weight: 600; color: #1E293B; }}
        .depth-font-3 {{ font-size: 0.90rem; font-weight: 700; color: #0F172A; }}
        .depth-font-4 {{ font-size: 0.86rem; font-weight: 500; color: #475569; }}

        tr.depth-0 {{ background-color: #FAF5FF; }}
        tr.depth-1 {{ background-color: #F0FDFA; }}
        tr.depth-2 {{ background-color: #FAFAFA; }}
        tr.depth-3 {{ background-color: #FFFFFF; font-weight: 600; }}
        tr.depth-4 {{ background-color: #FFFDF5; }}

        .mini-topic-pill {{
            display: inline-block;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 0.75rem;
            background: #F1F5F9;
            color: #334155;
            border: 1px solid #E2E8F0;
        }}
        .mini-topic-pill code {{ color: #1F4E79; font-weight: 700; }}

        .badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 0.8rem;
            font-weight: 600;
            background: #E2E8F0;
            color: #334155;
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
                <p>Operations, Accountability &amp; Topical Defect Intelligence &bull; As of {rep_date_str}</p>
            </div>
            <div class="header-nav">
                <a href="#top-categories" class="nav-btn">Top 10 Modules</a>
                <a href="#workforce-hierarchy" class="nav-btn">Workforce &amp; Topical Hierarchy</a>
                <a href="#defect-matrix" class="nav-btn">Cross-Module Matrix</a>
                <a href="#systemic-defects" class="nav-btn">Top Defect Drivers</a>
                <a href="#aging-register" class="nav-btn">Aging Alerts</a>
            </div>
        </header>

        <!-- KPI Metric Ribbon -->
        <section class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Total Ingested</div>
                <div class="kpi-value">{total:,}</div>
                <div class="kpi-subtext">Issues across all modules</div>
            </div>
            <div class="kpi-card alert">
                <div class="kpi-title">Open Backlog</div>
                <div class="kpi-value">{open_cnt:,}</div>
                <div class="kpi-subtext">Requires resolution action</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-title">Resolved / Closed</div>
                <div class="kpi-value">{resolved:,}</div>
                <div class="kpi-subtext">Resolution Rate: <strong>{rate}</strong></div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">EPFO Internal Queues</div>
                <div class="kpi-value" style="color: #2563EB;">{epfo_cnt:,}</div>
                <div class="kpi-subtext">Core tech team responsibility</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">CDAC + Field Offices</div>
                <div class="kpi-value" style="color: #0D9488;">{cdac_cnt + field_cnt:,}</div>
                <div class="kpi-subtext">CDAC: {cdac_cnt:,} &bull; RO Queues: {field_cnt:,}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Ownership Coverage</div>
                <div class="kpi-value" style="color: #16A34A;">{cov_pct}</div>
                <div class="kpi-subtext">Mapped in Issue_teams.csv</div>
            </div>
        </section>

        <!-- Section 1: Top 10 Major Problem Categories (Functionalities + Top Topical Drivers) -->
        <section class="panel" id="top-categories">
            <div class="panel-header">
                <div>
                    <div class="panel-title">Top 10 Major Problem Categories (Functionalities)</div>
                    <div class="panel-subtitle">Ranked by open backlog volume, share of total pendency, top defect drivers, and accountable leadership.</div>
                </div>
            </div>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 50px; text-align: center;">Rank</th>
                            <th style="width: 320px;">Category / Module &amp; Top Defect Drivers</th>
                            <th style="text-align: right; width: 90px;">Total</th>
                            <th style="text-align: right; width: 100px;">Open Backlog</th>
                            <th style="text-align: right; width: 90px;">Resolved</th>
                            <th style="text-align: right; width: 100px;">Resolution %</th>
                            <th style="width: 150px;">Share of Backlog</th>
                            <th>Accountable Official &amp; Leadership</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(top_10_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Section 2: 5-Tier Interactive People, Category & Topical Workload Hierarchy -->
        <section class="panel" id="workforce-hierarchy">
            <div class="panel-header">
                <div>
                    <div class="panel-title">Workforce &amp; Topical Defect Hierarchy Drilldown</div>
                    <div class="panel-subtitle">Interactive 5-tier accountability drilldown: <strong>JD(IS) &rarr; DD(IS) &rarr; AO/EO/DPA/Programmer &rarr; Category &rarr; Specific Problem Defect</strong></div>
                </div>
                <div class="table-controls">
                    <input type="text" id="treeSearchInput" class="search-input" placeholder="Search official, JD, DD, category, or defect topic..." onkeyup="filterTreeTable()">
                    <button class="action-btn" onclick="expandAllTree()">Expand All</button>
                    <button class="action-btn" onclick="collapseAllTree()">Collapse All</button>
                </div>
            </div>
            <div class="table-wrap">
                <table id="hierarchyTable">
                    <thead>
                        <tr>
                            <th style="width: 160px;">Tier Level</th>
                            <th>Name of Official / Category / Defect Topic</th>
                            <th style="text-align: right; width: 110px;">Total Issues</th>
                            <th style="text-align: right; width: 110px;">Open Backlog</th>
                            <th style="text-align: right; width: 100px;">Resolved</th>
                            <th style="text-align: right; width: 110px;">Resolution %</th>
                            <th style="text-align: right; width: 130px;">CDAC / RO Cases</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(hier_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Section 3: Cross-Module Defect Heatmap / Pivot Matrix -->
        <section class="panel" id="defect-matrix">
            <div class="panel-header">
                <div>
                    <div class="panel-title">Cross-Module Defect Distribution Matrix</div>
                    <div class="panel-subtitle">Top 10 Functionalities &times; 10 Major Defect Groups (identifying systemic vs module-specific anomalies).</div>
                </div>
            </div>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            {''.join(matrix_header_html)}
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(matrix_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Section 4: System-Wide Top 10 Root-Cause Defect Drivers -->
        <section class="panel" id="systemic-defects">
            <div class="panel-header">
                <div>
                    <div class="panel-title">System-Wide Top 10 Root-Cause Defect Drivers</div>
                    <div class="panel-subtitle">Overall defect categories ranked across all 5,086 tickets derived via deterministic text analysis (rules.yaml).</div>
                </div>
            </div>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 60px; text-align: center;">Rank</th>
                            <th style="width: 320px;">Defect Topic &amp; Rule Code</th>
                            <th>Major Problem Group</th>
                            <th style="text-align: right; width: 110px;">Total Issues</th>
                            <th style="text-align: right; width: 110px;">Open Backlog</th>
                            <th style="text-align: right; width: 100px;">Resolved</th>
                            <th style="text-align: right; width: 120px;">% of Total Issues</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(systemic_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Section 5: Aging Register & Escalations -->
        <section class="panel" id="aging-register">
            <div class="panel-header">
                <div>
                    <div class="panel-title">Daily Action &amp; Aging Exceptions Register (7+ Days)</div>
                    <div class="panel-subtitle">Issues requiring immediate management intervention due to prolonged pendency.</div>
                </div>
            </div>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 90px;">Issue ID</th>
                            <th style="width: 110px;">Age</th>
                            <th style="width: 180px;">Category</th>
                            <th style="width: 200px;">Assigned Queue</th>
                            <th style="width: 100px;">Status</th>
                            <th>Summary of Issue</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(aging_rows_html)}
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Footer -->
        <footer>
            CITES Operations Intelligence Platform &bull; Automated Decision Support &amp; Topical Intelligence &bull; Generated for {rep_date_str}
        </footer>
    </div>

    <!-- Interactive JavaScript -->
    <script>
        function toggleTreeNode(nodeId) {{
            const rows = document.querySelectorAll('#hierarchyTable tbody tr');
            const targetRow = document.querySelector(`tr[data-id="${{nodeId}}"]`);
            if (!targetRow) return;

            const icon = targetRow.querySelector('.tree-icon');
            const isExpanded = icon && icon.innerHTML === '▼';

            if (icon) {{
                icon.innerHTML = isExpanded ? '▶' : '▼';
            }}

            rows.forEach(row => {{
                const parentId = row.getAttribute('data-parent');
                if (parentId && parentId.startsWith(nodeId)) {{
                    if (isExpanded) {{
                        row.style.display = 'none';
                        const childIcon = row.querySelector('.tree-icon');
                        if (childIcon) childIcon.innerHTML = '▶';
                    }} else {{
                        if (parentId === nodeId) {{
                            row.style.display = '';
                        }}
                    }}
                }}
            }});
        }}

        function expandAllTree() {{
            const rows = document.querySelectorAll('#hierarchyTable tbody tr');
            rows.forEach(row => {{
                row.style.display = '';
                const icon = row.querySelector('.tree-icon');
                if (icon) icon.innerHTML = '▼';
            }});
        }}

        function collapseAllTree() {{
            const rows = document.querySelectorAll('#hierarchyTable tbody tr');
            rows.forEach(row => {{
                const depth = parseInt(row.getAttribute('data-depth') || '0', 10);
                const icon = row.querySelector('.tree-icon');
                if (depth === 0) {{
                    row.style.display = '';
                    if (icon) icon.innerHTML = '▶';
                }} else {{
                    row.style.display = 'none';
                    if (icon) icon.innerHTML = '▶';
                }}
            }});
        }}

        function filterTreeTable() {{
            const query = document.getElementById('treeSearchInput').value.toLowerCase().trim();
            const rows = document.querySelectorAll('#hierarchyTable tbody tr');

            if (!query) {{
                expandAllTree();
                return;
            }}

            rows.forEach(row => {{
                const text = row.textContent.toLowerCase();
                if (text.includes(query)) {{
                    row.style.display = '';
                    let pId = row.getAttribute('data-parent');
                    while (pId && pId !== 'root') {{
                        const pRow = document.querySelector(`tr[data-id="${{pId}}"]`);
                        if (pRow) {{
                            pRow.style.display = '';
                            const icon = pRow.querySelector('.tree-icon');
                            if (icon) icon.innerHTML = '▼';
                            pId = pRow.getAttribute('data-parent');
                        }} else {{
                            break;
                        }}
                    }}
                }} else {{
                    row.style.display = 'none';
                }}
            }});
        }}
    </script>
</body>
</html>
"""

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(html_template)

        return str(out_file)
