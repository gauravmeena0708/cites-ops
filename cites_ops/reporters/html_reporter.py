from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd

class HTMLReporter:
    """
    Generates self-contained, offline HTML management dashboards
    with responsive styling and interactive tables.
    Supports Jinja2 if installed, with a robust fallback to pure Python templating.
    """

    @classmethod
    def generate_html(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        report_date: Optional[Union[str, date]] = None,
        title: str = "CITES Operations Intelligence Dashboard",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        total = len(df_classified)
        status_series = df_classified["Status"].astype(str).str.lower() if "Status" in df_classified.columns else pd.Series([])
        resolved = int(status_series.isin(["resolved", "fixed", "closed"]).sum())
        open_cnt = total - resolved
        rate = f"{round(100 * resolved / total, 1)}%" if total else "0%"
        rep_date_str = str(report_date or date.today())

        # Category breakdown rows
        cat_rows_html = []
        if "major_topic_label" in df_classified.columns:
            vc = df_classified["major_topic_label"].value_counts()
            for cat_label, count in vc.items():
                desc = df_classified[df_classified["major_topic_label"] == cat_label]["category_description"].iloc[0] if "category_description" in df_classified.columns else ""
                pct = f"{round(100 * count / (total or 1), 1)}%"
                cat_rows_html.append(
                    f"<tr><td><b>{cat_label}</b></td>"
                    f"<td style='color: #64748B; font-size: 0.85rem;'>{desc}</td>"
                    f"<td style='text-align: right; font-weight: 600;'>{count:,}</td>"
                    f"<td style='text-align: right;'>{pct}</td></tr>"
                )

        # Aging sample rows
        aging_rows_html = []
        if "age_days" in df_classified.columns:
            aging_df = df_classified[df_classified["age_days"] >= 7].sort_values(by="age_days", ascending=False).head(15)
            for _, row in aging_df.iterrows():
                iss_id = row.get("Id", "")
                cat = row.get("Category", "")
                assigned = row.get("Assigned To", "")
                age = row.get("age_days", 0)
                st = row.get("Status", "")
                aging_rows_html.append(
                    f"<tr><td><b>{iss_id}</b></td><td>{cat}</td><td><code>{assigned}</code></td>"
                    f"<td><span class='badge' style='background: #FEE2E2; color: #991B1B;'>{age} days</span></td>"
                    f"<td><span class='badge'>{st}</span></td></tr>"
                )

        aging_section = ""
        if aging_rows_html:
            aging_section = f"""
            <section class="panel">
                <div class="panel-title">Queues Requiring Attention (Aging &gt; 7 Days)</div>
                <table>
                    <thead>
                        <tr>
                            <th>Issue ID</th>
                            <th>Category</th>
                            <th>Assigned Queue</th>
                            <th>Age (Days)</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(aging_rows_html)}
                    </tbody>
                </table>
            </section>
            """

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} - {rep_date_str}</title>
    <style>
        :root {{
            --primary: #1F4E79;
            --primary-light: #EBF1F5;
            --accent-red: #C00000;
            --accent-green: #2E7D32;
            --bg: #F8FAFC;
            --card-bg: #FFFFFF;
            --text-main: #1E293B;
            --text-muted: #64748B;
            --border: #E2E8F0;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            line-height: 1.5;
            padding: 2rem;
        }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        header {{
            margin-bottom: 2rem;
            padding-bottom: 1rem;
            border-bottom: 2px solid var(--border);
        }}
        .brand {{ font-size: 1.5rem; font-weight: 700; color: var(--primary); }}
        .subtitle {{ color: var(--text-muted); font-size: 0.95rem; margin-top: 0.25rem; }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }}
        .kpi-card {{
            background: var(--card-bg);
            padding: 1.25rem;
            border-radius: 8px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}
        .kpi-title {{ font-size: 0.85rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600; }}
        .kpi-value {{ font-size: 2rem; font-weight: 700; color: var(--primary); margin: 0.25rem 0; }}
        .kpi-card.alert .kpi-value {{ color: var(--accent-red); }}
        .kpi-card.success .kpi-value {{ color: var(--accent-green); }}

        .panel {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}
        .panel-title {{ font-size: 1.2rem; font-weight: 600; color: var(--primary); margin-bottom: 1rem; }}

        table {{ width: 100%; border-collapse: collapse; text-align: left; }}
        th, td {{ padding: 0.75rem 1rem; border-bottom: 1px solid var(--border); font-size: 0.95rem; }}
        th {{ background-color: var(--primary-light); color: var(--primary); font-weight: 600; }}
        tr:hover {{ background-color: #F1F5F9; }}
        .badge {{
            display: inline-block;
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            font-size: 0.8rem;
            font-weight: 600;
            background: #E2E8F0;
            color: #334155;
        }}
        footer {{ text-align: center; color: var(--text-muted); font-size: 0.85rem; margin-top: 3rem; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="brand">{title}</div>
            <div class="subtitle">Daily Snapshot &amp; Operational Intelligence · As of {rep_date_str}</div>
        </header>

        <section class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Total Ingested</div>
                <div class="kpi-value">{total:,}</div>
            </div>
            <div class="kpi-card alert">
                <div class="kpi-title">Open Backlog</div>
                <div class="kpi-value">{open_cnt:,}</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-title">Resolved / Closed</div>
                <div class="kpi-value">{resolved:,}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Resolution Rate</div>
                <div class="kpi-value">{rate}</div>
            </div>
        </section>

        <section class="panel">
            <div class="panel-title">Major Issue Categories Distribution</div>
            <table>
                <thead>
                    <tr>
                        <th>Major Problem Category</th>
                        <th>Description</th>
                        <th style="text-align: right;">Total Issues</th>
                        <th style="text-align: right;">Share</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(cat_rows_html)}
                </tbody>
            </table>
        </section>

        {aging_section}

        <footer>
            CITES Operations Intelligence Platform · Automated Offline Report
        </footer>
    </div>
</body>
</html>
"""

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        return str(out_file)
