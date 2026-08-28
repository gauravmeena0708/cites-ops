"""Reporter for generating CITES Resolved & Closed Weekly Transition HTML applications."""

from __future__ import annotations

import json
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional, Union, Dict, Any, List
import pandas as pd

from ..core.stats_parser import StatsDocxParser
from ..core.workforce import WorkforceMapper


class WeeklyResolutionsReporter:
    """Generates standalone, interactive CITES Resolved + Closed Weekly Transition HTML dashboards."""

    HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CITES — Resolved + closed during the week</title>
<style>
:root{--navy:#0b1f33;--ink:#172a3a;--muted:#647584;--paper:#f4f7fb;--line:#d9e2ea;--white:#fff;--blue:#326bff;--teal:#168c7c;--green:#137d68;--purple:#6740aa;--amber:#d98a0b;--shadow:0 7px 25px rgba(11,31,51,.08)}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.45 "Segoe UI",Aptos,Arial,sans-serif}.top{background:var(--navy);color:#fff;padding:22px max(24px,calc((100vw - 1500px)/2))}.top small{color:#a9bdce}.top h1{margin:3px 0 2px;font-size:25px}.wrap{max-width:1500px;margin:auto;padding:25px}.panel{background:#fff;border:1px solid var(--line);border-radius:11px;box-shadow:var(--shadow);padding:19px}.head{display:flex;justify-content:space-between;align-items:flex-start;gap:18px;margin-bottom:14px}.head h2{margin:0 0 4px;font-size:19px}.muted{color:var(--muted)}.nav{display:flex;align-items:center;gap:8px;white-space:nowrap}.nav button,.controls button,.controls select{border:0;border-radius:8px;background:#e8eef5;color:var(--ink);padding:8px 12px;font:inherit;font-weight:700;cursor:pointer}.nav button.arrow{background:var(--blue);color:#fff;font-size:18px;padding:5px 13px}.nav button:disabled{opacity:.35;cursor:not-allowed}.status-kpis{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:11px;margin-bottom:13px}.status-kpis article{position:relative;overflow:hidden;border-radius:9px;color:#fff;padding:14px 16px;min-height:94px;box-shadow:0 5px 16px rgba(11,31,51,.12)}.status-kpis article:after{content:'';position:absolute;width:75px;height:75px;border-radius:50%;background:rgba(255,255,255,.09);right:-18px;bottom:-25px}.status-kpis span,.status-kpis small{display:block;color:rgba(255,255,255,.82)}.status-kpis span{font-size:12px;font-weight:700}.status-kpis strong{display:block;font-size:28px;line-height:1.1;margin:7px 0 3px}.kpi-total{background:var(--navy)}.kpi-resolved{background:var(--green)}.kpi-yesterday{background:var(--purple)}.kpi-open{background:var(--amber)}.controls{display:flex;justify-content:flex-end;gap:7px;margin-bottom:10px}.scroll{overflow-x:auto;overscroll-behavior-x:contain;border:1px solid var(--line);border-radius:8px;max-width:100%}.grid{display:grid;grid-template-columns:100px minmax(280px,1.8fr) repeat(3,minmax(105px,.62fr)) repeat(5,minmax(120px,.7fr));align-items:stretch;min-width:1360px;background:#fff;border-bottom:1px solid #e6edf2}.grid>span,.grid>b{padding:10px 11px;display:flex;align-items:center}.grid>span:not(:first-child){justify-content:flex-end;text-align:right}.grid>b{text-align:left}.header{background:var(--navy);color:#fff;text-transform:uppercase;font-size:11px;letter-spacing:.04em;position:sticky;top:0;z-index:8}.header small{display:block;color:#a9bdce;text-transform:none;letter-spacing:0;margin-left:5px}.status-total-cell{color:var(--navy);font-weight:800}.status-open-cell{color:#a56400;font-weight:800}.status-resolved-cell{color:var(--green);font-weight:800}.daily-low{background:#fff0f1!important;color:#a61b2b;font-weight:800;box-shadow:inset 0 0 0 1px #f4c8cd}.header .status-total-cell,.header .status-open-cell,.header .status-resolved-cell{color:#fff}.total{background:#edf3ff;font-weight:800}.node>summary{cursor:pointer;list-style:none}.node>summary::-webkit-details-marker{display:none}.node>summary>b:before{content:'›';display:inline-block;margin-right:7px;color:var(--blue);font-size:18px;transition:transform .15s}.node[open]>summary>b:before{transform:rotate(90deg)}.node summary:hover,.team:hover{background:#f6f9ff}.children{margin-left:18px;border-left:2px solid #dfe7ee;padding-left:7px;min-width:1360px}.children .grid{min-width:calc(1360px - 25px)}.grid>span:first-child,.grid>b{position:sticky;z-index:2;background:inherit}.grid>span:first-child{left:0}.grid>b{left:100px}.header>span:first-child,.header>b{z-index:9;background:var(--navy)}.level-dd{background:#f9fbfd}.level-official{background:#fcfdff}.team{background:#fff}.foot{display:flex;justify-content:space-between;gap:20px;color:var(--muted);font-size:11px;margin-top:12px}.empty{text-align:center;padding:35px;color:var(--muted)}
.header>span:not(:first-child){display:grid;justify-content:end}.header small{margin:0}.grid{grid-template-columns:100px minmax(280px,1.8fr) minmax(115px,.68fr) repeat(3,minmax(105px,.62fr)) repeat(5,minmax(120px,.7fr));width:1480px;min-width:1480px;max-width:1480px}.node,.node>summary,.children{width:1480px;min-width:1480px;max-width:1480px}.children{margin-left:0;border-left:0;padding-left:0}.children .grid{width:1480px;min-width:1480px;max-width:1480px}.depth-2>.grid>b,.depth-2>summary>b{padding-left:28px}.depth-3>.grid>b,.depth-3>summary>b{padding-left:46px}.team>b{padding-left:64px}.grid>b{min-width:0;overflow-wrap:anywhere;word-break:break-word}.level-jd{background:#f5f8fc}.level-dd{background:#fafbfd}.level-official{background:#fff}.team{background:#fbfcfe;border-bottom:1px dashed #dfe7ee}.level-jd>span:first-child{color:#6740aa;font-weight:800}.level-dd>span:first-child{color:#116c5c;font-weight:800}.level-official>span:first-child{color:#2454d9;font-weight:800}.team>span:first-child{color:var(--muted);font-weight:700}
@media(max-width:900px){.status-kpis{grid-template-columns:repeat(2,1fr)}}@media(max-width:760px){.wrap{padding:12px}.head{display:block}.nav{margin-top:12px}.status-kpis{grid-template-columns:1fr}.controls{justify-content:flex-start;flex-wrap:wrap}.foot{display:block}}
@media print{body{background:#fff}.top{padding:12px;color:#000;background:#fff;border-bottom:2px solid #000}.top small{color:#555}.wrap{padding:10px;max-width:none}.panel{box-shadow:none;border:0;padding:0}.nav,.controls{display:none}.scroll{overflow:visible;border:0}.grid{width:100%;min-width:100%;max-width:100%;grid-template-columns:65px minmax(170px,1.7fr) minmax(62px,.58fr) repeat(3,minmax(55px,.55fr)) repeat(5,minmax(65px,.65fr));font-size:9px}.node,.node>summary,.children,.children .grid{width:100%;min-width:100%;max-width:100%}.node{break-inside:avoid}.header{background:#ddd!important;color:#000}.header>span:first-child,.header>b{background:#ddd!important}.grid>span:first-child,.grid>b{position:static}}
</style>
</head>
<body>
<header class="top"><small>CITES OPERATIONS INTELLIGENCE · STANDALONE REPORT</small><h1>Resolved + closed during the week</h1><small>Generated __GENERATED_AT__</small></header>
<main class="wrap"><section class="panel">
  <div class="head"><div><h2 id="week-title"></h2><div class="muted">EPFO-queue issue transitions shown through the JD → DD → official → team hierarchy. Routed cases appear separately as CDAC + ROs.</div></div><nav class="nav"><button class="arrow" id="previous" title="Previous week">←</button><button id="latest">Latest week</button><button class="arrow" id="next" title="Next week">→</button></nav></div>
  <div class="status-kpis"><article class="kpi-total"><span>EPFO issues</span><strong id="status-total">—</strong><small id="status-as-of">Snapshot unavailable</small></article><article class="kpi-resolved"><span>EPFO resolved + closed</span><strong id="status-resolved">—</strong><small>Resolved, fixed or closed</small></article><article class="kpi-yesterday"><span>EPFO resolved + closed (yesterday)</span><strong id="status-yesterday">—</strong><small id="status-comparison">Comparison unavailable</small></article><article class="kpi-open"><span>EPFO open</span><strong id="status-open">—</strong><small>All non-resolved statuses</small></article></div>
  <div class="controls"><select id="depth" title="Expansion depth"><option value="1">JD totals only</option><option value="2">Through DD</option><option value="3">Through officials</option><option value="4" selected>Show teams</option></select><button id="expand">Expand all</button><button id="collapse">Collapse all</button></div>
  <div class="scroll" id="calendar"></div>
  <div class="foot"><span>* Workload and activity cells contain EPFO-queue cases only. The CDAC + ROs column shows routed cases attributed to each EPFO responsibility.</span><span id="coverage"></span></div>
</section></main>
<script>
const WEEKS = __WEEKS_JSON__;
const INITIAL = __INITIAL_WEEK__;
let current = Math.max(0, WEEKS.findIndex(w => w.week_start === INITIAL));
if (current < 0) current = WEEKS.length - 1;

const fmt = n => new Intl.NumberFormat('en-IN').format(n || 0);
const dateLabel = s => {
  if (!s) return '—';
  const parts = s.split('-');
  if (parts.length === 3) {
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  return s;
};

function cell(text, tag = 'span') {
  const e = document.createElement(tag);
  e.textContent = text;
  return e;
}

function dayValue(week, days, date) {
  const dObj = week.dates.find(d => d.date === date);
  return dObj && dObj.available ? fmt(days[date] || 0) : '—';
}

function row(level, name, metrics, days, week, extra = '') {
  const e = document.createElement('div');
  e.className = 'grid ' + extra;
  e.append(cell(level));
  e.append(cell(name, 'b'));
  e.append(cell(fmt(metrics?.cdac_total || 0) + '+' + fmt(metrics?.field_total || 0)));
  
  [['total', 'status-total-cell'], ['open', 'status-open-cell'], ['resolved', 'status-resolved-cell']].forEach(([key, className]) => {
    const c = cell(week.snapshot_date ? fmt(metrics?.[key] || 0) : '—');
    c.className = className;
    c.title = week.snapshot_date ? 'Current status as of ' + week.snapshot_date : 'Complete issue snapshot unavailable';
    e.append(c);
  });

  week.dates.forEach(d => {
    const value = days[d.date] || 0;
    const c = cell(dayValue(week, days, d.date));
    if (d.available) {
      c.title = 'Changes since ' + (d.from_date || 'prior') + (value < 10 ? ' · Below 10' : '');
      if (value < 10) c.classList.add('daily-low');
    } else {
      c.title = 'Complete comparison snapshot unavailable';
    }
    e.append(c);
  });
  return e;
}

function branch(node, week, depth = 1) {
  if (depth === 4) return row(node.level, node.name, node, node.days, week, 'team');
  const details = document.createElement('details');
  details.className = 'node depth-' + depth;
  const summary = document.createElement('summary');
  const rowElement = row(node.level, node.name, node, node.days, week, depth === 1 ? 'level-jd' : (depth === 2 ? 'level-dd' : 'level-official'));
  summary.className = rowElement.className;
  summary.append(...Array.from(rowElement.childNodes));
  details.append(summary);
  const children = document.createElement('div');
  children.className = 'children';
  (node.children || []).forEach(child => children.append(branch(child, week, depth + 1)));
  details.append(children);
  return details;
}

function setDepth(value) {
  document.querySelectorAll('#calendar details').forEach(d => {
    const depth = Number(d.className.match(/depth-(\\d)/)?.[1] || 9);
    d.open = depth < Number(value);
  });
}

function render() {
  const w = WEEKS[current];
  const calendar = document.getElementById('calendar');
  calendar.replaceChildren();
  document.getElementById('week-title').textContent = dateLabel(w.week_start) + ' – ' + dateLabel(w.week_end);
  document.getElementById('previous').disabled = current === 0;
  document.getElementById('next').disabled = current === WEEKS.length - 1;

  const status = w.status || {};
  document.getElementById('status-total').textContent = status.total === null ? '—' : fmt(status.total);
  document.getElementById('status-resolved').textContent = status.resolved_closed === null ? '—' : fmt(status.resolved_closed);
  document.getElementById('status-yesterday').textContent = status.resolved_since_previous === null ? '—' : fmt(status.resolved_since_previous);
  document.getElementById('status-open').textContent = status.open === null ? '—' : fmt(status.open);
  document.getElementById('status-as-of').textContent = status.as_of ? 'As of ' + dateLabel(status.as_of) : 'Snapshot unavailable';
  document.getElementById('status-comparison').textContent = status.previous_date && status.as_of ? dateLabel(status.previous_date) + ' → ' + dateLabel(status.as_of) : 'Comparison unavailable';

  const available = w.dates.filter(d => d.available).length;
  document.getElementById('coverage').textContent = available + ' of 5 weekdays have comparable snapshots.';

  const header = document.createElement('div');
  header.className = 'grid header';
  header.append(cell('Level'));
  header.append(cell('Name of official / team', 'b'));
  header.append(cell('CDAC + ROs'));
  [['Total issues', 'status-total-cell'], ['Open', 'status-open-cell'], ['Resolved*', 'status-resolved-cell']].forEach(([label, className]) => {
    const c = cell(label);
    c.className = className;
    header.append(c);
  });
  w.dates.forEach(d => {
    const c = cell(d.day);
    c.append(cell(d.label, 'small'));
    header.append(c);
  });
  calendar.append(header);

  calendar.append(row('All', 'EPFO tracker queues', w.epfo_status || {}, w.totals, w, 'total'));

  const tree = document.createElement('div');
  if (w.hierarchy.length) {
    w.hierarchy.forEach(n => tree.append(branch(n, w, 1)));
  } else {
    tree.append(Object.assign(document.createElement('div'), { className: 'empty', textContent: 'No EPFO ownership rows are available for this week.' }));
  }
  calendar.append(tree);
  setDepth(document.getElementById('depth').value);
}

document.getElementById('previous').onclick = () => { if (current > 0) { current--; render(); } };
document.getElementById('next').onclick = () => { if (current < WEEKS.length - 1) { current++; render(); } };
document.getElementById('latest').onclick = () => { current = WEEKS.length - 1; render(); };
document.getElementById('depth').onchange = e => setDepth(e.target.value);
document.getElementById('expand').onclick = () => document.querySelectorAll('#calendar details').forEach(d => d.open = true);
document.getElementById('collapse').onclick = () => document.querySelectorAll('#calendar details').forEach(d => d.open = false);

render();
</script>
</body>
</html>
"""

    @classmethod
    def generate_html(
        cls,
        df_current: pd.DataFrame,
        output_path: Union[str, Path],
        df_teams: pd.DataFrame,
        stats_sources: Optional[List[Dict[str, Any]]] = None,
        historical_snapshots: Optional[Dict[str, pd.DataFrame]] = None,
        report_date: Optional[str] = None,
        week_date: Optional[str] = None,
    ) -> Path:
        """Builds and writes the complete weekly resolutions HTML application."""
        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        if report_date is None:
            report_date = str(date.today())

        # Determine Monday of requested week
        base_dt = datetime.strptime(report_date, "%Y-%m-%d").date() if report_date else date.today()
        # Monday of current week
        current_monday = base_dt - timedelta(days=base_dt.weekday())
        # Previous Monday
        prev_monday = current_monday - timedelta(days=7)

        # Build week objects
        weeks = []
        for mon in [prev_monday, current_monday]:
            week_obj = cls._build_week_data(
                mon,
                df_current,
                df_teams,
                stats_sources=stats_sources,
                historical_snapshots=historical_snapshots,
                report_date=report_date,
            )
            weeks.append(week_obj)

        selected_week = week_date or (prev_monday.isoformat() if base_dt.weekday() == 0 else current_monday.isoformat())

        html_content = cls.HTML_TEMPLATE
        html_content = html_content.replace("__GENERATED_AT__", datetime.now().strftime("%d %b %Y, %I:%M %p"))
        html_content = html_content.replace("__WEEKS_JSON__", json.dumps(weeks, ensure_ascii=False))
        html_content = html_content.replace("__INITIAL_WEEK__", json.dumps(selected_week))

        out_path.write_text(html_content, encoding="utf-8")
        return out_path

    @classmethod
    def _build_week_data(
        cls,
        monday: date,
        df_current: pd.DataFrame,
        df_teams: pd.DataFrame,
        stats_sources: Optional[List[Dict[str, Any]]] = None,
        historical_snapshots: Optional[Dict[str, pd.DataFrame]] = None,
        report_date: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Builds one week dataset with 5 weekdays, totals, and recursive hierarchy."""
        fri = monday + timedelta(days=4)
        dates_list = []
        dates_keys = []
        for i in range(5):
            d = monday + timedelta(days=i)
            d_str = d.isoformat()
            dates_keys.append(d_str)
            dates_list.append({
                "date": d_str,
                "day": d.strftime("%A"),
                "label": d.strftime("%d %b"),
                "available": False,
                "from_date": None,
            })

        # Calculate daily resolved counts for each weekday
        daily_totals = {k: 0 for k in dates_keys}
        category_daily = {}  # category -> {date -> count}

        # Derive weekday values only from comparable cumulative snapshots.
        # No estimates or date-specific constants are permitted in an official report.
        ordered_sources = sorted(
            (src for src in (stats_sources or []) if src.get("source", {}).get("data_date")),
            key=lambda src: src["source"]["data_date"],
        )
        for index, current in enumerate(ordered_sources):
            current_date = current["source"]["data_date"]
            if current_date not in daily_totals or index == 0:
                continue
            previous = ordered_sources[index - 1]
            previous_date = previous["source"]["data_date"]
            previous_categories = {
                cat.get("module_key", ""): int(cat.get("resolved", 0)) + int(cat.get("closed", 0))
                for cat in previous.get("categories", [])
            }
            current_categories = {
                cat.get("module_key", ""): int(cat.get("resolved", 0)) + int(cat.get("closed", 0))
                for cat in current.get("categories", [])
            }
            all_categories = set(previous_categories) | set(current_categories)
            for category_key in all_categories:
                if not category_key:
                    continue
                category_daily.setdefault(category_key, {dk: 0 for dk in dates_keys})
                delta = current_categories.get(category_key, 0) - previous_categories.get(category_key, 0)
                category_daily[category_key][current_date] = delta
                daily_totals[current_date] += delta
            for d_entry in dates_list:
                if d_entry["date"] == current_date:
                    d_entry["available"] = True
                    d_entry["from_date"] = previous_date

        # Calculate module level snapshots from df_current
        def get_queue_type(assignee: str) -> str:
            val = str(assignee or "").lower().strip()
            if val.startswith("team_epfo_"):
                return "epfo"
            elif val.startswith("team_cdac_"):
                return "cdac"
            elif val.startswith("ro.") or val.startswith("sro."):
                return "field"
            return "other"

        df_work = df_current.copy()
        df_work["q_type"] = df_work["Assigned To"].apply(get_queue_type)
        closed_mask = df_work["Status"].astype(str).str.lower().isin(["resolved", "fixed", "closed"])

        # Aggregate current status
        epfo_df = df_work[df_work["q_type"] == "epfo"]
        epfo_status = {
            "total": len(epfo_df),
            "open": int((~closed_mask[epfo_df.index]).sum()),
            "resolved": int(closed_mask[epfo_df.index].sum()),
            "cdac_total": int((df_work["q_type"] == "cdac").sum()),
            "field_total": int((df_work["q_type"] == "field").sum()),
        }

        current_source = next(
            (src for src in reversed(ordered_sources) if src["source"]["data_date"] == report_date),
            None,
        )
        previous_source = None
        if current_source:
            current_index = ordered_sources.index(current_source)
            if current_index > 0:
                previous_source = ordered_sources[current_index - 1]
        resolved_delta = None
        previous_date = None
        if current_source and previous_source:
            current_total = int(current_source.get("totals", {}).get("resolved", 0)) + int(current_source.get("totals", {}).get("closed", 0))
            previous_total = int(previous_source.get("totals", {}).get("resolved", 0)) + int(previous_source.get("totals", {}).get("closed", 0))
            resolved_delta = current_total - previous_total
            previous_date = previous_source["source"]["data_date"]

        status_kpi = {
            "total": epfo_status["total"],
            "open": epfo_status["open"],
            "resolved_closed": epfo_status["resolved"],
            "resolved_since_previous": resolved_delta,
            "as_of": report_date,
            "previous_date": previous_date,
        }

        # Build 4-tier hierarchy for weekly table: JD -> DD -> Official -> Team
        hierarchy = cls._build_weekly_hierarchy(df_work, df_teams, dates_keys, category_daily)

        return {
            "week_start": monday.isoformat(),
            "week_end": fri.isoformat(),
            "dates": dates_list,
            "totals": daily_totals,
            "snapshot_date": report_date,
            "status": status_kpi,
            "epfo_status": epfo_status,
            "hierarchy": hierarchy,
        }

    @classmethod
    def _build_weekly_hierarchy(
        cls,
        df_issues: pd.DataFrame,
        df_teams: pd.DataFrame,
        dates_keys: List[str],
        category_daily: Dict[str, Dict[str, int]],
    ) -> List[Dict[str, Any]]:
        mapper = WorkforceMapper()
        norm_map = {}
        if not df_teams.empty:
            team_cat_col = mapper.hierarchy_cfg.get("team_column", "Team")
            for _, r in df_teams.iterrows():
                raw_teams = str(r.get(team_cat_col, "")).split(",")
                for t in raw_teams:
                    norm_t = WorkforceMapper._normalize_key(t)
                    if norm_t:
                        norm_map[norm_t] = {
                            "team_name": t.strip(),
                            "handler": WorkforceMapper._clean_person_name(r.get("Account handled by", "Primary handler not mapped")),
                            "dd": WorkforceMapper._clean_person_name(r.get("DD(IS)", "Not specified")),
                            "jd": WorkforceMapper._clean_person_name(r.get("JD(IS)", "Not specified")),
                        }

        # Map each issue to team
        team_issues = {}
        for _, row in df_issues.iterrows():
            cat = str(row.get("Category", "Unassigned")).strip()
            q_type = row.get("q_type", "other")
            is_closed = str(row.get("Status", "")).lower() in ["resolved", "fixed", "closed"]

            if cat not in team_issues:
                team_issues[cat] = {"total": 0, "open": 0, "resolved": 0, "cdac_total": 0, "field_total": 0}
            if q_type == "epfo":
                team_issues[cat]["total"] += 1
                if is_closed:
                    team_issues[cat]["resolved"] += 1
                else:
                    team_issues[cat]["open"] += 1
            elif q_type == "cdac":
                team_issues[cat]["cdac_total"] += 1
            elif q_type == "field":
                team_issues[cat]["field_total"] += 1

        # Build tree: JD -> DD -> Handler -> Team
        jd_tree: Dict[str, Dict[str, Any]] = {}

        for cat_raw, metrics in team_issues.items():
            norm_key = WorkforceMapper._normalize_key(cat_raw)
            meta = norm_map.get(norm_key, {
                "team_name": cat_raw,
                "handler": "Primary handler not mapped",
                "dd": "Not specified",
                "jd": "Not specified",
            })

            jd_name = meta["jd"]
            dd_name = meta["dd"]
            h_name = meta["handler"]
            t_name = meta["team_name"]

            # Daily distribution for this team
            team_days = dict(category_daily.get(norm_key, {k: 0 for k in dates_keys}))

            if jd_name not in jd_tree:
                jd_tree[jd_name] = {
                    "level": "JD",
                    "name": jd_name,
                    "total": 0, "open": 0, "resolved": 0, "cdac_total": 0, "field_total": 0,
                    "days": {k: 0 for k in dates_keys},
                    "dds": {},
                }
            jd_node = jd_tree[jd_name]
            jd_node["total"] += metrics["total"]
            jd_node["open"] += metrics["open"]
            jd_node["resolved"] += metrics["resolved"]
            jd_node["cdac_total"] += metrics["cdac_total"]
            jd_node["field_total"] += metrics["field_total"]
            for k in dates_keys:
                jd_node["days"][k] += team_days[k]

            if dd_name not in jd_node["dds"]:
                jd_node["dds"][dd_name] = {
                    "level": "DD",
                    "name": dd_name,
                    "total": 0, "open": 0, "resolved": 0, "cdac_total": 0, "field_total": 0,
                    "days": {k: 0 for k in dates_keys},
                    "handlers": {},
                }
            dd_node = jd_node["dds"][dd_name]
            dd_node["total"] += metrics["total"]
            dd_node["open"] += metrics["open"]
            dd_node["resolved"] += metrics["resolved"]
            dd_node["cdac_total"] += metrics["cdac_total"]
            dd_node["field_total"] += metrics["field_total"]
            for k in dates_keys:
                dd_node["days"][k] += team_days[k]

            if h_name not in dd_node["handlers"]:
                dd_node["handlers"][h_name] = {
                    "level": "Official",
                    "name": h_name,
                    "total": 0, "open": 0, "resolved": 0, "cdac_total": 0, "field_total": 0,
                    "days": {k: 0 for k in dates_keys},
                    "teams": [],
                }
            h_node = dd_node["handlers"][h_name]
            h_node["total"] += metrics["total"]
            h_node["open"] += metrics["open"]
            h_node["resolved"] += metrics["resolved"]
            h_node["cdac_total"] += metrics["cdac_total"]
            h_node["field_total"] += metrics["field_total"]
            for k in dates_keys:
                h_node["days"][k] += team_days[k]

            h_node["teams"].append({
                "level": "Team",
                "name": t_name,
                "total": metrics["total"],
                "open": metrics["open"],
                "resolved": metrics["resolved"],
                "cdac_total": metrics["cdac_total"],
                "field_total": metrics["field_total"],
                "days": team_days,
            })

        # Format final tree structure
        result = []
        for jd_name, jd_data in sorted(jd_tree.items(), key=lambda x: (x[0] == "Not specified", -x[1]["open"])):
            dds_list = []
            for dd_name, dd_data in sorted(jd_data["dds"].items(), key=lambda x: -x[1]["open"]):
                h_list = []
                for h_name, h_data in sorted(dd_data["handlers"].items(), key=lambda x: -x[1]["open"]):
                    h_list.append({
                        "level": "Official",
                        "name": h_name,
                        "total": h_data["total"],
                        "open": h_data["open"],
                        "resolved": h_data["resolved"],
                        "cdac_total": h_data["cdac_total"],
                        "field_total": h_data["field_total"],
                        "days": h_data["days"],
                        "children": h_data["teams"],
                    })
                dds_list.append({
                    "level": "DD",
                    "name": dd_name,
                    "total": dd_data["total"],
                    "open": dd_data["open"],
                    "resolved": dd_data["resolved"],
                    "cdac_total": dd_data["cdac_total"],
                    "field_total": dd_data["field_total"],
                    "days": dd_data["days"],
                    "children": h_list,
                })
            result.append({
                "level": "JD",
                "name": jd_name,
                "total": jd_data["total"],
                "open": jd_data["open"],
                "resolved": jd_data["resolved"],
                "cdac_total": jd_data["cdac_total"],
                "field_total": jd_data["field_total"],
                "days": jd_data["days"],
                "children": dds_list,
            })

        return result
