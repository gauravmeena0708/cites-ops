"""Reporter for generating interactive issue topics and ownership hierarchy HTML applications."""

from __future__ import annotations

import json
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Union, Dict, Any, List
import pandas as pd

from ..core.classifier import IssueClassifier
from ..core.workforce import WorkforceMapper


class InteractiveTopicsReporter:
    """Generates standalone, self-contained interactive CITES Issue Topics & Ownership HTML dashboards."""

    HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<title>CITES Issue Topics and Ownership Hierarchy — Interactive Snapshot</title>
<style>
:root{--navy:#0b1f33;--ink:#172a3a;--muted:#627382;--paper:#f3f6fa;--white:#fff;--line:#d7e0e8;--blue:#326bff;--teal:#148b7b;--green:#137d68;--amber:#bd7400;--red:#a61b2b;--purple:#6740aa;--shadow:0 10px 34px rgba(11,31,51,.1)}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.45 "Segoe UI",Aptos,Arial,sans-serif}.top{background:linear-gradient(120deg,var(--navy),#183b5a);color:#fff;padding:25px max(18px,calc((100vw - 1480px)/2))}.top h1{font-size:27px;margin:4px 0}.eyebrow{color:#a9c4d8;font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase}.top-meta{display:flex;gap:14px;flex-wrap:wrap;color:#d5e2ec;font-size:12px}.wrap{max-width:1480px;margin:auto;padding:22px;overflow-x:hidden}.notice{background:#fff4d7;border:1px solid #edcf81;border-left:4px solid var(--amber);border-radius:9px;padding:11px 14px;margin-bottom:16px;color:#654408}.kpis{display:grid;grid-template-columns:repeat(5,minmax(150px,1fr));gap:11px;margin-bottom:16px}.kpi{background:#fff;border:1px solid var(--line);border-radius:10px;padding:14px 16px;box-shadow:0 4px 16px rgba(11,31,51,.05)}.kpi span{display:block;color:var(--muted);font-size:11px;font-weight:700}.kpi strong{display:block;font-size:25px;margin-top:4px}.kpi.accent{background:var(--teal);color:#fff;border-color:var(--teal)}.kpi.accent span{color:#dcf8f2}.panel{background:#fff;border:1px solid var(--line);border-radius:11px;box-shadow:var(--shadow);padding:17px;min-width:0}.panel-head{display:flex;justify-content:space-between;align-items:flex-end;gap:15px;margin-bottom:13px}.panel-head h2{font-size:19px;margin:0 0 3px}.panel-head p{margin:0;color:var(--muted);font-size:12px}.controls{display:flex;gap:8px;flex-wrap:wrap}.controls input,.controls select,.modal-tools input,.modal-tools select{border:1px solid #bdcbd6;background:#fff;border-radius:8px;padding:9px 10px;color:var(--ink);font:inherit}.controls input{min-width:275px}.table-scroll{overflow:auto;max-width:100%;max-height:70vh;border:1px solid var(--line);border-radius:8px}table{width:100%;border-collapse:collapse;min-width:1110px}th,td{padding:10px 11px;border-bottom:1px solid #e5ebf0;text-align:left;vertical-align:top;overflow-wrap:anywhere}th{position:sticky;top:0;z-index:3;background:var(--navy);color:#fff;font-size:10px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}th.num,td.num{text-align:right}.topic-row{cursor:pointer}.topic-row:hover,.topic-row:focus{background:#edf3ff;outline:none}.topic-name{font-weight:800;color:#214dbf}.topic-name small{display:block;color:var(--muted);font-size:10px;font-weight:500;margin-top:3px}.description{color:var(--muted);min-width:310px}.overall-row{position:sticky;top:37px;z-index:2;background:#e7edf4;font-weight:800}.route{display:inline-block;border-radius:15px;padding:3px 7px;margin:1px;font-size:10px;font-style:normal;font-weight:800;white-space:nowrap}.route-epfo{background:#e6eeff;color:#2454d9}.route-cdac{background:#dcf4ed;color:#116c5c}.route-field{background:#fff0cf;color:#895900}.route-other{background:#e9eef3;color:#50616e}.up{color:var(--red);font-weight:800}.down{color:var(--green);font-weight:800}.empty{text-align:center;color:var(--muted);padding:30px}.footer{color:var(--muted);font-size:11px;text-align:center;padding:18px}.modal{position:fixed;inset:0;z-index:30;display:none;background:rgba(6,19,31,.67);padding:18px}.modal.open{display:flex}.modal-card{margin:auto;background:#fff;width:min(1420px,100%);height:min(92vh,940px);border-radius:12px;box-shadow:0 24px 70px rgba(0,0,0,.3);display:grid;grid-template-rows:auto auto minmax(0,1fr) auto;overflow:hidden}.modal-head{background:var(--navy);color:#fff;padding:16px 18px;display:flex;justify-content:space-between;gap:14px}.modal-head h2{margin:0;font-size:21px;overflow-wrap:anywhere}.modal-head p{margin:3px 0 0;color:#bcd0df;font-size:12px}.close{border:0;background:#fff;color:var(--navy);border-radius:8px;font-size:23px;width:38px;height:38px;cursor:pointer}.modal-tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:11px 16px;border-bottom:1px solid var(--line);background:#f8fafc}.modal-tools input{min-width:300px;flex:1}.result-count{margin-left:auto;color:var(--muted);font-size:12px}.issue-scroll{overflow:auto;min-width:0}.issue-table{min-width:1180px}.issue-table th{top:0}.issue-table td{font-size:12px}.issue-id{font-weight:800;color:#214dbf;white-space:nowrap}.issue-summary{font-weight:700;max-width:460px}.issue-summary details{margin-top:5px;color:var(--muted);font-weight:400;max-width:100%}.issue-summary summary{cursor:pointer;color:#516d85;font-size:11px}.issue-summary p{white-space:pre-wrap;margin:6px 0 0;max-height:180px;max-width:100%;overflow:auto;overflow-wrap:anywhere}.status{display:inline-block;border-radius:13px;padding:3px 7px;font-size:10px;font-weight:800;background:#e9eef3}.status-open{background:#fff0cf;color:#895900}.status-closed{background:#dcf4ed;color:#116c5c}.queue-code{display:block;margin-top:4px;font:10px Consolas,monospace;overflow-wrap:anywhere}.modal-foot{display:flex;align-items:center;justify-content:center;gap:10px;padding:10px;border-top:1px solid var(--line);background:#f8fafc}.button{border:0;border-radius:8px;background:var(--blue);color:#fff;padding:8px 13px;font:inherit;font-weight:700;cursor:pointer}.button:disabled{opacity:.35;cursor:not-allowed}.page-label{color:var(--muted);font-size:12px;min-width:150px;text-align:center}
@media(max-width:900px){.kpis{grid-template-columns:repeat(2,1fr)}.panel-head{display:block}.controls{margin-top:10px}.controls input{min-width:0;flex:1}.wrap{padding:12px}.modal{padding:0}.modal-card{height:100vh;border-radius:0}.modal-tools input{min-width:0}.result-count{width:100%;margin:0}.top{padding:18px}.top h1{font-size:23px}.tree summary,.tree-topic{grid-template-columns:1fr}.tree-metrics{white-space:normal;flex-wrap:wrap}.tree-children{margin-left:10px}.tree-level{display:block;min-width:0;margin-bottom:2px}}
@media print{body{background:#fff}.top{background:#fff;color:#000;border-bottom:2px solid #000}.top-meta,.eyebrow{color:#444}.notice,.controls,.footer,.modal{display:none!important}.wrap{max-width:none;padding:8px}.panel{box-shadow:none;border:0;padding:0}.table-scroll{max-height:none;overflow:visible;border:0}table{min-width:100%;font-size:9px}th{position:static;background:#ddd;color:#000}.overall-row{position:static}.topic-row{break-inside:avoid}}
.major-row{background:#edf2f7;font-weight:800;border-top:2px solid #bdcbd6}.minor-row td:first-child{padding-left:34px}.minor-row td:first-child:before{content:'↳';color:var(--muted);margin-right:7px}
.hierarchy-panel{margin-top:18px}.hierarchy-tools{display:flex;gap:8px;flex-wrap:wrap}.tree{border:1px solid var(--line);border-radius:9px;overflow:hidden;max-width:100%;min-width:0}.tree details{border-bottom:1px solid var(--line);max-width:100%;min-width:0}.tree details:last-child{border-bottom:0}.tree summary{cursor:pointer;display:grid;grid-template-columns:minmax(0,1fr) minmax(0,auto);gap:14px;align-items:center;padding:11px 13px;background:#fff;max-width:100%;min-width:0}.tree summary:hover{background:#f4f8ff}.tree-children{margin-left:0;border-left:0;max-width:100%;min-width:0}.tree-children summary{padding-left:30px}.tree-children .tree-children summary{padding-left:47px}.tree-children .tree-children .tree-children summary{padding-left:64px}.tree-level{display:inline-block;min-width:108px;color:var(--muted);font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}.tree-name{font-weight:750;overflow-wrap:anywhere;word-break:break-word;min-width:0}.tree-metrics{display:flex;gap:12px;align-items:center;white-space:normal;flex-wrap:wrap;justify-content:flex-end;font-size:11px;color:var(--muted);min-width:0}.tree-metrics b{color:var(--ink);font-size:12px}.tree-topic{appearance:none;border:0;border-bottom:1px solid var(--line);background:#fbfdff;color:var(--ink);cursor:pointer;width:100%;max-width:100%;min-width:0;padding:10px 13px 10px 64px;display:grid;grid-template-columns:minmax(0,1fr) minmax(0,auto);gap:14px;text-align:left;align-items:center}.tree-topic:last-child{border-bottom:0}.tree-topic:hover,.tree-topic:focus{background:#edf3ff}.tree-topic-name{font-weight:800;color:#214dbf;overflow-wrap:anywhere;word-break:break-word;min-width:0}.tree-topic small{display:block;color:var(--muted);font-weight:400;margin-top:2px}.tree-empty{padding:22px;color:var(--muted);text-align:center}
</style>
</head>
<body>
<header class="top">
  <div class="eyebrow">CITES Operational Intelligence · Standalone Snapshot</div>
  <h1>Issue Topics and Ownership Hierarchy</h1>
  <div class="top-meta">
    <span>Data snapshot date: <strong>__DATA_THROUGH_DATE__</strong></span>
    <span>Generated: <strong>__GENERATED_AT__</strong></span>
    <span>Classification catalog: <strong>v2.01-hierarchical</strong></span>
  </div>
</header>
<main class="wrap">
  <div class="notice">
    <strong>Decision-Support Notice:</strong> Every issue is classified into a 2-tier problem taxonomy (Major Topic → Minor Problem Topic) and mapped to functional leadership (JD → DD → Officer → Team). Click any topic or team row to view the full drilldown table.
  </div>
  <section class="kpis">
    <div class="kpi"><span>Total Tracker Issues</span><strong id="kpi-total">0</strong></div>
    <div class="kpi"><span>Open Backlog</span><strong id="kpi-open">0</strong></div>
    <div class="kpi accent"><span>Resolved / Closed</span><strong id="kpi-resolved">0</strong></div>
    <div class="kpi"><span>Fresh Influx (7 Days)</span><strong id="kpi-new">0</strong></div>
    <div class="kpi"><span>Net Backlog Delta</span><strong id="kpi-delta">0</strong></div>
  </section>

  <section class="panel">
    <div class="panel-head">
      <div>
        <h2>Problem Topics Register</h2>
        <p>Sorted by total open backlog. Click any row to inspect all matching issues.</p>
      </div>
      <div class="controls">
        <input type="search" id="topic-filter" placeholder="Filter topics by name, description, rule...">
        <select id="route-filter">
          <option value="all">All Routing Queues</option>
          <option value="epfo">EPFO Core Queues</option>
          <option value="cdac">CDAC Vendor Queues</option>
          <option value="field">Regional Offices (ROs)</option>
        </select>
      </div>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th>Problem Topic</th>
            <th>Rule ID & Workflow Level</th>
            <th>Summary / Description</th>
            <th class="num">Total</th>
            <th class="num">Open</th>
            <th class="num">Resolved</th>
            <th>Routing Queues Breakdown</th>
          </tr>
        </thead>
        <tbody id="topics-body"></tbody>
      </table>
    </div>
  </section>

  <section class="panel hierarchy-panel">
    <div class="panel-head">
      <div>
        <h2>5-Tier Organizational & Topic Hierarchy</h2>
        <p>Expandable hierarchy from Joint Director (IS) down to Specific Problem Topics.</p>
      </div>
      <div class="hierarchy-tools">
        <button class="button" id="btn-expand-all">Expand All</button>
        <button class="button" id="btn-collapse-all" style="background:#516d85">Collapse All</button>
      </div>
    </div>
    <div class="tree" id="hierarchy-tree"></div>
  </section>
</main>

<div class="modal" id="issue-modal">
  <div class="modal-card">
    <div class="modal-head">
      <div>
        <h2 id="modal-title">Topic Issues Drilldown</h2>
        <p id="modal-subtitle">Showing matching issues</p>
      </div>
      <button class="close" id="modal-close" title="Close modal">&times;</button>
    </div>
    <div class="modal-tools">
      <input type="search" id="modal-search" placeholder="Search issues by ID, summary, description, reporter, assignee...">
      <select id="modal-status">
        <option value="all">All Statuses</option>
        <option value="open">Open Issues Only</option>
        <option value="resolved">Resolved / Closed Only</option>
      </select>
      <span class="result-count" id="modal-count">0 matching issues</span>
    </div>
    <div class="issue-scroll">
      <table class="issue-table">
        <thead>
          <tr>
            <th>ID</th>
            <th>Summary & Description</th>
            <th>Category / Module</th>
            <th>Assigned To</th>
            <th>Status</th>
            <th>Submitted</th>
            <th>Updated</th>
            <th>Workflow Rule</th>
          </tr>
        </thead>
        <tbody id="modal-body"></tbody>
      </table>
    </div>
    <div class="modal-foot">
      <button class="button" id="btn-prev-page">&larr; Prev</button>
      <span class="page-label" id="page-label">Page 1 of 1</span>
      <button class="button" id="btn-next-page">Next &rarr;</button>
    </div>
  </div>
</div>

<footer class="footer">
  CITES Operations Intelligence Framework · Generated on __GENERATED_AT__ · All Data Offline & Standalone
</footer>

<script>
const DATA = __PAYLOAD_JSON__;

const fmt = n => new Intl.NumberFormat('en-IN').format(n || 0);

// Populate KPIs
document.getElementById('kpi-total').textContent = fmt(DATA.overall.total);
document.getElementById('kpi-open').textContent = fmt(DATA.overall.open);
document.getElementById('kpi-resolved').textContent = fmt(DATA.overall.resolved);
document.getElementById('kpi-new').textContent = fmt(DATA.overall.new);
document.getElementById('kpi-delta').textContent = (DATA.overall.delta > 0 ? '+' : '') + fmt(DATA.overall.delta);

// Render Topics Table
function renderTopics() {
  const query = document.getElementById('topic-filter').value.toLowerCase().trim();
  const route = document.getElementById('route-filter').value;
  const tbody = document.getElementById('topics-body');
  tbody.replaceChildren();

  DATA.majors.forEach(major => {
    const majorMatching = DATA.topics.filter(t => t.major_key === major.key);
    const visibleTopics = majorMatching.filter(t => {
      const matchText = !query || t.name.toLowerCase().includes(query) || (t.description||'').toLowerCase().includes(query) || (t.rule_id||'').toLowerCase().includes(query);
      const matchRoute = route === 'all' || (t.routes && t.routes[route] > 0);
      return matchText && matchRoute;
    });

    if (visibleTopics.length === 0 && query) return;

    // Major row
    const mtr = document.createElement('tr');
    mtr.className = 'major-row topic-row';
    mtr.innerHTML = `
      <td>${major.name}</td>
      <td><strong>Major Topic</strong></td>
      <td class="description">${major.description || ''}</td>
      <td class="num">${fmt(major.total)}</td>
      <td class="num">${fmt(major.open)}</td>
      <td class="num">${fmt(major.resolved)}</td>
      <td>${renderRoutes(major.routes)}</td>
    `;
    mtr.onclick = () => openModalForMajor(major);
    tbody.appendChild(mtr);

    visibleTopics.forEach(t => {
      const tr = document.createElement('tr');
      tr.className = 'minor-row topic-row';
      tr.innerHTML = `
        <td class="topic-name">${t.name}<small>${t.workflow_level || ''}</small></td>
        <td><code>${t.rule_id || 'C00'}</code></td>
        <td class="description">${t.description || ''}</td>
        <td class="num">${fmt(t.total)}</td>
        <td class="num"><strong>${fmt(t.open)}</strong></td>
        <td class="num">${fmt(t.resolved)}</td>
        <td>${renderRoutes(t.routes)}</td>
      `;
      tr.onclick = () => openModalForTopic(t);
      tbody.appendChild(tr);
    });
  });
}

function renderRoutes(routes) {
  if (!routes) return '';
  let html = '';
  if (routes.epfo) html += `<span class="route route-epfo">EPFO: ${routes.epfo}</span> `;
  if (routes.cdac) html += `<span class="route route-cdac">CDAC: ${routes.cdac}</span> `;
  if (routes.field) html += `<span class="route route-field">ROs: ${routes.field}</span> `;
  if (routes.other) html += `<span class="route route-other">Other: ${routes.other}</span> `;
  return html;
}

// Render Hierarchy Tree
function renderHierarchy() {
  const container = document.getElementById('hierarchy-tree');
  container.replaceChildren();

  function buildBranch(node) {
    if (node.kind === 'topic') {
      const btn = document.createElement('button');
      btn.className = 'tree-topic';
      btn.innerHTML = `
        <span class="tree-topic-name">${node.name} <small>${node.rule_id} · ${node.workflow_level || ''}</small></span>
        <span class="tree-metrics"><span>Total: <b>${fmt(node.total)}</b></span><span>Open: <b>${fmt(node.open)}</b></span><span>Resolved: <b>${fmt(node.resolved)}</b></span></span>
      `;
      btn.onclick = () => openModalForTopic({ key: node.topic_key, name: node.name, rule_id: node.rule_id });
      return btn;
    }

    const details = document.createElement('details');
    details.open = node.level === 'JD' || node.level === 'Joint Director (IS)';
    const summary = document.createElement('summary');
    summary.innerHTML = `
      <span><span class="tree-level">${node.level}:</span> <span class="tree-name">${node.name}</span></span>
      <span class="tree-metrics"><span>Total: <b>${fmt(node.total)}</b></span><span>Open: <b>${fmt(node.open)}</b></span><span>Resolved: <b>${fmt(node.resolved)}</b></span></span>
    `;
    details.appendChild(summary);

    const childrenContainer = document.createElement('div');
    childrenContainer.className = 'tree-children';
    (node.children || []).forEach(child => childrenContainer.appendChild(buildBranch(child)));
    details.appendChild(childrenContainer);
    return details;
  }

  DATA.hierarchy.forEach(topNode => container.appendChild(buildBranch(topNode)));
}

// Modal State & Pagination
let currentFilteredIssues = [];
let currentPage = 1;
const PAGE_SIZE = 50;

function openModalForTopic(topic) {
  document.getElementById('modal-title').textContent = topic.name;
  document.getElementById('modal-subtitle').textContent = `Rule: ${topic.rule_id || ''} · ${topic.description || ''}`;
  currentFilteredIssues = DATA.issues.filter(i => i.topic_key === topic.key || i.rule_id === topic.rule_id);
  currentPage = 1;
  document.getElementById('modal-search').value = '';
  document.getElementById('modal-status').value = 'all';
  updateModalTable();
  document.getElementById('issue-modal').classList.add('open');
}

function openModalForMajor(major) {
  document.getElementById('modal-title').textContent = major.name;
  document.getElementById('modal-subtitle').textContent = `Major Defect Group · ${major.description || ''}`;
  currentFilteredIssues = DATA.issues.filter(i => i.major_topic_key === major.key);
  currentPage = 1;
  document.getElementById('modal-search').value = '';
  document.getElementById('modal-status').value = 'all';
  updateModalTable();
  document.getElementById('issue-modal').classList.add('open');
}

function updateModalTable() {
  const query = document.getElementById('modal-search').value.toLowerCase().trim();
  const statusFilter = document.getElementById('modal-status').value;

  const matched = currentFilteredIssues.filter(issue => {
    const matchStatus = statusFilter === 'all' ? true : (statusFilter === 'resolved' ? issue.is_closed : !issue.is_closed);
    if (!matchStatus) return false;
    if (!query) return true;
    return (
      (issue.tracker_id || '').toLowerCase().includes(query) ||
      (issue.summary || '').toLowerCase().includes(query) ||
      (issue.description || '').toLowerCase().includes(query) ||
      (issue.assigned_to || '').toLowerCase().includes(query) ||
      (issue.tracker_category || '').toLowerCase().includes(query)
    );
  });

  document.getElementById('modal-count').textContent = `${fmt(matched.length)} matching issues`;

  const totalPages = Math.max(1, Math.ceil(matched.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;

  document.getElementById('page-label').textContent = `Page ${currentPage} of ${totalPages}`;
  document.getElementById('btn-prev-page').disabled = currentPage === 1;
  document.getElementById('btn-next-page').disabled = currentPage === totalPages;

  const start = (currentPage - 1) * PAGE_SIZE;
  const pageItems = matched.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('modal-body');
  tbody.replaceChildren();

  pageItems.forEach(issue => {
    const tr = document.createElement('tr');
    const statusClass = issue.is_closed ? 'status-closed' : 'status-open';
    tr.innerHTML = `
      <td class="issue-id">#${issue.tracker_id}</td>
      <td class="issue-summary">
        <div>${issue.summary}</div>
        ${issue.description ? `<details><summary>View Full Description</summary><p>${escapeHtml(issue.description)}</p></details>` : ''}
      </td>
      <td><strong>${issue.tracker_category || ''}</strong></td>
      <td>${issue.assigned_to || 'Unassigned'}<span class="queue-code">${issue.queue_group || ''}</span></td>
      <td><span class="status ${statusClass}">${issue.status}</span></td>
      <td>${issue.date_submitted || ''}</td>
      <td>${issue.updated_date || ''}</td>
      <td><code>${issue.rule_id || 'C00'}</code></td>
    `;
    tbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Event Listeners
document.getElementById('topic-filter').oninput = renderTopics;
document.getElementById('route-filter').onchange = renderTopics;
document.getElementById('modal-search').oninput = () => { currentPage = 1; updateModalTable(); };
document.getElementById('modal-status').onchange = () => { currentPage = 1; updateModalTable(); };
document.getElementById('modal-close').onclick = () => document.getElementById('issue-modal').classList.remove('open');
document.getElementById('btn-prev-page').onclick = () => { if (currentPage > 1) { currentPage--; updateModalTable(); } };
document.getElementById('btn-next-page').onclick = () => { currentPage++; updateModalTable(); };
document.getElementById('btn-expand-all').onclick = () => document.querySelectorAll('#hierarchy-tree details').forEach(d => d.open = true);
document.getElementById('btn-collapse-all').onclick = () => document.querySelectorAll('#hierarchy-tree details').forEach(d => d.open = false);

// Init
renderTopics();
renderHierarchy();
</script>
</body>
</html>
"""

    @classmethod
    def generate_html(
        cls,
        df_enriched: pd.DataFrame,
        output_path: Union[str, Path],
        df_teams: Optional[pd.DataFrame] = None,
        report_date: Optional[str] = None,
        title: str = "CITES Issue Topics and Ownership Hierarchy",
    ) -> Path:
        """Builds and writes the complete standalone interactive topics HTML application."""
        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        if report_date is None:
            report_date = str(date.today())

        # Ensure classification columns exist
        if "rule_id" not in df_enriched.columns:
            classifier = IssueClassifier()
            df_enriched = classifier.classify_dataframe(df_enriched)

        total_issues = len(df_enriched)
        open_mask = ~df_enriched["Status"].astype(str).str.lower().isin(["resolved", "fixed", "closed"])
        open_issues = int(open_mask.sum())
        resolved_issues = total_issues - open_issues
        fresh_7d = int((df_enriched["age_days"] <= 7).sum()) if "age_days" in df_enriched.columns else 0

        # Build queue group column
        def get_queue_group(assignee: str) -> str:
            val = str(assignee or "").lower().strip()
            if val.startswith("team_epfo_"):
                return "epfo"
            elif val.startswith("team_cdac_"):
                return "cdac"
            elif val.startswith("ro.") or val.startswith("sro."):
                return "field"
            return "other"

        df_enriched["queue_group"] = df_enriched["Assigned To"].apply(get_queue_group)
        df_enriched["is_closed"] = ~open_mask

        # Build Major Topics
        majors_map: Dict[str, Dict[str, Any]] = {}
        topics_map: Dict[str, Dict[str, Any]] = {}

        for _, row in df_enriched.iterrows():
            m_key = str(row.get("major_topic_label", "General")).lower().replace(" ", "_")
            m_name = str(row.get("major_topic_label", "General Defects"))
            m_desc = str(row.get("category_description", ""))
            
            if m_key not in majors_map:
                majors_map[m_key] = {
                    "key": m_key,
                    "name": m_name,
                    "description": m_desc,
                    "total": 0,
                    "open": 0,
                    "resolved": 0,
                    "routes": {"epfo": 0, "cdac": 0, "field": 0, "other": 0},
                    "children": set(),
                }
            majors_map[m_key]["total"] += 1
            if row["is_closed"]:
                majors_map[m_key]["resolved"] += 1
            else:
                majors_map[m_key]["open"] += 1
            majors_map[m_key]["routes"][row["queue_group"]] += 1

            # Minor Topic
            t_key = str(row.get("topic_key", row.get("rule_id", "C00"))).lower()
            t_name = str(row.get("topic_label", row.get("category_label", "General")))
            t_rule = str(row.get("rule_id", "C00"))
            t_wf = str(row.get("workflow_level", "System"))
            t_desc = str(row.get("category_description", ""))

            majors_map[m_key]["children"].add(t_key)

            if t_key not in topics_map:
                topics_map[t_key] = {
                    "key": t_key,
                    "major_key": m_key,
                    "name": t_name,
                    "description": t_desc,
                    "workflow_level": t_wf,
                    "rule_id": t_rule,
                    "total": 0,
                    "open": 0,
                    "resolved": 0,
                    "routes": {"epfo": 0, "cdac": 0, "field": 0, "other": 0},
                }
            topics_map[t_key]["total"] += 1
            if row["is_closed"]:
                topics_map[t_key]["resolved"] += 1
            else:
                topics_map[t_key]["open"] += 1
            topics_map[t_key]["routes"][row["queue_group"]] += 1

        majors_list = []
        for m in majors_map.values():
            m["children"] = sorted(list(m["children"]))
            majors_list.append(m)
        majors_list.sort(key=lambda x: x["open"], reverse=True)

        topics_list = list(topics_map.values())
        topics_list.sort(key=lambda x: (x["open"], x["total"]), reverse=True)

        # Build 5-tier hierarchy
        mapper = WorkforceMapper()
        if df_teams is not None:
            workload = mapper.process_workload(df_enriched, df_teams)
            hierarchy = workload.get("hierarchy_tree", [])
        else:
            hierarchy = []

        # Build JSON issue list (lightweight dictionary)
        issues_payload = []
        for _, row in df_enriched.iterrows():
            issues_payload.append({
                "tracker_id": str(row.get("Id", "")),
                "summary": str(row.get("Summary", "")),
                "description": str(row.get("Description", "")) if pd.notna(row.get("Description")) else "",
                "status": str(row.get("Status", "assigned")),
                "assigned_to": str(row.get("Assigned To", "Unassigned")),
                "tracker_category": str(row.get("Category", "Unassigned")),
                "functional_team": str(row.get("Category", "")),
                "date_submitted": str(row.get("Date Submitted", "")),
                "updated_date": str(row.get("Updated", "")),
                "major_topic_key": str(row.get("major_topic_label", "General")).lower().replace(" ", "_"),
                "major_topic_label": str(row.get("major_topic_label", "General")),
                "topic_key": str(row.get("topic_key", row.get("rule_id", "C00"))).lower(),
                "topic_label": str(row.get("topic_label", "")),
                "workflow_level_label": str(row.get("workflow_level", "")),
                "rule_id": str(row.get("rule_id", "C00")),
                "queue_group": row["queue_group"],
                "is_closed": bool(row["is_closed"]),
                "is_new": bool(row.get("age_days", 999) <= 7),
            })

        payload = {
            "meta": {
                "title": title,
                "data_through_date": report_date,
                "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
                "classifier": "v2.01-hierarchical",
            },
            "overall": {
                "total": total_issues,
                "open": open_issues,
                "resolved": resolved_issues,
                "new": fresh_7d,
                "delta": open_issues - resolved_issues,
            },
            "majors": majors_list,
            "topics": topics_list,
            "hierarchy": hierarchy,
            "issues": issues_payload,
        }

        json_str = json.dumps(payload, ensure_ascii=False)

        html_content = cls.HTML_TEMPLATE
        html_content = html_content.replace("__DATA_THROUGH_DATE__", report_date)
        html_content = html_content.replace("__GENERATED_AT__", datetime.now().strftime("%d %b %Y, %I:%M %p"))
        html_content = html_content.replace("__PAYLOAD_JSON__", json_str)

        out_path.write_text(html_content, encoding="utf-8")
        return out_path
