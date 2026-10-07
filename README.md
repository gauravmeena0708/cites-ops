# CITES Operations Intelligence (`cites-ops`)

## Preferred workflow: a simplified local briefing

Use `brief` to turn a daily input folder into **one dashboard, one Excel workbook and two Top 15 PowerPoint decks**.
It runs entirely on local Python/JavaScript code and YAML rules, without an LLM, external model,
or network request. Reports use **Functionality → Issue Nature → Ticket**, with
Open, Resolved and Closed kept separate.

```powershell
cites-ops brief "C:\path\to\new_ingest\2026-09-07" --scope open
# For an export containing every ticket, including resolved and closed:
cites-ops brief "C:\path\to\new_ingest\2026-09-07" --scope all
```

Inputs: an issue CSV, `teams.csv` (or `issue_teams.csv`), and an optional matching
Samadhan Setu status DOCX. By default the output goes to:

```text
reports/briefings/2026-09-07/
  dashboard_2026-09-07.html
  issues_2026-09-07.xlsx
  top_15_open_2026-09-07.pptx
  top_15_resolved_closed_2026-09-07.pptx
  _internal/                  # Counts/history, classification cache and run log
```

The workbook has **Functionality Summary**, **Issue Natures**, **Issue Details**,
**Top Open Natures**, and **Top Resolved Closed**. The two decks use the CITES
reference styling and show counts plus up to three distinct examples per nature.
Use `--top-n 10` for Top 10 or `--ranking-period week` for last-week submissions.
The dashboard uses the CITES navy palette and expandable functionality/nature or
JD/DD/officer/functionality/nature views, with previous-week intake and searchable
ticket details. Major/minor labels and technical rule keys
are omitted from these reports.

PPT rendering needs Node.js and a **locally installed `@oai/artifact-tool` runtime**.
This workspace is configured at `.cache/presentations`. Elsewhere, pass
`--pptx-runtime PATH` (or set `CITES_ARTIFACT_WORKSPACE`) to the local Node workspace
containing that package. Use `--no-pptx` to run with Python only. See the guide below
for setup. No runtime model account or API key is used.

Optional WhatsApp exports generate a searchable `knowledge_base.html` of **review
candidates** for new functionality, reported fixes and workarounds:

```powershell
cites-ops brief "C:\path\to\new_ingest\2026-09-07" --scope open `
  --chats "C:\path\to\new_ingest\2026-08-31" --chat-since 2026-08-01
```

See [the simplified reporting guide](docs/SIMPLIFIED_REPORTING.md) for data coverage,
weekly definitions, rule customization, output controls and limitations. Existing
commands below remain available for older reporting workflows.

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**CITES Operations Intelligence** is an enterprise-grade issue triage, workforce accountability, and support chat analysis framework. It ingests multi-channel operational data (issue tracker CSVs, support chat dumps, and organizational hierarchy mappings), classifies tickets deterministically, extracts sanitized troubleshooting knowledge, and generates professional management deliverables (Excel workbooks, PowerPoint slide decks, Word administrative notes, and offline HTML dashboards).

---

## 🌟 Key Features

* **Deterministic Issue Classification**: Versioned rule-based text classifier (`rules.yaml`) with regex prioritization, workflow-level extraction (DA, SS, AO/APFC, RPFC), and SHA-256 text checksums for 100% reproducible categorization.
* **Support Chat Knowledge Mining & PII Masking**: Ingests WhatsApp exports (`.txt` or `.zip`), automatically identifies technical resolutions, masks sensitive identifiers (UANs, mobile numbers, member accounts), and formats shareable **Field Office Knowledge Notes (`.docx`)**.
* **Zero-Hardcoding Entity Cross-Referencing**: Built-in inverted index matches tickets with chat discussions by extracting common entity identifiers (UAN, Member ID, Grievance ID, Task ID) without modifying Python code.
* **Dynamic N-Tier Workforce Hierarchy**: Maps flat ticket queues into multi-level management structures (`Level 1 → Level 2 → Level 3 → Handler`) using custom `teams.csv` mappings.
* **Multi-Format Enterprise Deliverables**:
  * 📊 **Excel (`.xlsx`)**: Formatted multi-tab workbooks with KPI summaries, category pivots, and aging indicators (`openpyxl`).
  * 📽️ **PowerPoint (`.pptx`)**: Executive 16:9 slide decks for management reviews (`python-pptx`).
  * 📝 **Word (`.docx`)**: Formal administrative action notes and sanitized knowledge documents (`python-docx`).
  * 🌐 **Offline HTML Dashboard (`.html`)**: Self-contained, responsive morning brief dashboard.

---

## 🚀 Installation (Miniconda / Python 3.9+)

### 1. Create a Conda Environment
```bash
conda create -n cites python=3.11 -y
conda activate cites
```

### 2. Install the Package
```bash
# Option A: From local clone
git clone https://github.com/your-org/cites-ops.git
cd cites-ops
pip install -e .

# Option B: Direct install via Git
pip install git+https://github.com/your-org/cites-ops.git
```

---

## 💻 Command Line Interface (CLI)

### 1. Classify an Issue Tracker CSV
Categorizes raw issue exports into Major/Minor categories and generates an Excel summary:
```bash
cites-ops classify "issues/tracker.csv" --excel "reports/Categorized.xlsx"
```

### 2. Extract Technical Knowledge from Support Chats
Extracts problem–solution discussions, redacts PII, cross-references with tickets, and produces a Word Knowledge Note:
```bash
cites-ops chat-kb "chat_dumps/chat.txt" --issues "issues/tracker.csv" --output "Knowledge_Note.docx"
```

### 3. Generate Full Executive Daily Reporting Pack
Generates the complete suite (Excel + PowerPoint + HTML Dashboard + Word Note):
```bash
cites-ops report "issues/tracker.csv" --teams "config/teams.csv" --date 2026-08-14 --out-dir "reports/2026-08-14"
```

### 4. Map Workforce Workload & Accountability
Analyzes open backlog distribution across organizational management levels:
```bash
cites-ops workforce "issues/tracker.csv" "config/teams.csv"
```

### 5. Fetch Live Issues from MantisBT & Generate Daily Status
Connects to MantisBT REST API, ingests live issues, and generates daily issues CSV/Excel alongside the official **Samadhan Setu Status DOCX**:
```bash
cites-ops fetch-mantis --token "<YOUR_MANTIS_API_TOKEN>" --out-dir "new_ingest/2026-08-28"
```

### 6. Generate the Standardized End-to-End Daily Pack

The legacy `daily` workflow fetches MantisBT once and publishes its larger report pack under an ISO-date directory. For the simplified offline pack, use `brief` above:

```bash
cites-ops daily --date 2026-08-28 --input-dir "../tmp/input" --output-root "../tmp/output"
```

Expected inputs are `issue_teams.csv` and, when `MANTIS_API_TOKEN` is not set, an optional `token.txt`. The standard output is:

```text
../tmp/output/2026-08-28/
  cites_status_2026-08-28.docx
  cites_issues_2026-08-28.xlsx
  cites_issues_2026-08-28.csv
  cites_dashboard_2026-08-28.html
  cites_issue_topics_2026-08-28.html
  cites_defect_drilldown_2026-08-28.html
  cites_weekly_resolutions_2026-08-28.html
  manifest.json
  run.log
```

Use `--reports status,issues_xlsx,dashboard` to create a subset, `--exclude weekly_resolutions` to omit reports, and `--overwrite` to explicitly replace an existing date directory. TLS certificate verification is enabled by default; `--insecure` is intended only for local development.

---

## ⚙️ Configuration (Zero Hardcoding)

All categorization rules, entity patterns, and routing queue definitions live in declarative YAML files under `cites_ops/config/`:

| Config File | Purpose | Customization Example |
| :--- | :--- | :--- |
| **`rules.yaml`** | Issue classification rules & priority weights | Add custom regexes for new modules or defect types. |
| **`entities.yaml`** | Patterns for entity matching & PII redaction | Define UAN, Grievance ID, Member ID, PAN, or custom ticket numbers. |
| **`routing.yaml`** | Queue routing classifications | Define patterns for internal tech, vendors, and regional offices. |
| **`default_config.yaml`** | Aging thresholds & hierarchy column headers | Configure aging days (7d, 30d) and CSV column names. |

---

## 🐍 Python Library Usage

You can also import `cites_ops` directly into your Python scripts or Jupyter Notebooks:

```python
import pandas as pd
from cites_ops import IssueClassifier, EntityMatcher, ChatParser, ExcelReporter

# 1. Classify Tickets
df = pd.read_csv("issues.csv")
classifier = IssueClassifier()
df_classified = classifier.classify_dataframe(df)

# 2. Extract Entities & PII Masking
matcher = EntityMatcher()
masked_text = matcher.mask_pii("UAN 100123456789 reported issue in member MH/BAN/12345/678")

# 3. Export Formatted Excel
ExcelReporter.generate_report(df_classified, "Report.xlsx")
```

---

## 🧪 Running Tests

The test suite runs with standard Python `unittest`:

```bash
python -m unittest discover tests
```

---

## 📂 Repository Structure

```
cites-ops/
├── pyproject.toml               # Modern Python packaging configuration
├── setup.py                     # Setuptools configuration & console scripts
├── requirements.txt             # Dependency requirements
├── README.md                    # Documentation & user guide
├── .gitignore                   # Excludes sensitive data, dumps, and output files
├── cites_ops/
│   ├── __init__.py              # Package export definitions
│   ├── cli.py                   # Main CLI entry point ('cites-ops')
│   ├── config/                  # Declarative zero-hardcoding configurations
│   │   ├── rules.yaml           # Deterministic issue classification rules
│   │   ├── entities.yaml        # Entity regexes & PII masking templates
│   │   ├── routing.yaml         # Routing queue patterns (EPFO, CDAC, RO)
│   │   └── default_config.yaml  # Default settings & hierarchy column mappings
│   ├── core/                    # Core business logic & algorithms
│   │   ├── classifier.py        # Rule-based deterministic classification engine
│   │   ├── entity_matcher.py    # Entity extractor & inverted cross-ref index
│   │   ├── chat_parser.py       # WhatsApp chat parser & knowledge miner
│   │   ├── workforce.py         # N-tier dynamic hierarchy mapper
│   │   └── ingest.py            # CSV validator & delta/aging calculator
│   ├── reporters/               # Multi-format report builders
│   │   ├── excel_reporter.py    # Multi-tab formatted Excel workbooks (.xlsx)
│   │   ├── pptx_reporter.py     # Executive presentation slides (.pptx)
│   │   ├── docx_reporter.py     # Administrative notes (.docx)
│   │   └── html_reporter.py     # Offline responsive HTML dashboards (.html)
│   └── utils/                   # Helpers and loaders
│       ├── config_loader.py     # YAML loader & validator
│       └── helpers.py           # Text normalizers, date parsers, checksums
├── tests/                       # Unit test suite
└── examples/                    # Clean synthetic sample data for onboarding
```

---

## 📄 License

This project is licensed under the MIT License.
