# Simplified functionality reporting

`cites-ops brief INPUT_FOLDER` is the preferred folder-based workflow. It does not
call MantisBT, an LLM, or any remote service. Existing `daily`, `report`, `classify`,
`workforce`, and chat/report-specific commands retain their previous interfaces.

## Inputs and truthful coverage

Supply one issue CSV, a teams mapping and, optionally, the status Word document
for the same snapshot date. The ISO date is read from the input folder name or
provided with `--date YYYY-MM-DD`. Keep the input exports as the source archive.

Required issue fields: `Id`, `Category`, `Status`, `Summary`, `Description`,
`Date Submitted`. `Updated` (or `Last Update`), `Assigned To`, and `Project` are
used when available. The original tracker `Category` is the functionality.
The mapping requires `Team`, `Account handled by`, `DD(IS)`, and `JD(IS)`;
comma-separated functionalities in `Team` are supported. Conflicting mappings
are rejected. Unmapped officers remain visible.

| Scope | Meaning | Portfolio counts |
| --- | --- | --- |
| `--scope all` | You confirm the CSV contains all issues in the intended portfolio. | Computed from the CSV; reconciled with the status document when supplied. |
| `--scope open` | Complete export of currently open issues. | The status document supplies total/resolved/closed. Without it these remain unavailable; the open count comes from the CSV. |
| `--scope partial` | Any filtered export. | Available only from a matching status document. Textual analysis stays labelled as partial. |
| `--scope auto` | Default; recognizes a CSV containing only open states. | Same as open. Mixed statuses require explicit `all` or `partial`; the tool cannot infer export completeness. |

`new`, `assigned`, `feedback`, `open`, `acknowledged`, `confirmed`, and `reopened`
count as Open. `resolved` and `fixed` count as Resolved. `closed` counts as Closed.
Unknown statuses fail validation instead of silently changing totals.

When supplied, the status document must match the report date and reconcile
internally and with the CSV's declared coverage, by functionality and overall.
An open CSV never acquires invented details for resolved/closed cases. The
dashboard labels the scope of counts separately from the scope of ticket analysis.
Use a full export to analyze natures across all statuses.

Exact duplicate rows are removed; conflicting duplicate IDs fail. Blank IDs fail.
Missing/unreadable submission dates produce a warning and are excluded from week
counts; ages remain unavailable. Tickets submitted or updated after the report
date are rejected because a current export cannot reconstruct a past snapshot.

## Four routine outputs

- `dashboard_DATE.html`: CITES navy palette, coloured status cards, expandable
  Functionality → Nature or JD → DD → Officer → Functionality → Nature hierarchy;
  expand/collapse all and depth selection; top five natures per functionality;
  overall Top 10/15; four status counts and weekly intake;
  search, status/period/ownership filters and paginated ticket details. Full
  description is displayed on click. Everything needed is embedded for offline use.
- `issues_DATE.xlsx`: Functionality Summary, Issue Natures, Issue Details,
  Top Open Natures and Top Resolved Closed. The two ranked sheets contain counts,
  affected functionalities and up to three distinct, masked examples per nature.
  Summary counts use the same prepared data as the dashboard. All natures appear
  in Excel so counts reconcile; flags indicate the top five. Date and numeric
  cells retain their types. DD/JD, original status and description columns are
  hidden initially in Issue Details and can be expanded in Excel.
- `top_15_open_DATE.pptx`: title/scope, ranked table and one evidence slide per
  nature, with its count and up to three representative ticket examples.
- `top_15_resolved_closed_DATE.pptx`: the same structure for resolved + closed,
  retaining separate resolved and closed counts in the ranked table.

`--top-n 10` changes both ranking sheets and decks to Top 10. The default ranks
all supplied tickets. `--ranking-period week` instead selects last-calendar-week
submissions; the HTML opens with the same period selected. These filters never
turn update dates into completion dates. The ranking uses distinct tickets,
deterministic label ordering to break ties, and excludes Other / needs review,
whose count remains visible. Fewer than 15 defined natures produce fewer slides.
Examples prioritize rules matched in the summary, then recent tickets, while
preferring different functionalities/statuses and distinct summaries; fewer
than three examples are shown when there are fewer distinct descriptions.
Configured identifiers, email addresses and explicitly labelled credentials are
masked in these excerpts. Names and other free text can remain; full source text
is retained in the ticket-detail workbook and dashboard.

On an open-only export, the resolved/closed deck is a one-slide availability
notice and its Excel sheet has no invented ranks. A partial export is labelled
as such throughout. Zero supplied closed tickets does not mean no closed tickets
exist in the portfolio.

The source status DOCX is retained as the formal input; `brief` does not regenerate
it. Separate defect/topic/weekly dashboards, raw CSV copies and the
generic administrative note are not generated by this workflow.

`_internal` holds a manifest with counts, input/output hashes, run timing, scope and
rules fingerprint; a classification cache; a log; and sanitized knowledge entries.
This is support data, not a management report. Keep dated output folders for
overall snapshot comparisons and continued cache/knowledge reuse. No remote
storage is involved.

### Local PowerPoint runtime

The HTML/Excel/classification pipeline uses Python. PPT rendering uses Node.js
with `@oai/artifact-tool` installed locally; it does not call an external model.
The renderer edits cloned slides from the sanitized CITES template included in
the package. It preserves source geometry, typography, logo and inherited layouts.
The template contains placeholders rather than source ticket data.

In Codex, initialize a workspace with the installed Presentations skill helper:

```powershell
node "<Presentations skill folder>/container_tools/setup_artifact_tool_workspace.mjs" --workspace .cache/presentations
```

The current repository has this local runtime configured. For another local
installation, supply its Node workspace with `--pptx-runtime PATH` or
`CITES_ARTIFACT_WORKSPACE`. Installing this Python package alone does not install
the Node rendering library. Missing runtime or failed rendering fails the run
and preserves earlier reports; `--no-pptx` explicitly selects HTML/Excel only.
No automatic downloads or API credentials are used by `brief`.

Set `CITES_PPTX_PREVIEW=1` when reviewing template changes to render every slide
and save layout JSON under `_internal`; it is unnecessary for routine generation.

## Issue nature and the previous week

Each ticket has **one primary issue nature** inferred from Summary and Description.
Rules use a priority and prefer matches in the summary. Unmatched tickets remain
`Other / needs review`. The labels describe reported symptoms, not verified causes.

The default view shows up to five defined natures **for each functionality**, ranked
by open tickets and then total tickets. The review bucket stays visible separately.
There is no requirement to invent two natures for a functionality with only one.
Turn off the Top 5 checkbox to display every nature.

“Last week” means the last completed Monday–Sunday period. For 2026-09-07 this is
2026-08-31 through 2026-09-06 inclusive; the preceding comparison week is
2026-08-24 through 2026-08-30. The date boundaries are shown. The last-week view
ranks natures by submissions, not current open backlog. Status columns describe
the tickets' **current snapshot status**, not their status at the end of that week.

On an open-only export, week counts cover only tickets from those weeks that are
still open. They are not complete intake totals. Resolution-event counts require
ticket status history and are not inferred from the update date. Overall changes
between saved complete portfolios are labelled net snapshot changes, with the
actual prior date shown; a gap is not silently labelled “yesterday.”

## Customizing the local rules

The default is `cites_ops/config/rules.yaml`; existing major-family metadata is
ignored by the simplified outputs. Use `--rules PATH` for a replacement catalogue.
`nature_label`, when supplied, is the standalone symptom label used by `brief`;
legacy commands retain their original `label`. This lets labels such as
“At DA level” become “Claim/task is not visible or routed at DA level” in briefs.
Explicit missing paths fail. See `examples/nature_rules.yaml` for a minimal,
synthetic example. A rule supports:

```yaml
rule_id: TRANSFER_VISIBILITY
label: Transfer claim not visible at receiving office
priority: 100
functionalities: [Form-13]       # Optional; absent applies across functionalities
patterns:
  - '(?i)claim.{0,60}not visible.{0,60}receiving office'
exclude_patterns:
  - '(?i)claim is now visible'   # Optional; suppress known false matches
```

Changing rule content invalidates cached classifications even if the version label
is unchanged. Cache keys also include functionality and normalized ticket text.
Only hashes, labels and rule IDs are cached, not raw ticket text. Repeated or
unchanged text reuses prior results. The default rules are deterministic heuristics,
not an accuracy guarantee. Review a representative labelled sample across
functionalities and the review bucket before relying on classification decisions.

## Optional chat knowledge

`--chats` accepts TXT/ZIP files or directories containing those exports. All text
transcripts in a ZIP are parsed without extracting its media. Default dates are
day/month/year; select `--chat-date-order mdy` for month/day/year exports. Use
`--chat-since DATE` to choose the first message date; the default is the start of
last week. Messages after the report date are excluded.

`knowledge_base.html` contains explicit announcements selected by
`cites_ops/config/knowledge.yaml`:

- New functionality announcements with explicit development/deployment wording.
- Reported fixes, excluding negative statements, promises and common questions.
- Explicit workarounds or instructions.

Vague replies such as “done” do not become entries. Nearby messages within one
hour can supply context only when they share a functionality or a linked issue;
context does not establish a confirmed resolution. Source export/transcript,
message number and date are retained. Ticket references and configured entities
can link an announcement to supplied tickets. Unknown functionality remains
`Needs identification`. All entries are review candidates, not generated claims
that a popular issue is fixed everywhere. Summaries use message text, not invented
instructions. Media-only announcements, indirect phrasing, multilingual text and
ambiguous discussions may be missed.

Configured identifiers and email addresses are masked in candidate text and
supporting messages. Automated masking is not a substitute for review before
sharing. Original senders are not displayed. Entries are deduplicated and retained
from the latest briefing in the same output root, so that root should belong to a
single reporting portfolio. This is a searchable candidate collection; approval
or editorial management is not implemented. Do not treat it as a verified KB.

## Commands

```powershell
# Current open-only intake, reconciled with the matching Word status document:
cites-ops brief "C:\Users\IT\Downloads\CITES\new_ingest\2026-09-07" --scope open

# Full issue export with optional August chat announcements:
cites-ops brief "C:\path\to\2026-09-07" --scope all `
  --chats "C:\path\to\2026-08-31" --chat-since 2026-08-01

# Disambiguate files and override output location:
cites-ops brief "C:\path\to\input" --date 2026-09-07 --scope all `
  --issues "C:\path\to\all.csv" --teams "C:\path\to\teams.csv" `
  --stats "C:\path\to\status.docx" --output-root reports/briefings

# Explicitly ignore Word statistics and use complete CSV counts:
cites-ops brief "C:\path\to\2026-09-07" --scope all --no-stats
```

An existing date folder requires `--overwrite`. Only a previously generated
simplified briefing can be replaced; legacy packs and unrelated folders are
refused. Reports are built in staging before replacement. Inputs are never edited.

## Verification

```text
python -m unittest discover tests
node tests/check_brief_dashboard.cjs
```

Tests cover status/coverage reconciliation, week boundaries, conflicting inputs,
classification scope/exclusions/cache invalidation, workbook data types, HTML
payload escaping, offline execution, replacement protection and chat extraction.
The default classifier catalogue still needs domain review against real tickets.
