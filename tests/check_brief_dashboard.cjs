// Dependency-free DOM contract checks for the actual embedded report scripts.
// These complement Python data tests; they do not replace visual browser QA.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

class Element {
  constructor(tag) { this.tagName = tag; this.children = []; this.value = ''; this.checked = false; this.hidden = false; this._text = ''; }
  set textContent(value) { this._text = String(value); this.children = []; }
  get textContent() { return this._text + this.children.map(c => c.textContent).join(''); }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this._text = ''; this.children = children; }
  scrollIntoView() {}
  showModal() { this.open = true; }
  close() { this.open = false; }
}

function load(name, payload, dataId) {
  const html = fs.readFileSync(path.join(__dirname, '../cites_ops/templates', name), 'utf8');
  const elements = new Map([...html.matchAll(/\bid="([^"]+)"/g)].map(m => [m[1], new Element('div')]));
  elements.get(dataId).textContent = JSON.stringify(payload);
  if (elements.has('period-filter')) elements.get('period-filter').value = 'all';
  if (elements.has('top-only')) elements.get('top-only').checked = true;
  const context = {
    document: { getElementById: id => { assert.ok(elements.has(id), `Missing element: ${id}`); return elements.get(id); }, createElement: tag => new Element(tag) },
    setTimeout: callback => { callback(); return 1; }, clearTimeout() {}, console,
  };
  vm.createContext(context);
  for (const match of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/g)) {
    if (!match[1].includes('application/json')) vm.runInContext(match[2], context, { timeout: 5000 });
  }
  return id => elements.get(id);
}

const issues = Array.from({ length: 8 }, (_, i) => ({
  id: String(i + 1), functionality: 'Form-13', nature: i === 7 ? 'Other / needs review' : 'Nature ' + i,
  status: i === 0 ? 'closed' : i === 1 ? 'resolved' : 'open', summary: 'Summary ' + i,
  description: 'Description ' + i, officer: 'Officer A', dd: 'Deputy A', jd: 'Joint A',
  assigned_to: 'Queue A', submitted: '2026-09-01', updated: '2026-09-07', age_days: 6,
  last_week: i < 4, prior_week: i >= 4,
}));
const $ = load('briefing.html', {
  date: '2026-09-07', count_source: 'test.csv', totals: { total: 8, open: 6, resolved: 1, closed: 1 },
  warnings: ['Scope warning'], scope: 'all', issues, review_count: 1, comparison: null,
  week: { start: '2026-08-31', end: '2026-09-06', prior_start: '2026-08-24', prior_end: '2026-08-30' },
  functions: [{ functionality: 'Form-13', total: 8, open: 6, resolved: 1, closed: 1, detail_total: 8,
    officer: 'Officer A', dd: 'Deputy A', jd: 'Joint A' }],
}, 'report-data');
assert.equal($('kpis').children.length, 4);
assert.equal($('nature-rows').children.length, 7); // defined natures ranked globally; review separate
assert.equal($('ticket-rows').children.length, 8);
const topBranch = $('tree').children[1];
assert.equal(topBranch.open, false);
$('expand-all').onclick(); assert.equal(topBranch.open, true);
$('collapse-all').onclick(); assert.equal(topBranch.open, false);
$('tree-view').value = 'ownership'; $('tree-view').onchange();
assert.match($('tree').textContent, /Joint A/);
$('status-filter').value = 'resolved_closed'; $('status-filter').onchange();
assert.equal($('ticket-rows').children.length, 2);
$('status-filter').value = ''; $('status-filter').onchange();
$('period-filter').value = 'week'; $('period-filter').onchange();
assert.equal($('ticket-rows').children.length, 4);
assert.match($('nature-title').textContent, /last week/);
$('status-filter').value = 'closed'; $('status-filter').onchange();
assert.equal($('ticket-rows').children.length, 1);
$('ticket-rows').children[0].children[0].children[0].onclick();
assert.equal($('ticket-dialog').open, true);
assert.equal($('dialog-description').textContent, 'Description 0');
$('close-dialog').onclick(); assert.equal($('ticket-dialog').open, false);
$('reset').onclick();
$('top-only').checked = false; $('top-only').onchange();
assert.equal($('nature-rows').children.length, 7);
$('nature-rows').children[0].children[1].children[0].onclick();
assert.equal($('ticket-rows').children.length, 1);
assert.equal($('clear-nature').hidden, false);
$('clear-nature').onclick(); assert.equal($('ticket-rows').children.length, 8);
$('search').value = 'Description 2'; $('search').oninput();
assert.equal($('ticket-rows').children.length, 1);
$('search').value = 'no matching ticket'; $('search').oninput();
assert.equal($('ticket-rows').children.length, 0);
assert.equal($('next').disabled, true);

const kb = load('knowledge.html', { date: '2026-09-07', entries: [
  { date: '2026-09-01', type: 'Reported fix', state: 'Candidate — needs review', text: 'Patch deployed',
    functionalities: ['Form-13'], tickets: ['1'], evidence: [{ source: 'chat.txt', message: 2, date: '2026-09-01', text: 'Patch deployed' }] },
  { date: '2026-09-02', type: 'Workaround', state: 'Candidate — needs review', text: 'Follow these steps',
    functionalities: ['Form-19'], tickets: [], evidence: [] },
] }, 'data');
assert.equal(kb('entries').children.length, 2);
kb('kind').value = 'Workaround'; kb('kind').onchange();
assert.equal(kb('entries').children.length, 1);
kb('search').value = 'Form-13'; kb('search').oninput();
assert.equal(kb('entries').children.length, 0);
console.log('Dashboard and knowledge scripts passed render/filter/drilldown/dialog/search checks.');
