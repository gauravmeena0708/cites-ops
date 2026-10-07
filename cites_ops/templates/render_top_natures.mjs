// Local rendering only: this module makes no model or network calls.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath, pathToFileURL} from 'node:url';

const [runtime, input, output] = process.argv.slice(2);
if (!runtime || !input || !output) throw new Error('Expected runtime workspace, ranking JSON, output directory');
const requireRuntime = createRequire(path.join(path.resolve(runtime), 'package.json'));
const {FileBlob, PresentationFile} = await import(pathToFileURL(requireRuntime.resolve('@oai/artifact-tool')).href);
const base = path.dirname(fileURLToPath(import.meta.url));
const data = JSON.parse(await fs.readFile(input, 'utf8'));
const mapping = JSON.parse(await fs.readFile(path.join(base, 'top_natures.map.json'), 'utf8'));
const short = (value, limit) => value.length <= limit ? value : value.slice(0, limit - 1).replace(/\s+\S*$/, '') + '…';
const fmt = n => Number(n).toLocaleString('en-IN');
const period = data.rankings.period === 'week' ? `${data.week.start} to ${data.week.end}` : 'Full supplied snapshot';
const scope = data.scope === 'all' ? 'Complete issue export' : data.scope === 'open' ? 'Open-ticket export' : 'Filtered issue export';

for (const [key, label] of [['open', 'Open'], ['resolved_closed', 'Resolved + Closed']]) {
 const group = data.rankings[key];
 const p = await PresentationFile.importPptx(await FileBlob.load(path.join(base, 'top_natures.pptx')));
 const originals = [...p.slides.items], slides = [];
 const clone = (index, values) => {
  const slide = originals[index].duplicate();
  for (const edit of mapping[index].editTargets) {
   if (!edit.token) continue;
   const shape = slide.shapes.items.find(s => s.name === edit.shapeName);
   if (!shape) throw new Error(`Template missing ${edit.shapeName}`);
   shape.text.replace(edit.token, values[edit.token] ?? '');
  }
  slides.push(slide);
  return slide;
 };
 const token = (number, paragraph = 1) => `{{TextBox_${number}_${paragraph}}}`;
 const footer = `CITES Operations Monitoring | Snapshot ${data.date} | ${data.rankings.period === 'week' ? 'Last-week submissions' : 'All supplied tickets'} | Confidential`;
 const coverage = group.available ? `${fmt(group.total)} ${label.toLowerCase()} tickets in scope. Top ${group.rows.length} natures cover ${fmt(group.ranked_total)} tickets; ${fmt(group.review_count)} need nature review.` : 'Completed-ticket details are unavailable for this snapshot. Counts by issue nature and representative examples require a resolved/closed or complete issue export.';
 clone(0, {
  [token(3)]: 'CITES OPERATIONS INTELLIGENCE',
  [token(4)]: `Top Issue Natures: ${label}`,
  [token(5)]: coverage,
  [token(8)]: 'RANKING PERIOD', [token(8,2)]: period,
  [token(9)]: 'RANKING LIMIT', [token(9,2)]: `Top ${data.rankings.top_n} issue natures`,
  [token(10)]: 'DATA SNAPSHOT', [token(10,2)]: data.date,
  [token(11)]: 'SOURCE SCOPE', [token(11,2)]: scope,
 });
 if (group.available) {
  for (let offset = 0; offset < Math.max(1, group.rows.length); offset += 10) {
   const rows = group.rows.slice(offset, offset + 10);
   const slide = clone(1, {
    [token(1)]: `${label.toUpperCase()} · DISTINCT TICKETS`,
    [token(2)]: rows.length ? `Issue natures · ranks ${offset + 1}–${offset + rows.length}` : 'No ranked issue natures in this scope',
    [token(3)]: `TOP ${data.rankings.top_n} · ${data.date}`,
    [token(6)]: footer,
   });
   const table = slide.tables.items[0];
   ['Rank', 'Issue Nature', 'Tickets', 'Resolved', 'Closed', 'Functions'].forEach((v,c) => table.cells.set(0,c,v));
   rows.forEach((r,index) => [String(r.rank),short(r.nature,50),fmt(r.count), key === 'open' ? '—' : fmt(r.resolved), key === 'open' ? '—' : fmt(r.closed),String(r.functionalities.length)].forEach((v,c) => table.cells.set(index+1,c,v)));
  }
  for (const row of group.rows) {
   const values = {
    [token(1)]: `RANK ${row.rank} · ${fmt(row.count)} ${label.toUpperCase()} TICKETS`,
    [token(2)]: short(row.nature,72),
    [token(3)]: data.date,
    [token(14)]: footer,
   };
   for (const [index, number] of [6,9,12].entries()) {
    const sample = row.samples[index];
    values[token(number)] = sample ? `${short(sample.functionality,57)} | #${sample.id} | ${sample.status} | Filed: ${sample.submitted || 'unknown'}` : 'No additional distinct example';
    values[token(number,2)] = sample ? short(sample.summary || 'No summary supplied',135) : '';
    values[token(number,3)] = sample ? short(sample.description || 'No description supplied',245) : '';
   }
   const slide = clone(2,values);
   slide.speakerNotes.textFrame.setText(`Issue nature: ${row.nature}\nCount: ${row.count}; resolved ${row.resolved}; closed ${row.closed}.\nFunctionalities: ${row.functionalities.map(f=>`${f.name} (${f.count})`).join('; ')}\nExamples: ${row.samples.map(s=>'#'+s.id).join(', ')}\n[Sources]\nUser-supplied issue export for ${data.date}; deterministic local issue-nature rules. Layout and EPFO mark adapted from the user-provided CITES_7Day_Top_Issues_2026-08-31.pptx.\n[/Sources]`);
  }
 }
 for (const slide of originals) slide.delete();
 for (const [index,slide] of slides.entries()) {
  slide.moveTo(index);
  // Footer text was deliberately left as an editable marker until the final slide count was known.
  for (const shape of slide.shapes.items) {
   if (shape.name === (index === 0 ? '' : (slide.tables.items.length ? 'TextBox 7' : 'TextBox 15'))) shape.text = `SLIDE ${index+1} OF ${slides.length}`;
  }
  if (index === 0 || slide.tables.items.length) slide.speakerNotes.textFrame.setText(`${coverage}\nPeriod: ${period}. Current ticket status is used; these are not resolution-event counts. Review/unclassified tickets are excluded from the nature ranking and retained in totals.\n[Sources]\nUser-supplied issue export for ${data.date}. Layout and EPFO mark adapted from the user-provided CITES_7Day_Top_Issues_2026-08-31.pptx.\n[/Sources]`);
 }
 const name = `top_${data.rankings.top_n}_${key}_${data.date}`;
 const deck = path.join(output, name + '.pptx');
 await (await PresentationFile.exportPptx(p)).save(deck);
 // The library's inspection sidecar is technical output, not a user-facing report.
 const sidecar = deck + '.inspect.ndjson';
 try { await fs.rename(sidecar, path.join(output, '_internal', name + '.inspect.ndjson')); } catch(e) { if(e.code !== 'ENOENT') throw e; }
 if (process.env.CITES_PPTX_PREVIEW === '1') {
  const previews = path.join(output, '_internal', name);
  await fs.mkdir(previews, {recursive:true});
  for (const [index,slide] of slides.entries()) {
   for (const format of ['png','layout']) {
    const blob = await p.export({slide, format, scale:1});
    await fs.writeFile(path.join(previews, `${String(index+1).padStart(2,'0')}.${format === 'layout' ? 'json' : 'png'}`),new Uint8Array(await blob.arrayBuffer()));
   }
  }
 }
 console.log(deck);
}
