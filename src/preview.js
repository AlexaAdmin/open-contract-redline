// The redline as a web page: deletions struck through in red, insertions in green, and the AI's
// comment next to each change. Built from the finished .docx (previewDocx), so it matches Word.
// One self-contained HTML file: used inside the local web page and saved by the command line.

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const RISK = (r) => String(r || '').toLowerCase();

function paragraphHtml(p, i, comments, num) {
  const marks = [];
  const body = p.segs.map((s) => {
    if (s.t === 'cm') {
      if (!comments[s.id]) return '';
      if (!num.has(s.id)) num.set(s.id, num.size + 1);
      marks.push(s.id);
      return `<sup class="cm" data-c="${esc(s.id)}" title="${esc(comments[s.id].text)}">${num.get(s.id)}</sup>`;
    }
    const t = esc(s.text);
    return s.t === 'del' ? `<del>${t}</del>` : s.t === 'ins' ? `<ins>${t}</ins>` : t;
  }).join('');
  const notes = marks.map((id) => `<div class="note" id="c${esc(id)}"><b>${num.get(id)}</b> ${esc(comments[id].text)}</div>`).join('');
  const empty = !p.segs.some((s) => s.text && s.text.trim());
  return `<div class="row${p.changed ? ' chg' : ''}${empty ? ' empty' : ''}" data-i="${i}"><div class="para">${body || '&nbsp;'}</div><div class="notes">${notes}</div></div>`;
}

/**
 * preview: from previewDocx. meta: { filename, result, applied, skipped }
 */
export function redlinePage(preview, { filename = 'contract', result = {}, applied = 0, skipped = [] } = {}) {
  const num = new Map();
  const rows = preview.paragraphs.map((p, i) => paragraphHtml(p, i, preview.comments, num)).join('\n');
  const changes = preview.paragraphs.filter((p) => p.changed).length;
  const issues = (result.issues || []).map((i) => `<div class="issue"><b>${esc(i.id)}. ${esc(i.title)}</b> <span class="risk ${esc(RISK(i.risk))}">${esc(RISK(i.risk).toUpperCase())}</span>
    <div class="clause">${esc(i.clause)}</div><p>${esc(i.explanation)}</p><p><b>Recommendation:</b> ${esc(i.recommendation)}</p></div>`).join('');
  const notPlaced = skipped.length ? `<details class="box"><summary>${skipped.length} suggested change(s) could not be placed in the redline</summary>
    <ul>${skipped.map((s) => `<li>[P${esc(s.paragraph)}] ${esc(s.type)}: “${esc(s.find)}” → “${esc(s.replace)}” <span class="clause">(${esc(s.reason)})</span></li>`).join('')}</ul></details>` : '';
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Redline: ${esc(filename)}</title>
<style>
  :root { --ink:#111827; --sub:#4b5563; --line:#e5e7eb; --brand:#0f766e; --del:#b91c1c; --ins:#047857; --hi:#dc2626; --med:#d97706; --low:#059669; }
  * { box-sizing:border-box; }
  body { margin:0; font:15px/1.6 system-ui,-apple-system,Segoe UI,Roboto,sans-serif; color:var(--ink); background:#f6f6f7; }
  header { padding:18px 22px 6px; } h1 { font-size:19px; margin:0 0 2px; } .meta { color:var(--sub); font-size:13px; }
  .bar { position:sticky; top:0; z-index:5; display:flex; flex-wrap:wrap; gap:8px; align-items:center; padding:10px 22px; background:#fff; border-bottom:1px solid var(--line); }
  .seg { display:inline-flex; border:1px solid var(--line); border-radius:8px; overflow:hidden; }
  .seg button, .btn { font:600 13px system-ui; padding:7px 12px; border:0; background:#fff; color:var(--sub); cursor:pointer; }
  .seg button + button { border-left:1px solid var(--line); }
  .seg button.on { background:#f0fdfa; color:var(--brand); }
  .btn { border:1px solid var(--line); border-radius:8px; }
  .bar label { font-size:13px; color:var(--sub); display:flex; gap:6px; align-items:center; cursor:pointer; }
  .count { font-size:13px; color:var(--sub); margin-left:auto; }
  .legend { font-size:12.5px; color:var(--sub); padding:6px 22px 0; } .legend del, .legend ins { padding:0 3px; }
  .wrap { padding:14px 22px 40px; }
  .box { background:#fff; border:1px solid var(--line); border-radius:10px; padding:12px 16px; margin-bottom:12px; }
  .box summary { cursor:pointer; font-weight:600; }
  .doc { background:#fff; border:1px solid var(--line); border-radius:10px; padding:10px 0; }
  .row { display:grid; grid-template-columns:minmax(0,1fr) 260px; gap:18px; padding:4px 22px; }
  .row.empty .para { min-height:.6em; }
  .row.chg { background:#fffdf5; }
  .row.cur { outline:2px solid #5eead4; outline-offset:-2px; }
  .para { white-space:pre-wrap; overflow-wrap:anywhere; }
  del { color:var(--del); text-decoration:line-through; text-decoration-thickness:1.5px; background:#fef2f2; }
  ins { color:var(--ins); text-decoration:underline; text-decoration-thickness:1.5px; text-underline-offset:2px; background:#ecfdf5; }
  .cm { display:inline-block; min-width:17px; padding:0 4px; margin:0 1px; border-radius:9px; background:var(--brand); color:#fff; font:700 10.5px/17px system-ui; text-align:center; cursor:pointer; vertical-align:super; }
  .note { font-size:13px; line-height:1.5; background:#f0fdfa; border:1px solid #ccfbf1; border-radius:8px; padding:8px 10px; margin:2px 0 8px; color:#134e4a; }
  .note b { display:inline-block; min-width:17px; text-align:center; border-radius:9px; background:var(--brand); color:#fff; font-size:10.5px; margin-right:4px; }
  .note.flash { box-shadow:0 0 0 3px #5eead4; }
  body.clean del, body.orig ins, body.clean .cm, body.orig .cm, body.clean .notes, body.orig .notes { display:none; }
  body.clean ins, body.orig del { color:inherit; background:none; text-decoration:none; }
  body.clean .row, body.orig .row { grid-template-columns:1fr; } body.clean .row.chg, body.orig .row.chg { background:none; }
  body.only .row:not(.chg) { display:none; }
  .risk { display:inline-block; padding:1px 7px; border-radius:6px; font-size:11.5px; font-weight:700; border:1px solid; }
  .risk.high { color:var(--hi); } .risk.medium { color:var(--med); } .risk.low { color:var(--low); }
  .issue { border-top:1px solid var(--line); padding:10px 0; } .issue:first-of-type { border-top:0; } .issue p { margin:4px 0; }
  .clause { color:var(--sub); font-size:13px; }
  footer { color:var(--sub); font-size:12px; padding:0 22px 30px; }
  @media (max-width:760px) { .row { grid-template-columns:1fr; gap:4px; } .count { margin-left:0; } }
  @media print { .bar { position:static; } .row.cur { outline:0; } }
</style></head>
<body>
<header><h1>Redline: ${esc(filename)}</h1>
  <div class="meta">${result.overall_risk ? `<span class="risk ${esc(RISK(result.overall_risk))}">${esc(RISK(result.overall_risk).toUpperCase())} RISK</span> &middot; ` : ''}${esc(result.provider || '')}${result.model ? ' / ' + esc(result.model) : ''} &middot; ${applied} tracked change(s) in ${changes} paragraph(s)</div></header>
<div class="bar">
  <div class="seg" role="group" aria-label="View"><button data-v="red" class="on">Redline</button><button data-v="clean">Clean (all accepted)</button><button data-v="orig">Original</button></div>
  <label><input type="checkbox" id="only"> Changes only</label>
  <button class="btn" id="prev">&uarr; Previous change</button><button class="btn" id="next">&darr; Next change</button>
  <span class="count" id="count"></span>
</div>
<div class="legend"><del>deleted text</del> <ins>inserted text</ins> <span class="cm" style="cursor:default">1</span> AI comment explaining the change</div>
<div class="wrap">
  ${result.summary ? `<details class="box"${issues ? '' : ' open'}><summary>Summary and ${(result.issues || []).length} issue(s)</summary><p>${esc(result.summary)}</p>${issues}</details>` : ''}
  ${notPlaced}
  <div class="doc">
${rows}
  </div>
</div>
<footer>AI-generated first draft for a qualified lawyer to check. Not legal advice. Made with Open Contract Redline.</footer>
<script>
(function () {
  var b = document.body, rows = [].slice.call(document.querySelectorAll('.row.chg')), at = -1;
  function count() { document.getElementById('count').textContent = rows.length ? (at < 0 ? rows.length + ' changed paragraph(s)' : 'Change ' + (at + 1) + ' of ' + rows.length) : 'No changes'; }
  function go(d) { if (!rows.length) return; if (at >= 0) rows[at].classList.remove('cur'); at = (at + d + rows.length) % rows.length; rows[at].classList.add('cur'); rows[at].scrollIntoView({ behavior: 'smooth', block: 'center' }); count(); }
  document.getElementById('next').onclick = function () { go(1); };
  document.getElementById('prev').onclick = function () { go(at < 0 ? 0 : -1); };
  document.getElementById('only').onchange = function (e) { b.classList.toggle('only', e.target.checked); };
  [].forEach.call(document.querySelectorAll('.seg button'), function (x) { x.onclick = function () {
    [].forEach.call(document.querySelectorAll('.seg button'), function (y) { y.classList.toggle('on', y === x); });
    b.classList.toggle('clean', x.dataset.v === 'clean'); b.classList.toggle('orig', x.dataset.v === 'orig'); }; });
  [].forEach.call(document.querySelectorAll('sup.cm'), function (s) { s.onclick = function () { var n = document.getElementById('c' + s.dataset.c); if (!n) return;
    n.classList.add('flash'); setTimeout(function () { n.classList.remove('flash'); }, 1200); }; });
  document.addEventListener('keydown', function (e) { if (e.target.tagName === 'INPUT') return; if (e.key === 'j' || e.key === 'ArrowDown' && e.altKey) go(1); if (e.key === 'k' || e.key === 'ArrowUp' && e.altKey) go(-1); });
  count();
})();
</script>
</body></html>`;
}
