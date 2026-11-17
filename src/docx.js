// Read a .docx into numbered paragraphs, and write the AI's edits back as real Word
// tracked changes (insertions + deletions) with a margin comment for each change.
// Paragraphs the AI does not touch are left byte-for-byte as they were.
import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&amp;/g, '&');
const encode = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = (s) => encode(s).replace(/"/g, '&quot;');

// ---- tiny XML helpers (enough for WordprocessingML; no external parser) ----

/** Split xml into its top-level elements (and text between them). */
function children(xml) {
  const out = [];
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) { out.push({ text: xml.slice(i) }); break; }
    if (lt > i) out.push({ text: xml.slice(i, lt) });
    const end = elementEnd(xml, lt);
    const raw = xml.slice(lt, end);
    const m = /^<([\w:.-]+)/.exec(raw);
    out.push({ tag: m ? m[1] : '', raw });
    i = end;
  }
  return out;
}

/** Index just past the element that starts at xml[start] ('<'). */
function elementEnd(xml, start) {
  const close = xml.indexOf('>', start);
  if (xml[close - 1] === '/' || xml.startsWith('<?', start) || xml.startsWith('<!', start)) return close + 1;
  const name = /^<([\w:.-]+)/.exec(xml.slice(start, close + 1))[1];
  let depth = 1, i = close + 1;
  const open = new RegExp(`<${name.replace('.', '\\.')}(?=[\\s>/])`, 'g');
  while (depth > 0) {
    const nextClose = xml.indexOf(`</${name}>`, i);
    if (nextClose < 0) throw new Error(`Malformed XML: <${name}> is not closed`);
    open.lastIndex = i;
    let m;
    while ((m = open.exec(xml)) && m.index < nextClose) {
      const gt = xml.indexOf('>', m.index);
      if (xml[gt - 1] !== '/') depth++;
      open.lastIndex = gt + 1;
    }
    depth--;
    i = nextClose + name.length + 3;
  }
  return i;
}

const inner = (raw) => /\/>$/.test(raw) && !raw.includes('</') ? '' : raw.slice(raw.indexOf('>') + 1, raw.lastIndexOf('</'));
const textOf = (xml) => {
  let t = '';
  // w:t only: text inside existing tracked deletions (w:delText) is not part of the current wording
  xml.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:tab\s[^>]*\/>/g, (m, txt) => { t += txt !== undefined ? decode(txt) : '\t'; return m; });
  return t;
};

// ---- paragraphs ----

/** Find every body paragraph: { start, end, raw } in document order. */
function findParagraphs(docXml) {
  const body0 = docXml.indexOf('<w:body>');
  const body1 = docXml.lastIndexOf('</w:body>');
  const list = [];
  const walk = (from, to) => {
    let i = from;
    while (i < to) {
      const lt = docXml.indexOf('<', i);
      if (lt < 0 || lt >= to) break;
      if (docXml[lt + 1] === '/' || docXml[lt + 1] === '?' || docXml[lt + 1] === '!') { i = docXml.indexOf('>', lt) + 1; continue; }
      const end = elementEnd(docXml, lt);
      const tag = /^<([\w:.-]+)/.exec(docXml.slice(lt, lt + 40))[1];
      if (tag === 'w:p') list.push({ start: lt, end, raw: docXml.slice(lt, end) });
      else if (['w:tbl', 'w:tr', 'w:tc', 'w:sdt', 'w:sdtContent', 'w:customXml'].includes(tag)) {
        const open = docXml.indexOf('>', lt) + 1;
        walk(open, end - (tag.length + 3));
      }
      i = end;
    }
  };
  walk(docXml.indexOf('>', body0) + 1, body1);
  return list;
}

/** Break a paragraph into pPr + "atoms". Text atoms can be split and redlined; opaque ones cannot. */
function atomize(pRaw) {
  const open = pRaw.slice(0, pRaw.indexOf('>') + 1);
  const atoms = [];
  let pPr = '';
  for (const c of children(inner(pRaw))) {
    if (!c.tag) continue;
    if (c.tag === 'w:pPr') { pPr = c.raw; continue; }
    if (c.tag === 'w:r') {
      let rPr = '';
      for (const rc of children(inner(c.raw))) {
        if (!rc.tag) continue;
        if (rc.tag === 'w:rPr') { rPr = rc.raw; continue; }
        if (rc.tag === 'w:t') atoms.push({ kind: 'text', rPr, text: decode(inner(rc.raw)) });
        else if (rc.tag === 'w:tab') atoms.push({ kind: 'run', rPr, xml: rc.raw, text: '\t' });
        else atoms.push({ kind: 'run', rPr, xml: rc.raw, text: '' });
      }
      continue;
    }
    atoms.push({ kind: 'opaque', xml: c.raw, text: textOf(c.raw) }); // hyperlinks, fields, existing tracked changes...
  }
  return { open, pPr, atoms };
}

const runXml = (a) => a.kind === 'text'
  ? `<w:r>${a.rPr}<w:t xml:space="preserve">${encode(a.text)}</w:t></w:r>`
  : `<w:r>${a.rPr}${a.xml}</w:r>`;

/** Read the paragraphs of a .docx (Buffer) as plain text, in order. */
export async function readDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file('word/document.xml');
  if (!f) throw new Error('This file is not a Word .docx document.');
  const xml = await f.async('string');
  return findParagraphs(xml).map((p) => atomize(p.raw).atoms.map((a) => a.text).join(''));
}

// Loose matching: models often change curly quotes, dashes or spacing when copying text.
const norm = (ch) => ({ '\u2018': "'", '\u2019': "'", '\u201C': '"', '\u201D': '"', '\u2013': '-', '\u2014': '-', '\u00A0': ' ', '\t': ' ' })[ch] || ch;
function locate(hay, needle) {
  let i = hay.indexOf(needle);
  if (i >= 0) return [i, i + needle.length];
  // normalised search that maps back to original positions
  const map = [], h = [];
  for (let k = 0; k < hay.length; k++) {
    const c = norm(hay[k]);
    if (c === ' ' && h[h.length - 1] === ' ') continue;
    h.push(c); map.push(k);
  }
  const n = [...needle].map(norm).join('').replace(/ +/g, ' ').trim();
  if (!n) return null;
  i = h.join('').indexOf(n);
  if (i < 0) return null;
  return [map[i], map[i + n.length - 1] + 1];
}

/** Split atoms so that a boundary falls exactly at character offset pos. */
function splitAt(atoms, pos) {
  let off = 0;
  for (let k = 0; k < atoms.length; k++) {
    const a = atoms[k], len = a.text.length;
    if (pos > off && pos < off + len) {
      if (a.kind !== 'text') return false;
      const cut = pos - off;
      atoms.splice(k, 1, { ...a, text: a.text.slice(0, cut) }, { ...a, text: a.text.slice(cut) });
      return true;
    }
    off += len;
  }
  return true;
}

function rangeIdx(atoms, s, e) {
  let off = 0, first = -1, last = -1;
  atoms.forEach((a, k) => {
    const len = a.text.length;
    if (len > 0 && off >= s && off + len <= e) { if (first < 0) first = k; last = k; }
    off += len;
  });
  return [first, last];
}

/**
 * Apply edits to a .docx and return { buffer, applied[], skipped[] }.
 * edits: [{ type: 'replace'|'delete'|'insert_after', paragraph, find, replace, comment }]
 */
export async function redlineDocx(buf, edits, { author = 'AI Review', date = new Date() } = {}) {
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  const paras = findParagraphs(xml);
  const when = new Date(date).toISOString().replace(/\.\d+Z$/, 'Z');
  const A = `w:author="${attr(author)}" w:date="${when}"`;
  let revId = 90000;

  // comments.xml (create or extend)
  let commentsXml = zip.file('word/comments.xml') ? await zip.file('word/comments.xml').async('string') : '';
  let nextComment = 0;
  commentsXml.replace(/<w:comment\s[^>]*w:id="(\d+)"/g, (m, id) => { nextComment = Math.max(nextComment, +id + 1); return m; });
  const newComments = [];
  const addComment = (text) => {
    const id = nextComment++;
    const paras = String(text).split(/\n+/).map((t) => `<w:p><w:r><w:t xml:space="preserve">${encode(t)}</w:t></w:r></w:p>`).join('');
    newComments.push(`<w:comment w:id="${id}" ${A} w:initials="AI">${paras}</w:comment>`);
    return id;
  };

  const applied = [], skipped = [];
  const byPara = new Map();
  edits.forEach((e, i) => {
    const n = Number(e.paragraph);
    if (!Number.isInteger(n) || n < 0 || n >= paras.length) { skipped.push({ ...e, reason: 'paragraph number not found' }); return; }
    if (!byPara.has(n)) byPara.set(n, []);
    byPara.get(n).push({ ...e, _i: i });
  });

  const replacements = []; // { start, end, xml } on the original document string
  for (const [n, list] of byPara) {
    const { open, pPr, atoms } = atomize(paras[n].raw);
    const text = atoms.map((a) => a.text).join('');
    const inserts = list.filter((e) => e.type === 'insert_after');
    // locate in-paragraph edits on the original text, drop overlaps
    const spans = [];
    for (const e of list.filter((x) => x.type !== 'insert_after')) {
      const hit = e.find ? locate(text, e.find) : null;
      if (!hit) { skipped.push({ ...e, reason: 'text not found in that paragraph' }); continue; }
      if (spans.some((s) => hit[0] < s.e && s.s < hit[1])) { skipped.push({ ...e, reason: 'overlaps another edit' }); continue; }
      spans.push({ s: hit[0], e: hit[1], edit: e });
    }
    spans.sort((x, y) => y.s - x.s); // apply from the end so earlier offsets stay valid
    let items = atoms.slice();
    for (const sp of spans) {
      if (!splitAt(items, sp.s) || !splitAt(items, sp.e)) { skipped.push({ ...sp.edit, reason: 'starts or ends inside a link or field' }); continue; }
      const [f, l] = rangeIdx(items, sp.s, sp.e);
      const all = f < 0 ? [] : items.slice(f, l + 1);
      // zero-length markers inside the range (bookmarks, spell-check marks) stay in place, just before the change
      const marks = all.filter((a) => a.kind === 'opaque' && !a.text);
      const slice = all.filter((a) => !(a.kind === 'opaque' && !a.text));
      if (!slice.length || slice.some((a) => a.kind === 'opaque' || a.group)) { skipped.push({ ...sp.edit, reason: 'touches a link, field or existing tracked change' }); continue; }
      const rPr = slice.find((a) => a.rPr)?.rPr || '';
      const cid = addComment(sp.edit.comment || 'Proposed change');
      const del = `<w:del w:id="${revId++}" ${A}>${slice.map((a) => runXml(a).replace(/<w:t( |>)/g, '<w:delText$1').replace(/<\/w:t>/g, '</w:delText>')).join('')}</w:del>`;
      const ins = sp.edit.type === 'delete' || !sp.edit.replace ? '' :
        `<w:ins w:id="${revId++}" ${A}><w:r>${rPr}<w:t xml:space="preserve">${encode(sp.edit.replace)}</w:t></w:r></w:ins>`;
      const group = marks.map((a) => a.xml).join('') + `<w:commentRangeStart w:id="${cid}"/>${del}${ins}<w:commentRangeEnd w:id="${cid}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${cid}"/></w:r>`;
      items.splice(f, l - f + 1, { group, text: slice.map((a) => a.text).join('') });
      applied.push(sp.edit);
    }
    let out = spans.length ? open + pPr + items.map((a) => a.group || (a.kind === 'opaque' ? a.xml : runXml(a))).join('') + '</w:p>' : paras[n].raw;
    // new paragraphs after this one
    for (const e of inserts) {
      if (!String(e.replace || '').trim()) { skipped.push({ ...e, reason: 'empty new paragraph' }); continue; }
      let np = pPr.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, '').replace(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/, '');
      const mark = `<w:ins w:id="${revId++}" ${A}/>`;
      if (!np) np = `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>`;
      else if (/<w:rPr>/.test(np)) np = np.replace('<w:rPr>', `<w:rPr>${mark}`);
      else if (/<w:rPr\/>/.test(np)) np = np.replace('<w:rPr/>', `<w:rPr>${mark}</w:rPr>`);
      else np = np.replace('</w:pPr>', `<w:rPr>${mark}</w:rPr></w:pPr>`);
      const rPr = atoms.find((a) => a.kind === 'text')?.rPr || '';
      const cid = addComment(e.comment || 'Proposed new wording');
      out += `<w:p>${np}<w:commentRangeStart w:id="${cid}"/><w:ins w:id="${revId++}" ${A}><w:r>${rPr}<w:t xml:space="preserve">${encode(e.replace)}</w:t></w:r></w:ins><w:commentRangeEnd w:id="${cid}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${cid}"/></w:r></w:p>`;
      applied.push(e);
    }
    if (out !== paras[n].raw) replacements.push({ start: paras[n].start, end: paras[n].end, xml: out });
  }
  replacements.sort((a, b) => b.start - a.start).forEach((r) => { xml = xml.slice(0, r.start) + r.xml + xml.slice(r.end); });
  zip.file('word/document.xml', xml);

  if (newComments.length) {
    if (commentsXml) commentsXml = commentsXml.replace(/<\/w:comments>\s*$/, newComments.join('') + '</w:comments>');
    else {
      commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W_NS}">${newComments.join('')}</w:comments>`;
      const relsPath = 'word/_rels/document.xml.rels';
      let rels = await zip.file(relsPath).async('string');
      if (!rels.includes(COMMENTS_REL)) {
        let rid = 1; rels.replace(/Id="rId(\d+)"/g, (m, d) => { rid = Math.max(rid, +d + 1); return m; });
        rels = rels.replace('</Relationships>', `<Relationship Id="rId${rid}" Type="${COMMENTS_REL}" Target="comments.xml"/></Relationships>`);
        zip.file(relsPath, rels);
      }
      let ct = await zip.file('[Content_Types].xml').async('string');
      if (!ct.includes('/word/comments.xml')) ct = ct.replace('</Types>', `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CT}"/></Types>`);
      zip.file('[Content_Types].xml', ct);
    }
    zip.file('word/comments.xml', commentsXml);
  }
  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buffer, applied, skipped };
}

/**
 * Read a (redlined) .docx back for on-screen review, exactly as Word shows it.
 * Returns { paragraphs: [{ segs: [{ t: 'eq'|'del'|'ins'|'cm', text?, id? }], changed }], comments: { id: { author, text } } }.
 */
export async function previewDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('word/document.xml').async('string');
  const cx = zip.file('word/comments.xml') ? await zip.file('word/comments.xml').async('string') : '';
  const comments = {};
  cx.replace(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g, (m, a, body) => {
    const id = /w:id="(\d+)"/.exec(a)?.[1];
    if (id !== undefined) comments[id] = { author: decode(/w:author="([^"]*)"/.exec(a)?.[1] || ''), text: body.split(/<\/w:p>/).map(textOf).filter((t) => t.trim()).join('\n') };
    return m;
  });
  return { paragraphs: findParagraphs(xml).map((p) => segments(p.raw)), comments };
}

function segments(pRaw) {
  const segs = [];
  let del = 0, ins = 0;
  const add = (t, text) => {
    if (!text) return;
    const last = segs[segs.length - 1];
    if (last && last.t === t) last.text += text; else segs.push({ t, text });
  };
  const kind = () => (del ? 'del' : ins ? 'ins' : 'eq');
  const re = /<(\/?)(w:del|w:ins|w:t|w:delText|w:tab|w:br|w:cr|w:commentReference|w:pPr|w:rPr)(?=[\s>/])([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(pRaw))) {
    const [all, close, tag, attrs, self] = m;
    if (tag === 'w:pPr' || tag === 'w:rPr') {   // formatting only (a self-closing <w:ins/> in here marks a paragraph mark)
      if (!close && !self) { const end = pRaw.indexOf(`</${tag}>`, re.lastIndex); if (end >= 0) re.lastIndex = end + tag.length + 3; }
      continue;
    }
    if (tag === 'w:del' && !self) { del += close ? -1 : 1; continue; }
    if (tag === 'w:ins' && !self) { ins += close ? -1 : 1; continue; }
    if ((tag === 'w:t' || tag === 'w:delText') && !close && !self) {
      const end = pRaw.indexOf(`</${tag}>`, re.lastIndex);
      add(kind(), decode(pRaw.slice(re.lastIndex, end)));
      re.lastIndex = end + tag.length + 3;
      continue;
    }
    if (tag === 'w:tab' && !close) add(kind(), '\t');
    else if ((tag === 'w:br' || tag === 'w:cr') && !close) add(kind(), '\n');
    else if (tag === 'w:commentReference') { const id = /w:id="(\d+)"/.exec(attrs)?.[1]; if (id !== undefined) segs.push({ t: 'cm', id }); }
    void all;
  }
  return { segs, changed: segs.some((s) => s.t !== 'eq') };
}

/** Turn plain text into a minimal .docx so .txt / .md contracts can be redlined too. */
export async function textToDocx(text) {
  const zip = new JSZip();
  const paras = String(text).replace(/\r\n?/g, '\n').split('\n').map((l) => `<w:p><w:r><w:t xml:space="preserve">${encode(l)}</w:t></w:r></w:p>`).join('');
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}"><w:body>${paras}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
