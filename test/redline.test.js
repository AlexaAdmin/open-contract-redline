// npm test   (no API key needed: the AI model is replaced by a fixed answer)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import JSZip from 'jszip';
import { readDocx, redlineDocx } from '../src/docx.js';
import { reviewContract } from '../src/review.js';

const SAMPLE = fs.readFileSync(new URL('../examples/sample-services-agreement.docx', import.meta.url));

const EDITS = [
  { type: 'replace', paragraph: 3, find: 'seven (7) days', replace: 'thirty (30) days', comment: '30 days is market standard.', issue_id: 'I1' },
  { type: 'replace', paragraph: 3, find: '5% per month', replace: '1% per month', comment: 'Interest rate too high.', issue_id: 'I1' },
  { type: 'replace', paragraph: 2, find: '(the “Services”) with reasonable skill', replace: '(the “Services”) promptly and with reasonable skill', comment: 'Crosses plain and italic text.', issue_id: 'I2' },
  { type: 'replace', paragraph: 4, find: '3 Term. This Agreement continues for 36 months', replace: '3 Term. This Agreement continues for 12 months', comment: 'Crosses a bookmark.', issue_id: 'I3' },
  { type: 'replace', paragraph: 5, find: 'the fees paid in the one (1) month', replace: 'the fees paid in the twelve (12) months', comment: 'Cap too low.', issue_id: 'I4' },
  { type: 'replace', paragraph: 5, find: "The Customer's liability is unlimited", replace: 'The same cap applies to the Customer', comment: 'Mutual cap. (straight quote on purpose)', issue_id: 'I4' },
  { type: 'delete', paragraph: 7, find: ', including training its own products', replace: '', comment: 'No training on our data.', issue_id: 'I5' },
  { type: 'replace', paragraph: 8, find: 'the cover page', replace: 'Schedule 2', comment: 'Inside a link: must be skipped.', issue_id: 'I6' },
  { type: 'replace', paragraph: 6, find: 'text that is not there', replace: 'x', comment: 'Must be skipped.', issue_id: 'I7' },
  { type: 'replace', paragraph: 99, find: 'x', replace: 'y', comment: 'Bad paragraph.', issue_id: 'I7' },
  { type: 'replace', paragraph: 12, find: 'Name:', replace: 'Name and title:', comment: 'Table cell.', issue_id: 'I8' },
  { type: 'insert_after', paragraph: 9, find: '', replace: '9 Assignment. Neither party may assign this Agreement without the other’s written consent.', comment: 'New clause.', issue_id: 'I9' },
];

// What the document says if every change is accepted / rejected in Word.
async function versions(buf) {
  const xml = await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string');
  const paras = xml.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || [];
  const txt = (s) => (s.match(/<w:t(?: [^>]*)?>[\s\S]*?<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  const accepted = paras.map((p) => txt(p.replace(/<w:del [\s\S]*?<\/w:del>/g, '')));
  const rejected = paras.filter((p) => !/<w:pPr>[\s\S]*<w:ins [^>]*\/>[\s\S]*<\/w:pPr>/.test(p))
    .map((p) => txt(p.replace(/<w:ins [^/]*?>[\s\S]*?<\/w:ins>/g, '').replace(/w:delText/g, 'w:t')));
  return { accepted, rejected, xml };
}

test('reads paragraphs, including table cells', async () => {
  const p = await readDocx(SAMPLE);
  assert.equal(p[3], '2 Fees. The Customer shall pay all invoices within seven (7) days of the invoice date. Late payments bear interest at 5% per month.');
  assert.equal(p[12], 'Name:');
});

test('writes tracked changes that accept and reject cleanly', async () => {
  const original = await readDocx(SAMPLE);
  const { buffer, applied, skipped } = await redlineDocx(SAMPLE, EDITS, { author: 'Test', date: '2026-01-01T00:00:00Z' });
  assert.equal(applied.length, 9);
  assert.deepEqual(skipped.map((s) => s.reason).sort(), ['paragraph number not found', 'text not found in that paragraph', 'touches a link, field or existing tracked change']);

  const v = await versions(buffer);
  assert.deepEqual(v.rejected, original, 'rejecting every change gives back the original text');
  assert.equal(v.accepted[3], '2 Fees. The Customer shall pay all invoices within thirty (30) days of the invoice date. Late payments bear interest at 1% per month.');
  assert.equal(v.accepted[2], '1 Services. The Supplier shall provide the data analytics services described in Schedule 1 (the “Services”) promptly and with reasonable skill and care.');
  assert.equal(v.accepted[4], '3 Term. This Agreement continues for 12 months and renews automatically for further 36-month periods unless the Customer gives 180 days’ notice.');
  assert.equal(v.accepted[5], '4 Liability. The Supplier’s total liability under this Agreement shall not exceed the fees paid in the twelve (12) months before the claim. The same cap applies to the Customer.');
  assert.equal(v.accepted[7], '6 Data. The Supplier may use Customer data for any purpose.');
  assert.equal(v.accepted[8], '7 Notices. Notices must be sent to the addresses on the cover page.');
  assert.equal(v.accepted[10], '9 Assignment. Neither party may assign this Agreement without the other’s written consent.');
  assert.equal(v.accepted[13], 'Name and title:');
  assert.match(v.xml, /<w:bookmarkStart [^>]*w:name="_RefTerm"/, 'bookmark kept');
  assert.match(v.xml, /<w:bookmarkEnd [^>]*w:id="10"/, 'bookmark end kept');
  assert.match(v.xml, /<w:hyperlink /, 'hyperlink kept');

  const zip = await JSZip.loadAsync(buffer);
  const comments = await zip.file('word/comments.xml').async('string');
  assert.equal((comments.match(/<w:comment /g) || []).length, 9);
  assert.match(await zip.file('word/_rels/document.xml.rels').async('string'), /relationships\/comments" Target="comments.xml"/);
  assert.match(await zip.file('[Content_Types].xml').async('string'), /\/word\/comments.xml/);
  fs.mkdirSync(new URL('./out/', import.meta.url), { recursive: true });
  fs.writeFileSync(new URL('./out/sample - redline.docx', import.meta.url), buffer);
});

test('a second pass adds comments to the existing comments part', async () => {
  const first = (await redlineDocx(SAMPLE, EDITS.slice(0, 1))).buffer;
  const second = (await redlineDocx(first, [EDITS[4]])).buffer;
  const c = await (await JSZip.loadAsync(second)).file('word/comments.xml').async('string');
  assert.deepEqual([...c.matchAll(/<w:comment w:id="(\d+)"/g)].map((m) => m[1]), ['0', '1']);
});

test('full review flow with each provider name (model replaced by a fixed answer)', async () => {
  for (const provider of ['anthropic', 'openai', 'deepseek', 'mistral']) {
    const fake = '```json\n' + JSON.stringify({ summary: 'One-sided supplier paper.', overall_risk: 'high',
      issues: [{ id: 'I1', title: 'Payment terms', risk: 'medium', clause: '2', explanation: 'Seven days is short.', recommendation: 'Ask for 30 days.' }],
      edits: EDITS.slice(0, 2) }) + '\n```';
    const r = await reviewContract({ buffer: SAMPLE, filename: 'sample.docx', provider, apiKey: 'test-key', _call: async () => fake });
    assert.equal(r.result.provider, provider);
    assert.equal(r.applied.length, 2);
    assert.match(r.report, /Overall risk:\*\* HIGH/);
  }
});

test('missing key gives a clear message', async () => {
  delete process.env.MISTRAL_API_KEY;
  await assert.rejects(reviewContract({ buffer: SAMPLE, filename: 'a.docx', provider: 'mistral' }), /Set MISTRAL_API_KEY/);
});
