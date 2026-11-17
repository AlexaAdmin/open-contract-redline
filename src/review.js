// review(file) -> { report (markdown), result (JSON), redline (.docx Buffer), html (redline as a web page) }
import { readDocx, redlineDocx, textToDocx, previewDocx } from './docx.js';
import { redlinePage } from './preview.js';
import fs from 'node:fs';

const SAMPLE_URL = new URL('../examples/sample-services-agreement.docx', import.meta.url);
const DEMO_URL = new URL('../examples/demo-review.json', import.meta.url);

/** The demo: the sample contract with a prepared answer. No API key, no AI call. */
export function demoReview() {
  return reviewContract({ buffer: fs.readFileSync(SAMPLE_URL), filename: 'sample-services-agreement.docx', demo: true });
}
import { reviewWithModel, normalizeResult } from './providers.js';

async function load(buffer, filename) {
  const ext = filename.toLowerCase().split('.').pop();
  let docx = buffer;
  if (ext === 'txt' || ext === 'md') docx = await textToDocx(buffer.toString('utf8'));
  else if (ext !== 'docx') throw new Error('Please upload a Word .docx file (or .txt / .md). For a PDF, save it as Word first.');
  const paragraphs = await readDocx(docx);
  if (!paragraphs.join('').trim()) throw new Error('No text found in this document. Is it a scanned image?');
  return { docx, paragraphs };
}

/** Agent mode, step 1: the contract as numbered paragraphs, for the brief an AI agent reads. */
export async function contractParagraphs({ buffer, filename = 'contract.docx' }) {
  return (await load(buffer, filename)).paragraphs;
}

/**
 * opts.review: an answer already written by an AI agent (agent mode, no API key).
 * opts.demo:   the prepared sample answer. Otherwise one API call to the chosen provider.
 */
export async function reviewContract({ buffer, filename = 'contract.docx', ...opts }) {
  const { docx, paragraphs } = await load(buffer, filename);
  const result = opts.review
    ? { ...normalizeResult(opts.review), provider: 'agent mode', model: opts.agentName || 'your AI agent (no API key)' }
    : opts.demo
      ? { ...JSON.parse(fs.readFileSync(DEMO_URL, 'utf8')), provider: 'demo', model: 'prepared sample answer (no AI call)' }
      : await reviewWithModel(paragraphs, opts);
  delete result._note;
  const { buffer: redline, applied, skipped } = await redlineDocx(docx, result.edits, { author: opts.author || 'AI Review' });
  const preview = await previewDocx(redline);   // read back from the finished file, so the screen matches Word
  const html = redlinePage(preview, { filename, result, applied: applied.length, skipped });
  return { result, redline, applied, skipped, preview, html, report: toMarkdown(result, applied, skipped, filename) };
}

export function toMarkdown(r, applied, skipped, filename) {
  const L = [];
  L.push(`# Contract review: ${filename}`, '');
  L.push(`**Overall risk:** ${r.overall_risk.toUpperCase()}  `);
  L.push(`**Reviewed by:** ${r.provider} / ${r.model}  `);
  L.push(`**Redline:** ${applied.length} tracked change(s) applied${skipped.length ? `, ${skipped.length} could not be placed (listed below)` : ''}`, '');
  L.push('## Summary', '', r.summary, '');
  L.push('## Issues', '');
  for (const i of r.issues) {
    L.push(`### ${i.id}. ${i.title} (${String(i.risk).toUpperCase()})`, `*Clause:* ${i.clause}`, '', i.explanation, '', `**Recommendation:** ${i.recommendation}`, '');
  }
  if (skipped.length) {
    L.push('## Suggested changes not placed in the redline', '', 'Apply these by hand:', '');
    for (const s of skipped) L.push(`- [P${s.paragraph}] ${s.type}: "${s.find}" -> "${s.replace}" (${s.reason})`);
    L.push('');
  }
  L.push('---', '*AI-generated first draft for a qualified lawyer to check. Not legal advice.*');
  return L.join('\n');
}
