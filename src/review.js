// review(file) -> { report (markdown), result (JSON), redline (.docx Buffer) }
import { readDocx, redlineDocx, textToDocx } from './docx.js';
import { reviewWithModel } from './providers.js';

export async function reviewContract({ buffer, filename = 'contract.docx', ...opts }) {
  const ext = filename.toLowerCase().split('.').pop();
  let docx = buffer;
  if (ext === 'txt' || ext === 'md') docx = await textToDocx(buffer.toString('utf8'));
  else if (ext !== 'docx') throw new Error('Please upload a Word .docx file (or .txt / .md). For a PDF, save it as Word first.');

  const paragraphs = await readDocx(docx);
  if (!paragraphs.join('').trim()) throw new Error('No text found in this document. Is it a scanned image?');

  const result = await reviewWithModel(paragraphs, opts);
  const { buffer: redline, applied, skipped } = await redlineDocx(docx, result.edits, { author: opts.author || 'AI Review' });
  return { result, redline, applied, skipped, report: toMarkdown(result, applied, skipped, filename) };
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
