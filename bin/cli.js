#!/usr/bin/env node
// Usage: node bin/cli.js <contract.docx> [--provider anthropic|openai|deepseek|mistral] [--model ID]
//                                         [--party "Customer"] [--playbook playbook.md] [--out folder]
import fs from 'node:fs';
import path from 'node:path';
import { loadEnv } from '../src/env.js';
import { reviewContract } from '../src/review.js';
import { PROVIDERS } from '../src/providers.js';

loadEnv();
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(`--${name}`); return i >= 0 ? args.splice(i, 2)[1] : undefined; };
const provider = opt('provider'), model = opt('model'), party = opt('party'), playbookFile = opt('playbook'), outDir = opt('out');
const file = args[0];

if (!file || file === '--help' || file === '-h') {
  console.log(`Open Contract Redline: AI contract review with a Word redline.

  node bin/cli.js contract.docx --provider anthropic --party "Customer"

  --provider   ${Object.keys(PROVIDERS).join(' | ')}  (default: AI_PROVIDER in .env, else anthropic)
  --model      model id (default per provider, see README)
  --party      who you act for, e.g. "Supplier"
  --playbook   a text/markdown file with your firm's positions
  --out        output folder (default: next to the contract)`);
  process.exit(file ? 0 : 1);
}

try {
  const buffer = fs.readFileSync(file);
  const playbook = playbookFile ? fs.readFileSync(playbookFile, 'utf8') : '';
  console.log(`Reviewing ${path.basename(file)} ...`);
  const r = await reviewContract({ buffer, filename: path.basename(file), provider, model, party, playbook });
  const dir = outDir || path.dirname(file);
  const base = path.basename(file).replace(/\.[^.]+$/, '');
  fs.mkdirSync(dir, { recursive: true });
  const docxOut = path.join(dir, `${base} - redline.docx`), mdOut = path.join(dir, `${base} - review.md`), htmlOut = path.join(dir, `${base} - redline.html`);
  fs.writeFileSync(docxOut, r.redline);
  fs.writeFileSync(mdOut, r.report);
  fs.writeFileSync(htmlOut, r.html);
  console.log(`Done with ${r.result.provider} / ${r.result.model}: ${r.result.issues.length} issues, ${r.applied.length} tracked changes.`);
  if (r.skipped.length) console.log(`${r.skipped.length} suggested change(s) could not be placed; see the review file.`);
  console.log(`Redline (Word):    ${docxOut}\nRedline (browser): ${htmlOut}\nReview:            ${mdOut}`);
} catch (e) {
  console.error('Error:', e.message);
  process.exit(1);
}
