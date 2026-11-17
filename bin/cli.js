#!/usr/bin/env node
// Usage:
//   With an API key:  node bin/cli.js <contract.docx> [--provider anthropic|openai|deepseek|mistral] [--model ID]
//                                     [--party "Customer"] [--playbook playbook.md] [--out folder] [--open]
//   Agent mode (no API key, for Claude Code / Codex):
//     1. node bin/cli.js <contract.docx> --prepare [--party "Customer"] [--playbook playbook.md]
//     2. the AI agent reads "<name> - agent-brief.md" and writes "<name> - review.json"
//     3. node bin/cli.js <contract.docx> --from "<name> - review.json" --open
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { loadEnv } from '../src/env.js';
import { reviewContract, contractParagraphs } from '../src/review.js';
import { agentBrief, readAgentReview } from '../src/agent.js';
import { PROVIDERS } from '../src/providers.js';

loadEnv();
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(`--${name}`); return i >= 0 ? args.splice(i, 2)[1] : undefined; };
const flag = (name) => { const i = args.indexOf(`--${name}`); if (i < 0) return false; args.splice(i, 1); return true; };
const provider = opt('provider'), model = opt('model'), party = opt('party'), playbookFile = opt('playbook'), outDir = opt('out');
const from = opt('from'), agentName = opt('agent');
const prepare = flag('prepare'), open = flag('open');
const file = args[0];

if (!file || file === '--help' || file === '-h') {
  console.log(`Open Contract Redline: AI contract review with a redline in your browser and in Word.

  With an API key:
    node bin/cli.js contract.docx --provider anthropic --party "Customer" --open

  Agent mode, no API key (run by Claude Code, Codex or another AI agent):
    node bin/cli.js contract.docx --prepare --party "Customer"     writes "<name> - agent-brief.md"
    (the agent reads the brief and writes "<name> - review.json")
    node bin/cli.js contract.docx --from "<name> - review.json" --open

  --provider   ${Object.keys(PROVIDERS).join(' | ')}  (default: AI_PROVIDER in .env, else anthropic)
  --model      model id (default per provider, see README)
  --party      who you act for, e.g. "Supplier"
  --playbook   a text/markdown file with your firm's positions
  --out        output folder (default: next to the contract)
  --open       open the browser redline when done`);
  process.exit(file ? 0 : 1);
}

function openFile(p) {
  const cmd = process.platform === 'win32' ? 'explorer.exe' : process.platform === 'darwin' ? 'open' : 'xdg-open';
  try { spawn(cmd, [p], { detached: true, stdio: 'ignore' }).on('error', () => {}).unref(); } catch {}
  console.log(`Opening the redline in your browser. If nothing opens, open this file: ${p}`);
}

try {
  const buffer = fs.readFileSync(file);
  const playbook = playbookFile ? fs.readFileSync(playbookFile, 'utf8') : '';
  const dir = outDir || path.dirname(file);
  const base = path.basename(file).replace(/\.[^.]+$/, '');
  fs.mkdirSync(dir, { recursive: true });

  if (prepare) {
    const paragraphs = await contractParagraphs({ buffer, filename: path.basename(file) });
    const reviewFile = path.join(dir, `${base} - review.json`);
    const briefOut = path.join(dir, `${base} - agent-brief.md`);
    fs.writeFileSync(briefOut, agentBrief({ paragraphs, filename: path.basename(file), party, playbook, reviewFile, contractPath: file }));
    console.log(`Agent brief written (${paragraphs.length} paragraphs): ${briefOut}
Next: read the brief, write your review as JSON to "${reviewFile}", then run:
  node bin/cli.js "${file}" --from "${reviewFile}" --open`);
    process.exit(0);
  }

  let review;
  if (from) review = readAgentReview(fs.readFileSync(from, 'utf8'));
  else console.log(`Reviewing ${path.basename(file)} ...`);
  const r = await reviewContract({ buffer, filename: path.basename(file), provider, model, party, playbook, review, agentName });
  const docxOut = path.join(dir, `${base} - redline.docx`), mdOut = path.join(dir, `${base} - review.md`), htmlOut = path.join(dir, `${base} - redline.html`);
  fs.writeFileSync(docxOut, r.redline);
  fs.writeFileSync(mdOut, r.report);
  fs.writeFileSync(htmlOut, r.html);
  console.log(`Done with ${r.result.provider} / ${r.result.model}: ${r.result.issues.length} issues, ${r.applied.length} tracked changes.`);
  if (r.skipped.length) {
    console.log(`${r.skipped.length} suggested change(s) could not be placed:`);
    for (const s of r.skipped) console.log(`  - [P${s.paragraph}] "${s.find}" (${s.reason})`);
    if (from) console.log(`Fix their "find" text in ${from} (copy it exactly from the paragraph) and run again.`);
  }
  console.log(`Redline (browser): ${htmlOut}\nRedline (Word):    ${docxOut}\nReview:            ${mdOut}`);
  if (open) openFile(path.resolve(htmlOut));
} catch (e) {
  console.error('Error:', e.message);
  process.exit(1);
}
