// Agent mode: no API key. An AI coding agent (Claude Code, Codex, ...) does the review itself,
// using the subscription it already runs on. This file writes the brief the agent follows, and
// reads its answer back. The redline is then built exactly as for an API review.
import { RESULT_SCHEMA, systemPrompt } from './prompt.js';
import { parseJson, normalizeResult } from './providers.js';

/** The brief an AI agent reads: the same instructions and JSON shape the API models get, plus the contract. */
export function agentBrief({ paragraphs, filename, party, playbook, reviewFile, contractPath }) {
  const body = paragraphs.map((p, i) => `[P${i}] ${p}`).join('\n');
  return `# Agent brief: review "${filename}"

You are an AI agent doing a contract review for Open Contract Redline, in agent mode (no API key).
Do the review yourself, following the instructions below exactly, then save your answer as JSON.

## 1. Instructions

${systemPrompt({ party, playbook })}

## 2. Save your answer

Write ONLY the JSON object (no other text, no code fences) to this file:

    ${reviewFile}

It must match this JSON schema:

\`\`\`json
${JSON.stringify(RESULT_SCHEMA, null, 2)}
\`\`\`

Checklist before saving:
- Every "find" is copied character for character from that one paragraph (curly quotes included).
- No edit spans two paragraphs, and no two edits overlap.
- Every edit has a "comment" and the "issue_id" of an issue in "issues".

## 3. Build the redline

Then run:

    node bin/cli.js "${contractPath}" --from "${reviewFile}" --open

It writes the Word redline, the browser redline and the review report, and opens the redline in the browser.
If it reports changes that could not be placed, fix their "find" text in the JSON and run it again.

## 4. The contract

<contract>
${body}
</contract>
`;
}

/** Read the agent's JSON answer. */
export function readAgentReview(text) {
  const out = normalizeResult(parseJson(text));
  if (!out.summary && !out.issues.length && !out.edits.length) throw new Error('The review file is empty. It needs summary, overall_risk, issues and edits.');
  return out;
}
