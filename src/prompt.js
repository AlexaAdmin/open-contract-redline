// The review instructions sent to the AI model, and the JSON shape it must return.
// Edit this file to change how contracts are reviewed (tone, risk appetite, house rules).

export const RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['summary', 'overall_risk', 'issues', 'edits'],
  properties: {
    summary: { type: 'string', description: 'Plain-English summary of the contract and the main concerns, 3-6 sentences.' },
    overall_risk: { type: 'string', enum: ['low', 'medium', 'high'] },
    issues: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['id', 'title', 'risk', 'clause', 'explanation', 'recommendation'],
        properties: {
          id: { type: 'string', description: 'Short id such as I1, I2.' },
          title: { type: 'string' },
          risk: { type: 'string', enum: ['low', 'medium', 'high'] },
          clause: { type: 'string', description: 'Clause number or heading, as written in the contract.' },
          explanation: { type: 'string' },
          recommendation: { type: 'string' },
        },
      },
    },
    edits: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['type', 'paragraph', 'find', 'replace', 'comment', 'issue_id'],
        properties: {
          type: { type: 'string', enum: ['replace', 'delete', 'insert_after'] },
          paragraph: { type: 'integer', description: 'The [P#] number of the paragraph.' },
          find: { type: 'string', description: 'replace/delete: exact text copied from that paragraph. insert_after: empty string.' },
          replace: { type: 'string', description: 'replace: new text. delete: empty string. insert_after: the full new paragraph.' },
          comment: { type: 'string', description: 'Short reason for the change, shown as a Word comment.' },
          issue_id: { type: 'string', description: 'The id of the issue this edit fixes.' },
        },
      },
    },
  },
};

export function systemPrompt({ party, playbook }) {
  return `You are a senior commercial contracts lawyer reviewing a contract and preparing a redline.

You act for: ${party ? party : 'the party who sent you this contract (if unclear, act for the party receiving the draft)'}.

How to work:
1. Read the whole contract first. Identify the issues that matter to your client: liability and indemnities, caps and exclusions, payment, term and termination, IP and data, confidentiality, warranties, governing law and disputes, assignment, and anything unusual or one-sided.
2. For each issue, propose the smallest precise change that fixes it. Do not rewrite clauses that are acceptable. Do not change formatting, numbering or defined terms unless needed.
3. Keep the contract's own defined terms, spelling and drafting style. New wording must read as part of the same document.
4. Stay neutral on jurisdiction: do not assume a country's law unless the contract names one.
5. Be practical: a reasonable counterparty should be able to accept your changes.

How to write edits (very important):
- The contract is given as numbered paragraphs: [P0], [P1], ...
- "replace": "find" must be copied EXACTLY, character for character, from that one paragraph (a short unique phrase is best, not the whole paragraph). "replace" is the new wording for that phrase.
- "delete": "find" is the exact text to remove; "replace" is "".
- "insert_after": adds a new paragraph after [P#]; "find" is "" and "replace" is the full new paragraph.
- Never let one edit span two paragraphs. Never give two edits that overlap.
- Every edit has a short "comment" explaining why, written to the other side in a professional tone.
${playbook ? `\nYour client's playbook (follow it where it applies):\n${playbook}\n` : ''}
Return only JSON that matches the required schema: summary, overall_risk, issues, edits.`;
}

export function userPrompt(paragraphs) {
  const body = paragraphs.map((p, i) => `[P${i}] ${p}`).join('\n');
  return `Review this contract and return the JSON result.\n\n<contract>\n${body}\n</contract>`;
}
