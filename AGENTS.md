# Notes for AI agents (Claude Code, Codex)

Open Contract Redline reviews a contract and shows the redline in the browser, then writes it back into the .docx as Word tracked changes with comments.

## If the user asks you to install it, run it, or redline a contract

Use **agent mode** unless the user has an API key in `.env` and asks to use it. In agent mode **you** do the review, so the user needs no API key.

1. Install once: `npm install` (Node.js 20 or later).
2. Find the contract. If the user did not name a file, ask for it. If they just want to see it work, use `examples/sample-services-agreement.docx`. Ask who they act for (e.g. "Customer") if it is not clear; if they do not know, leave it out.
3. Prepare the brief:
   `node bin/cli.js "<contract>" --prepare --party "<party>"`
   (add `--playbook <file>` if the user gives you their positions). This writes `<name> - agent-brief.md` next to the contract.
4. Read the whole brief. Review the contract yourself, following its instructions exactly, and write ONLY the JSON answer to `<name> - review.json` (the path is in the brief). Copy every `find` text exactly from its paragraph.
5. Build the redline and open it:
   `node bin/cli.js "<contract>" --from "<name> - review.json" --open`
   If it lists changes that could not be placed, correct their `find` text in the JSON and run step 5 again.
6. Tell the user where the files are: `<name> - redline.html` (opens in the browser), `<name> - redline.docx` (Word, real tracked changes) and `<name> - review.md`. Remind them it is a first draft for a lawyer to check.

To show the local web page instead (for users with an API key, or to try the built-in demo): `npm start`, then open http://localhost:8787.

## Layout
- `src/prompt.js`: review instructions and the JSON schema. Used by the API models and by the agent brief, so both review the same way. Change review behaviour here.
- `src/agent.js`: agent mode. Writes the brief and reads the agent's JSON answer.
- `src/providers.js`: one function per provider (Anthropic, OpenAI, DeepSeek, Mistral). Each makes exactly one model call per review.
- `src/docx.js`: reads paragraphs from `word/document.xml` and writes tracked changes (`w:del`/`w:ins`) plus comments (`word/comments.xml`). Untouched paragraphs must stay byte-for-byte identical.
- `src/preview.js`: the on-screen redline (one self-contained HTML page), built from the finished .docx via `previewDocx` in `src/docx.js`, so it always matches Word. Escape every piece of text that goes into it.
- `src/review.js`: ties it together, builds the markdown report, and runs the no-key demo (`examples/demo-review.json`).
- `bin/cli.js`: command line (API mode and agent mode). `server.js` + `public/index.html`: local web page on 127.0.0.1.
- `test/redline.test.js`: tests with a fixed answer (no API key needed).

## Rules
- Run `npm test` after every change to the code. Add a test for any new behaviour in `src/docx.js`.
- Keep a review to a single model call (API mode) or a single review pass (agent mode) unless the user explicitly asks for more.
- Never place an edit that would corrupt the document. If an edit cannot be placed safely, skip it and report it (see `skipped` in `redlineDocx`).
- Never log, store or send contract text or API keys anywhere except the chosen provider. Never change the user's original contract file.
- Keep dependencies few. Plain JavaScript (ES modules), Node 20+.
- Output is a draft for a lawyer to check; keep the "not legal advice" notice in the UI and report.
