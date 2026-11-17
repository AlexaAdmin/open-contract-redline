# Notes for coding agents (Claude Code, Codex)

Open Contract Redline reviews a contract with ONE AI model and writes the result back into the .docx as Word tracked changes with comments.

## Layout
- `src/prompt.js`: review instructions and the JSON schema the model must return. Change review behaviour here.
- `src/providers.js`: one function per provider (Anthropic, OpenAI, DeepSeek, Mistral). Each makes exactly one model call per review.
- `src/docx.js`: reads paragraphs from `word/document.xml` and writes tracked changes (`w:del`/`w:ins`) plus comments (`word/comments.xml`). Untouched paragraphs must stay byte-for-byte identical.
- `src/preview.js`: the on-screen redline (one self-contained HTML page), built from the finished .docx via `previewDocx` in `src/docx.js`, so it always matches Word. Escape every piece of text that goes into it.
- `src/review.js`: ties it together, builds the markdown report, and runs the no-key demo (`examples/demo-review.json`).
- `bin/cli.js`: command line. `server.js` + `public/index.html`: local web page on 127.0.0.1.
- `test/redline.test.js`: tests with a fake model answer (no API key needed).

## Rules
- Run `npm test` after every change. Add a test for any new behaviour in `src/docx.js`.
- Keep a review to a single model call unless the user explicitly asks for more.
- Never place an edit that would corrupt the document. If an edit cannot be placed safely, skip it and report it (see `skipped` in `redlineDocx`).
- Never log, store or send contract text or API keys anywhere except the chosen provider.
- Keep dependencies few. Plain JavaScript (ES modules), Node 20+.
- Output is a draft for a lawyer to check; keep the "not legal advice" notice in the UI and report.
