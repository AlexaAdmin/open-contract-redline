# Open Contract Redline

**Free, open-source AI contract review that gives you a real Word redline.**

Upload a contract. One AI model reads it, lists the issues by risk, and returns your `.docx` with **tracked changes and margin comments**, ready to open in Microsoft Word and send back to the other side.

- Bring your own API key: **Anthropic (Claude), OpenAI, DeepSeek or Mistral**
- One model per review, the one you choose. No hidden second model.
- Runs on your own computer. Your contract goes straight to the AI provider you pick, and nowhere else.
- Works as a local web page or from the command line
- Plain JavaScript, small codebase: built to be changed with **Claude Code** or **Codex**

> AI output is a first draft for a qualified lawyer to check. It is not legal advice.

---

## Quick start (5 minutes)

You need [Node.js 20 or later](https://nodejs.org) and an API key from one provider.

```bash
git clone https://github.com/AlexaAdmin/open-contract-redline.git
cd open-contract-redline
npm install
cp .env.example .env      # then paste your API key into .env
npm start                 # open http://localhost:8787
```

Or from the command line:

```bash
node bin/cli.js "My Contract.docx" --provider anthropic --party "Customer"
```

You get two files next to the contract:

| File | What it is |
|---|---|
| `My Contract - redline.docx` | Your contract with tracked changes and a comment explaining each change |
| `My Contract - review.md` | Summary, overall risk, and every issue with a recommendation |

Try it first on the fictional contract in `examples/sample-services-agreement.docx`.

## Choosing a provider and model

| Provider | `--provider` | Key in `.env` | Default model |
|---|---|---|---|
| Anthropic (Claude) | `anthropic` | `ANTHROPIC_API_KEY` | `claude-opus-5-5` |
| OpenAI | `openai` | `OPENAI_API_KEY` | `gpt-6-astra` |
| DeepSeek | `deepseek` | `DEEPSEEK_API_KEY` | `deepseek-v4-pro` |
| Mistral | `mistral` | `MISTRAL_API_KEY` | `mistral-large-latest` |

Use any other model your key can access with `--model <id>`, or set `ANTHROPIC_MODEL`, `OPENAI_MODEL`, `DEEPSEEK_MODEL` or `MISTRAL_MODEL` in `.env`. Providers rename models from time to time: if a default stops working, check the provider's model list and set the new id.

## Options

| Option | Meaning |
|---|---|
| `--party "Supplier"` | Who you act for. The review protects this party. |
| `--playbook positions.md` | Your firm's positions (liability caps, governing law, red lines). The model follows them where they apply. |
| `--model <id>` | Use a specific model. |
| `--out folder` | Where to save the two files. |

The web page has the same options.

## How it works

1. `src/docx.js` reads the Word file into numbered paragraphs (including table cells).
2. `src/providers.js` sends them, with the instructions in `src/prompt.js`, to **one** model and asks for JSON: summary, issues and precise edits.
3. `src/docx.js` writes each edit back as a Word tracked change (deletion + insertion) with a comment. Paragraphs the model did not touch are left exactly as they were, formatting included.
4. If an edit cannot be placed safely (for example, inside a hyperlink or a field), it is not forced into the document. It is listed in the review file so you can apply it by hand.

Supported files: `.docx`, `.txt`, `.md`. For a PDF, save it as Word first.

## Your data

- Nothing is stored by this tool. The local web page listens on `127.0.0.1` only.
- Your contract is sent to the AI provider you choose, under that provider's API terms. Check their data retention and training settings, and use an account your firm has approved for client documents.
- Keep your `.env` file private. It is listed in `.gitignore`.

## Make it yours with Claude Code or Codex

Open the folder in Claude Code or Codex and ask in plain words, for example:

- "Add a clause library: when the model suggests a new clause, use our standard wording from `clauses/`."
- "Add a second pass that checks every cross-reference in the redline."
- "Export the issue list to Excel as well."
- "Make the comments come from 'Jane Smith, Legal' instead of 'AI Review'."
- "Support PDF uploads."

`AGENTS.md` (Codex) and `CLAUDE.md` (Claude Code) tell the coding agent how the project is laid out and how to test a change. Run `npm test` after every change: the tests check the redline without needing an API key.

## Want the full workspace?

This repository does one job: review and redline a single contract. **[ContractApp](https://www.ai2bundle.co/contractapp-enterprise)** is the full AI workspace for law firms and in-house teams, with matters, firm playbooks, team approvals, due diligence, drafting, research and secure firm-wide storage.

## Licence

[Apache License 2.0](LICENSE). You may use, change and share this code, including commercially, as long as you keep the licence and the notices in [NOTICE](NOTICE). The licence does not cover our names or logos.

Contributions are welcome: open an issue or a pull request.
