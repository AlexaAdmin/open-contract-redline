// One review = one call to ONE AI model. Pick the provider and model; bring your own API key.
// Supported: anthropic (Claude), openai, deepseek, mistral.
import Anthropic from '@anthropic-ai/sdk';
import OpenAI from 'openai';
import { Mistral } from '@mistralai/mistralai';
import { RESULT_SCHEMA, systemPrompt, userPrompt } from './prompt.js';

export const PROVIDERS = {
  anthropic: { label: 'Anthropic (Claude)', env: 'ANTHROPIC_API_KEY', model: 'claude-opus-5-5' },
  openai: { label: 'OpenAI', env: 'OPENAI_API_KEY', model: 'gpt-6-astra' },
  deepseek: { label: 'DeepSeek', env: 'DEEPSEEK_API_KEY', model: 'deepseek-v4-pro' },
  mistral: { label: 'Mistral', env: 'MISTRAL_API_KEY', model: 'mistral-large-latest' },
};

export function resolveConfig({ provider, model, apiKey } = {}) {
  const name = String(provider || process.env.AI_PROVIDER || 'anthropic').toLowerCase();
  const p = PROVIDERS[name];
  if (!p) throw new Error(`Unknown provider "${name}". Use one of: ${Object.keys(PROVIDERS).join(', ')}.`);
  const key = apiKey || process.env[p.env];
  if (!key) throw new Error(`No API key for ${p.label}. Set ${p.env} in your .env file.`);
  const envModel = process.env[`${name.toUpperCase()}_MODEL`];
  return { provider: name, model: model || envModel || p.model, apiKey: key };
}

// Models sometimes wrap JSON in ```json fences or add a sentence; take the outermost object.
export function parseJson(text) {
  const s = String(text || '').trim();
  try { return JSON.parse(s); } catch {}
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) { try { return JSON.parse(s.slice(a, b + 1)); } catch {} }
  throw new Error('The AI model did not return valid JSON. Try again, or try another model.');
}

const MAX_OUT = 32000;

async function callAnthropic(cfg, system, user) {
  const client = new Anthropic({ apiKey: cfg.apiKey });
  // Streaming avoids HTTP timeouts on long contracts; finalMessage() returns the complete reply.
  const stream = client.messages.stream({
    model: cfg.model,
    max_tokens: MAX_OUT,
    system,
    messages: [{ role: 'user', content: user }],
    output_config: { effort: 'high', format: { type: 'json_schema', schema: RESULT_SCHEMA } },
  });
  const msg = await stream.finalMessage();
  if (msg.stop_reason === 'refusal') throw new Error('The model declined to review this document.');
  if (msg.stop_reason === 'max_tokens') throw new Error('The answer was cut off: the contract is too long for one pass. Split it and try again.');
  return msg.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
}

async function callOpenAICompatible(cfg, system, user, baseURL) {
  const client = new OpenAI({ apiKey: cfg.apiKey, ...(baseURL ? { baseURL } : {}) });
  const strict = cfg.provider === 'openai';
  const res = await client.chat.completions.create({
    model: cfg.model,
    messages: [{ role: 'system', content: system }, { role: 'user', content: user }],
    response_format: strict
      ? { type: 'json_schema', json_schema: { name: 'contract_review', strict: true, schema: RESULT_SCHEMA } }
      : { type: 'json_object' },
    ...(strict ? { max_completion_tokens: MAX_OUT } : { max_tokens: 8192 }),
  });
  const choice = res.choices?.[0];
  if (choice?.finish_reason === 'length') throw new Error('The answer was cut off: the contract is too long for one pass. Split it and try again.');
  return choice?.message?.content || '';
}

async function callMistral(cfg, system, user) {
  const client = new Mistral({ apiKey: cfg.apiKey });
  const res = await client.chat.complete({
    model: cfg.model,
    messages: [{ role: 'system', content: system }, { role: 'user', content: user }],
    responseFormat: { type: 'json_object' },
  });
  const c = res.choices?.[0]?.message?.content;
  return Array.isArray(c) ? c.map((x) => x.text || '').join('') : (c || '');
}

/** Ask ONE model to review the contract. Returns { summary, overall_risk, issues[], edits[] }. */
export async function reviewWithModel(paragraphs, opts = {}) {
  const cfg = resolveConfig(opts);
  const system = systemPrompt({ party: opts.party, playbook: opts.playbook });
  const user = userPrompt(paragraphs);
  const call = opts._call || ({
    anthropic: () => callAnthropic(cfg, system, user),
    openai: () => callOpenAICompatible(cfg, system, user),
    deepseek: () => callOpenAICompatible(cfg, system, user, 'https://api.deepseek.com'),
    mistral: () => callMistral(cfg, system, user),
  })[cfg.provider];
  const text = await call(cfg, system, user);
  const out = parseJson(text);
  return {
    summary: String(out.summary || ''),
    overall_risk: ['low', 'medium', 'high'].includes(out.overall_risk) ? out.overall_risk : 'medium',
    issues: Array.isArray(out.issues) ? out.issues : [],
    edits: Array.isArray(out.edits) ? out.edits : [],
    provider: cfg.provider,
    model: cfg.model,
  };
}
