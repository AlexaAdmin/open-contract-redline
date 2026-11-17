// Local web page: npm start, then open http://localhost:8787
// Runs only on your computer. Your contract goes straight from here to the AI provider you pick.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadEnv } from './src/env.js';
import { reviewContract, demoReview } from './src/review.js';
import { PROVIDERS } from './src/providers.js';

loadEnv();
const here = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 8787);
const MAX_BYTES = 25 * 1024 * 1024;

const send = (res, code, body, type = 'application/json') => {
  res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
};

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
      return send(res, 200, fs.readFileSync(path.join(here, 'public', 'index.html')), 'text/html; charset=utf-8');
    }
    if (req.method === 'GET' && req.url === '/api/providers') {
      return send(res, 200, Object.entries(PROVIDERS).map(([id, p]) => ({
        id, label: p.label, model: process.env[`${id.toUpperCase()}_MODEL`] || p.model, hasKey: !!process.env[p.env],
      })));
    }
    if (req.method === 'POST' && req.url === '/api/demo') {
      const r = await demoReview();
      return send(res, 200, { result: r.result, report: r.report, applied: r.applied.length, skipped: r.skipped, redline: r.redline.toString('base64'), html: r.html, filename: 'sample-services-agreement.docx' });
    }
    if (req.method === 'POST' && req.url === '/api/review') {
      const chunks = []; let size = 0;
      for await (const c of req) { size += c.length; if (size > MAX_BYTES) return send(res, 413, { error: 'File too large (25 MB max).' }); chunks.push(c); }
      const b = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      const r = await reviewContract({
        buffer: Buffer.from(b.file, 'base64'), filename: b.filename,
        provider: b.provider, model: b.model || undefined, apiKey: b.apiKey || undefined,
        party: b.party, playbook: b.playbook,
      });
      return send(res, 200, { result: r.result, report: r.report, applied: r.applied.length, skipped: r.skipped, redline: r.redline.toString('base64'), html: r.html });
    }
    send(res, 404, { error: 'Not found' });
  } catch (e) {
    send(res, 400, { error: e.message });
  }
});

server.listen(PORT, '127.0.0.1', () => console.log(`Open Contract Redline is running: http://localhost:${PORT}`));
