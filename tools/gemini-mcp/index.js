#!/usr/bin/env node
/**
 * Gemini MCP server (stdio transport, zero dependencies).
 *
 * Claude Code から Gemini API を呼ぶための MCP サーバー。
 * 公開ツール:
 *   - ask_gemini      : Gemini に自由に質問する
 *   - gemini_research : Google 検索グラウンディング付きで調べる（出典付き）
 *
 * stdout は JSON-RPC 専用。ログは必ず stderr に出すこと。
 */
'use strict';

const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const SERVER_NAME = 'gemini';
const SERVER_VERSION = '1.0.0';
const FALLBACK_PROTOCOL = '2025-06-18';
const API_BASE = 'https://generativelanguage.googleapis.com/v1beta';
const REQUEST_TIMEOUT_MS = Number(process.env.GEMINI_TIMEOUT_MS || 120000);

const log = (...args) => process.stderr.write(`[gemini-mcp] ${args.join(' ')}\n`);

// ---------------------------------------------------------------------------
// .env の読み込み（リポジトリ直下の .env を探す。既存の環境変数を上書きしない）
// ---------------------------------------------------------------------------

function loadDotEnv() {
  let dir = __dirname;
  for (let i = 0; i < 5; i++) {
    const candidate = path.join(dir, '.env');
    if (fs.existsSync(candidate)) {
      for (const rawLine of fs.readFileSync(candidate, 'utf8').split('\n')) {
        const line = rawLine.trim();
        if (!line || line.startsWith('#')) continue;
        const eq = line.indexOf('=');
        if (eq === -1) continue;
        const key = line.slice(0, eq).trim();
        let value = line.slice(eq + 1).trim();
        if (
          (value.startsWith('"') && value.endsWith('"')) ||
          (value.startsWith("'") && value.endsWith("'"))
        ) {
          value = value.slice(1, -1);
        }
        if (key && process.env[key] === undefined) process.env[key] = value;
      }
      return candidate;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return null;
}

loadDotEnv();

const apiKey = () => process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || '';
const defaultModel = () => process.env.GEMINI_MODEL || 'gemini-2.5-pro';

// ---------------------------------------------------------------------------
// HTTP。プロキシ環境（Claude Code on the web など）では curl 経由にフォールバックする。
// Node の fetch は HTTPS_PROXY を見ないため。
// ---------------------------------------------------------------------------

const proxyUrl = () =>
  process.env.HTTPS_PROXY || process.env.https_proxy || process.env.ALL_PROXY || '';

function curlQuote(value) {
  return `"${String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

/** curl 経由の POST。API キーは argv に載せず、設定を stdin から渡す。 */
function postViaCurl(url, headers, body) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gemini-mcp-'));
  const bodyFile = path.join(dir, 'body.json');
  try {
    fs.writeFileSync(bodyFile, body, { mode: 0o600 });
    const config = [
      `url = ${curlQuote(url)}`,
      ...Object.entries(headers).map(([k, v]) => `header = ${curlQuote(`${k}: ${v}`)}`),
      `data-binary = ${curlQuote(`@${bodyFile}`)}`,
      'request = "POST"',
      'silent',
      'show-error',
      `max-time = ${Math.ceil(REQUEST_TIMEOUT_MS / 1000)}`,
      'write-out = "\\n%{http_code}"',
    ].join('\n');

    const result = spawnSync('curl', ['--config', '-'], {
      input: config,
      encoding: 'utf8',
      maxBuffer: 32 * 1024 * 1024,
    });

    if (result.error) throw new Error(`curl の起動に失敗しました: ${result.error.message}`);
    if (result.status !== 0) {
      throw new Error(`curl がエラー終了しました (exit ${result.status}): ${result.stderr.trim()}`);
    }

    const out = result.stdout;
    const cut = out.lastIndexOf('\n');
    const status = Number(out.slice(cut + 1).trim());
    return { status, text: out.slice(0, cut === -1 ? 0 : cut) };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

async function postViaFetch(url, headers, body) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body,
      signal: controller.signal,
    });
    return { status: response.status, text: await response.text() };
  } finally {
    clearTimeout(timer);
  }
}

async function postJson(url, headers, payload) {
  const body = JSON.stringify(payload);
  if (proxyUrl()) return postViaCurl(url, headers, body);
  return postViaFetch(url, headers, body);
}

// ---------------------------------------------------------------------------
// Gemini API
// ---------------------------------------------------------------------------

class ToolError extends Error {}

async function generateContent({ model, prompt, systemInstruction, temperature, useSearch }) {
  const key = apiKey();
  if (!key) {
    throw new ToolError(
      'GEMINI_API_KEY が設定されていません。\n' +
        'Google AI Studio (https://aistudio.google.com/apikey) でキーを取得し、\n' +
        'リポジトリ直下に .env を作って GEMINI_API_KEY=... を書くか、環境変数として設定してください。'
    );
  }

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature },
  };
  if (systemInstruction) payload.systemInstruction = { parts: [{ text: systemInstruction }] };
  if (useSearch) payload.tools = [{ google_search: {} }];

  const url = `${API_BASE}/models/${encodeURIComponent(model)}:generateContent`;
  const { status, text } = await postJson(
    url,
    { 'content-type': 'application/json', 'x-goog-api-key': key },
    payload
  );

  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new ToolError(`Gemini API の応答を解析できませんでした (HTTP ${status}): ${text.slice(0, 500)}`);
  }

  if (status < 200 || status >= 300) {
    const message = data?.error?.message || text.slice(0, 500);
    throw new ToolError(`Gemini API エラー (HTTP ${status}): ${message}`);
  }

  const candidate = data.candidates?.[0];
  if (!candidate) {
    const blockReason = data.promptFeedback?.blockReason;
    throw new ToolError(
      blockReason
        ? `Gemini が応答を返しませんでした (blockReason: ${blockReason})`
        : 'Gemini が応答を返しませんでした。'
    );
  }

  const answer = (candidate.content?.parts || [])
    .map((part) => part.text)
    .filter(Boolean)
    .join('\n')
    .trim();

  const sources = [];
  for (const chunk of candidate.groundingMetadata?.groundingChunks || []) {
    if (chunk.web?.uri) sources.push({ title: chunk.web.title || chunk.web.uri, uri: chunk.web.uri });
  }

  if (!answer) {
    throw new ToolError(`Gemini の応答が空でした (finishReason: ${candidate.finishReason || '不明'})`);
  }

  return { answer, sources, model, finishReason: candidate.finishReason };
}

// ---------------------------------------------------------------------------
// ツール定義
// ---------------------------------------------------------------------------

const TOOLS = [
  {
    name: 'ask_gemini',
    description:
      'Google Gemini に質問し、その回答を取得する。セカンドオピニオン、別視点からの検討、' +
      '長文の要約や下調べに使う。事実確認や最新情報が必要な場合は gemini_research を使うこと。',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: { type: 'string', description: 'Gemini に送る質問・指示。前提となる文脈も含めて具体的に書く。' },
        system_instruction: { type: 'string', description: 'Gemini に与える役割や出力形式の指定（任意）。' },
        model: { type: 'string', description: '使用するモデル。既定は gemini-2.5-pro。速度優先なら gemini-2.5-flash。' },
        temperature: { type: 'number', description: '0.0〜2.0。既定は 0.7。事実重視なら低め。' },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'gemini_research',
    description:
      'Google 検索グラウンディングを有効にした Gemini で調べ物をする。回答に加えて参照した URL を返す。' +
      '教育・研修コンテンツの下調べ、事実確認、最新情報の収集に使う。',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: '調べたい内容。トピックだけでなく、何を知りたいのかまで具体的に書く。' },
        system_instruction: { type: 'string', description: '出力形式や観点の指定（任意）。' },
        model: { type: 'string', description: '使用するモデル。既定は gemini-2.5-pro。' },
      },
      required: ['query'],
    },
  },
];

async function callTool(name, args) {
  if (name === 'ask_gemini') {
    const prompt = typeof args.prompt === 'string' ? args.prompt.trim() : '';
    if (!prompt) throw new ToolError('prompt は必須です。');
    const result = await generateContent({
      model: args.model || defaultModel(),
      prompt,
      systemInstruction: args.system_instruction,
      temperature: typeof args.temperature === 'number' ? args.temperature : 0.7,
      useSearch: false,
    });
    return `${result.answer}\n\n---\n(model: ${result.model})`;
  }

  if (name === 'gemini_research') {
    const query = typeof args.query === 'string' ? args.query.trim() : '';
    if (!query) throw new ToolError('query は必須です。');
    const result = await generateContent({
      model: args.model || defaultModel(),
      prompt: query,
      systemInstruction: args.system_instruction,
      temperature: 0.3,
      useSearch: true,
    });
    const sources = result.sources.length
      ? result.sources.map((s, i) => `${i + 1}. ${s.title} — ${s.uri}`).join('\n')
      : '(グラウンディングされた出典はありません。内容の裏取りが必要です)';
    return `${result.answer}\n\n## 出典\n${sources}\n\n---\n(model: ${result.model})`;
  }

  throw new ToolError(`未知のツールです: ${name}`);
}

// ---------------------------------------------------------------------------
// JSON-RPC / MCP
// ---------------------------------------------------------------------------

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function sendResult(id, result) {
  send({ jsonrpc: '2.0', id, result });
}

function sendError(id, code, message) {
  send({ jsonrpc: '2.0', id, error: { code, message } });
}

async function handleMessage(message) {
  const { id, method, params } = message;
  const isNotification = id === undefined || id === null;

  switch (method) {
    case 'initialize': {
      const requested = params?.protocolVersion;
      sendResult(id, {
        protocolVersion: typeof requested === 'string' ? requested : FALLBACK_PROTOCOL,
        capabilities: { tools: {} },
        serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
      });
      return;
    }

    case 'notifications/initialized':
    case 'notifications/cancelled':
      return;

    case 'ping':
      if (!isNotification) sendResult(id, {});
      return;

    case 'tools/list':
      sendResult(id, { tools: TOOLS });
      return;

    case 'tools/call': {
      const name = params?.name;
      const args = params?.arguments || {};
      try {
        const text = await callTool(name, args);
        sendResult(id, { content: [{ type: 'text', text }], isError: false });
      } catch (error) {
        const text =
          error instanceof ToolError
            ? error.message
            : `Gemini の呼び出しに失敗しました: ${error?.message || String(error)}`;
        log(`tools/call failed (${name}): ${error?.message || error}`);
        sendResult(id, { content: [{ type: 'text', text }], isError: true });
      }
      return;
    }

    default:
      if (!isNotification) sendError(id, -32601, `Method not found: ${method}`);
  }
}

function main() {
  let buffer = '';
  let queue = Promise.resolve();

  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (chunk) => {
    buffer += chunk;
    let index;
    while ((index = buffer.indexOf('\n')) !== -1) {
      const line = buffer.slice(0, index).trim();
      buffer = buffer.slice(index + 1);
      if (!line) continue;

      let message;
      try {
        message = JSON.parse(line);
      } catch {
        log(`JSON の解析に失敗しました: ${line.slice(0, 200)}`);
        continue;
      }
      // メッセージは受信順に処理する
      queue = queue.then(() => handleMessage(message)).catch((error) => {
        log(`ハンドラでエラーが発生しました: ${error?.message || error}`);
      });
    }
  });

  process.stdin.on('end', () => process.exit(0));
  log(`起動しました (model: ${defaultModel()}, API key: ${apiKey() ? 'あり' : 'なし'})`);
}

main();
