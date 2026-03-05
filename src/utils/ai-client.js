/**
 * ai-client.js – LM Studio / Ollama client
 *
 * Sends messages to any OpenAI-compatible endpoint (LM Studio or Ollama).
 * All requests are made entirely from the browser to localhost – no data
 * leaves the user's machine.
 */

"use strict";

/* ─────────────────────────────────────────────
   Default configuration (overridden from settings)
───────────────────────────────────────────── */

export const DEFAULT_CONFIG = {
  baseUrl: "http://localhost:1234/v1",   // LM Studio default
  model: null,                            // null = auto-detect
  maxTokens: 2048,
  temperature: 0.7,
  stream: true,
};

/* ─────────────────────────────────────────────
   Config loader / saver (localStorage)
───────────────────────────────────────────── */

const STORAGE_KEY = "excel_ai_config";

export function loadConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? { ...DEFAULT_CONFIG, ...JSON.parse(raw) } : { ...DEFAULT_CONFIG };
  } catch {
    return { ...DEFAULT_CONFIG };
  }
}

export function saveConfig(cfg) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
}

/* ─────────────────────────────────────────────
   Model detection
───────────────────────────────────────────── */

/**
 * Fetch the first available model from the /models endpoint.
 * Returns model id string or null on failure.
 */
export async function detectModel(baseUrl) {
  try {
    const resp = await fetch(`${baseUrl}/models`, {
      headers: { "Content-Type": "application/json" },
      signal: AbortSignal.timeout(5000),
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    const models = data.data || data.models || [];
    return models.length > 0 ? (models[0].id || models[0].name || null) : null;
  } catch {
    return null;
  }
}

/* ─────────────────────────────────────────────
   Chat completion (streaming)
───────────────────────────────────────────── */

/**
 * Send a chat completion request.
 *
 * @param {Array}    messages        OpenAI-format messages array
 * @param {object}   cfg             Config (from loadConfig())
 * @param {function} onChunk         Called with each text chunk as it streams in
 * @param {function} onDone          Called when streaming is complete with full text
 * @param {function} onError         Called on error with Error object
 * @param {AbortSignal} abortSignal  Optional AbortSignal
 */
export async function streamChat({
  messages,
  cfg,
  onChunk = () => {},
  onDone = () => {},
  onError = () => {},
  abortSignal = null,
}) {
  const config = cfg || loadConfig();
  const model = config.model || (await detectModel(config.baseUrl));

  const body = {
    model: model || "local-model",
    messages,
    max_tokens: config.maxTokens,
    temperature: config.temperature,
    stream: true,
  };

  try {
    const resp = await fetch(`${config.baseUrl}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: abortSignal,
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`AI backend returned ${resp.status}: ${errText}`);
    }

    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let fullText = "";
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop(); // keep incomplete last line

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed === "data: [DONE]") continue;
        if (!trimmed.startsWith("data: ")) continue;

        try {
          const json = JSON.parse(trimmed.slice(6));
          const chunk =
            json.choices?.[0]?.delta?.content ||
            json.choices?.[0]?.text ||
            "";
          if (chunk) {
            fullText += chunk;
            onChunk(chunk, fullText);
          }
        } catch {
          // Ignore malformed SSE lines
        }
      }
    }

    onDone(fullText);
    return fullText;
  } catch (err) {
    if (err.name !== "AbortError") {
      onError(err);
    }
    return null;
  }
}

/**
 * Non-streaming completion (for background tasks).
 */
export async function chatOnce(messages, cfg) {
  const config = cfg || loadConfig();
  const model = config.model || (await detectModel(config.baseUrl));

  const resp = await fetch(`${config.baseUrl}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: model || "local-model",
      messages,
      max_tokens: config.maxTokens,
      temperature: config.temperature,
      stream: false,
    }),
    signal: AbortSignal.timeout(60_000),
  });

  if (!resp.ok) throw new Error(`AI backend error: ${resp.status}`);
  const data = await resp.json();
  return (
    data.choices?.[0]?.message?.content ||
    data.choices?.[0]?.text ||
    ""
  );
}

/* ─────────────────────────────────────────────
   System prompt builder
───────────────────────────────────────────── */

export function buildSystemPrompt(dataContext = "") {
  return `You are an expert Excel data analyst and AI assistant integrated directly inside Microsoft Excel.

Your capabilities:
- Analyse tabular data and provide statistical insights
- Suggest and explain appropriate chart types for different data
- Detect anomalies, outliers, and trends in datasets
- Write Excel formulas and VBA macros when asked
- Answer questions about data science, statistics, and Excel

Rules:
- Be concise but thorough
- Format numbers clearly (e.g., use commas for thousands)
- When suggesting charts, explain WHY that chart type is appropriate
- When detecting anomalies, specify the row/column and explain why it is anomalous
- Never request or suggest sending data outside this machine (privacy-first)
${dataContext ? `\nCurrent Excel data context:\n${dataContext}` : ""}`;
}
