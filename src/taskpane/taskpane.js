/**
 * taskpane.js – Main Office Add-in task-pane controller
 *
 * Orchestrates the chat UI, Excel API, AI client, data analysis, smart
 * charting, and anomaly / trend detection.  All processing is local –
 * no data leaves the user's machine.
 */

"use strict";

import {
  readSelection,
  readUsedRange,
  highlightCells,
  writeValues,
  summariseData,
  toA1,
  colIndexToLetter,
} from "../utils/excel-api.js";

import {
  loadConfig,
  saveConfig,
  detectModel,
  streamChat,
  buildSystemPrompt,
} from "../utils/ai-client.js";

import {
  analyseDataset,
  detectAnomalies,
  detectTrends,
  formatAnalysisReport,
  formatAnomalyReport,
  formatTrendReport,
} from "../utils/data-analyzer.js";

import { autoChart, insertChart, recommendChartType } from "../utils/chart-manager.js";

/* ─────────────────────────────────────────────
   State
───────────────────────────────────────────── */
let cfg           = loadConfig();
let chatHistory   = [];          // OpenAI-format messages
let isGenerating  = false;
let abortCtrl     = null;
let mediaRecorder = null;
let audioChunks   = [];
let videoStream   = null;
let currentModel  = null;

/* ─────────────────────────────────────────────
   Office.js bootstrap
───────────────────────────────────────────── */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initUI();
    checkAIConnection();
  }
});

/* ─────────────────────────────────────────────
   UI Initialisation
───────────────────────────────────────────── */
function initUI() {
  // ---- Tabs ----
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
    });
  });

  // ---- Chat send ----
  document.getElementById("sendBtn").addEventListener("click", sendMessage);
  document.getElementById("chatInput").addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") sendMessage();
  });

  // ---- Settings ----
  document.getElementById("settingsBtn").addEventListener("click", openSettings);
  document.getElementById("closeSettings").addEventListener("click", closeSettings);
  document.getElementById("saveSettings").addEventListener("click", saveSettingsUI);
  document.getElementById("testConnection").addEventListener("click", testConnectionUI);
  document.getElementById("temperature").addEventListener("input", (e) => {
    document.getElementById("tempDisplay").textContent = parseFloat(e.target.value).toFixed(1);
  });

  // ---- Clear chat ----
  document.getElementById("clearChatBtn").addEventListener("click", clearChat);

  // ---- Voice ----
  document.getElementById("voiceBtn").addEventListener("click", toggleVoiceRecording);

  // ---- Video ----
  document.getElementById("videoBtn").addEventListener("click", toggleVideo);
  document.getElementById("captureBtn").addEventListener("click", captureFrame);
  document.getElementById("stopVideoBtn").addEventListener("click", stopVideo);

  // ---- File attach ----
  document.getElementById("fileBtn").addEventListener("click", () =>
    document.getElementById("fileInput").click()
  );
  document.getElementById("fileInput").addEventListener("change", handleFileUpload);

  // ---- Analyse tab ----
  document.getElementById("analyzeSelectionBtn").addEventListener("click", analyseSelection);
  document.getElementById("analyzeSheetBtn").addEventListener("click", analyseSheet);

  // ---- Charts tab ----
  document.getElementById("autoChartBtn").addEventListener("click", doAutoChart);
  document.querySelectorAll(".chart-type-btn").forEach((btn) => {
    btn.addEventListener("click", () => doInsertChart(btn.dataset.type));
  });

  // ---- Detect tab ----
  document.getElementById("detectAnomaliesBtn").addEventListener("click", doDetectAnomalies);
  document.getElementById("detectTrendsBtn").addEventListener("click", doDetectTrends);
  document.getElementById("highlightAnomaliesBtn").addEventListener("click", doHighlightAnomalies);

  // ---- Load saved settings into modal ----
  document.getElementById("aiBackendUrl").value  = cfg.baseUrl;
  document.getElementById("aiModelName").value   = cfg.model || "";
  document.getElementById("pyBackendUrl").value  = cfg.pyBackendUrl || "";
  document.getElementById("maxTokens").value     = cfg.maxTokens;
  document.getElementById("temperature").value   = cfg.temperature;
  document.getElementById("tempDisplay").textContent = cfg.temperature;
}

/* ─────────────────────────────────────────────
   AI Connection status
───────────────────────────────────────────── */
async function checkAIConnection() {
  setStatus("connecting");
  try {
    const model = await detectModel(cfg.baseUrl);
    if (model) {
      currentModel = model;
      cfg.model = cfg.model || model;
      setStatus("connected", model);
    } else {
      setStatus("disconnected", "No model found");
    }
  } catch {
    setStatus("disconnected", "Cannot connect");
  }
}

function setStatus(state, label) {
  const dot  = document.getElementById("statusDot");
  const text = document.getElementById("statusText");
  const badge = document.getElementById("modelBadge");

  dot.className = `status-dot ${state}`;
  text.textContent =
    state === "connected"    ? "Connected" :
    state === "disconnected" ? "Disconnected" :
    state === "thinking"     ? "Thinking…" : "Connecting…";
  if (label) badge.textContent = label;
}

/* ─────────────────────────────────────────────
   Chat
───────────────────────────────────────────── */
async function sendMessage() {
  const input = document.getElementById("chatInput");
  const text = input.value.trim();
  if (!text || isGenerating) return;

  input.value = "";
  appendMessage("user", text);
  chatHistory.push({ role: "user", content: text });

  isGenerating = true;
  document.getElementById("sendBtn").disabled = true;
  setStatus("thinking", currentModel || "…");

  // Optionally include Excel data as context
  let dataContext = "";
  if (document.getElementById("includeDataToggle").checked) {
    try {
      const data = await readSelection() || await readUsedRange();
      if (data) dataContext = summariseData(data);
    } catch {
      // non-fatal
    }
  }

  const systemPrompt = buildSystemPrompt(dataContext);
  const messages = [
    { role: "system", content: systemPrompt },
    ...chatHistory,
  ];

  // Streaming response
  abortCtrl = new AbortController();
  const bubble = appendMessage("assistant", "");
  const thinkEl = createThinkingDots();
  bubble.appendChild(thinkEl);

  await streamChat({
    messages,
    cfg,
    abortSignal: abortCtrl.signal,
    onChunk: (chunk, full) => {
      thinkEl.remove();
      bubble.innerHTML = markdownToHtml(full);
    },
    onDone: (full) => {
      chatHistory.push({ role: "assistant", content: full });
      bubble.innerHTML = markdownToHtml(full);
      isGenerating = false;
      document.getElementById("sendBtn").disabled = false;
      setStatus("connected", currentModel || "—");
    },
    onError: (err) => {
      thinkEl.remove();
      bubble.innerHTML = `<p class="error-text">⚠️ ${err.message}</p>`;
      isGenerating = false;
      document.getElementById("sendBtn").disabled = false;
      setStatus("disconnected", "Error");
    },
  });
}

function appendMessage(role, text) {
  const messages = document.getElementById("messages");
  const wrapper  = document.createElement("div");
  wrapper.className = `message ${role === "user" ? "user-msg" : "assistant-msg"}`;

  const avatar = document.createElement("div");
  avatar.className = "msg-avatar";
  avatar.textContent = role === "user" ? "👤" : "🤖";

  const bubble = document.createElement("div");
  bubble.className = "msg-bubble";
  bubble.innerHTML = markdownToHtml(text);

  wrapper.appendChild(avatar);
  wrapper.appendChild(bubble);
  messages.appendChild(wrapper);
  messages.scrollTop = messages.scrollHeight;
  return bubble;
}

function createThinkingDots() {
  const el = document.createElement("div");
  el.className = "thinking-dots";
  el.innerHTML = "<span>●</span><span>●</span><span>●</span>";
  return el;
}

function clearChat() {
  chatHistory = [];
  const messages = document.getElementById("messages");
  messages.innerHTML = "";
  appendMessage("assistant", "Chat cleared. Ask me anything about your Excel data!");
}

/* ─────────────────────────────────────────────
   Voice recording  (Web Speech API / MediaRecorder)
───────────────────────────────────────────── */
async function toggleVoiceRecording() {
  if (mediaRecorder && mediaRecorder.state === "recording") {
    mediaRecorder.stop();
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    mediaRecorder = new MediaRecorder(stream);
    audioChunks   = [];

    mediaRecorder.ondataavailable = (e) => audioChunks.push(e.data);
    mediaRecorder.onstop = async () => {
      document.getElementById("recordingIndicator").classList.add("hidden");
      stream.getTracks().forEach((t) => t.stop());

      // Try Web Speech API transcription first, fall back to caption
      if ("webkitSpeechRecognition" in window || "SpeechRecognition" in window) {
        transcribeWithSpeechAPI();
      } else {
        appendUserFromVoice("[Voice message recorded – transcription not available in this browser]");
      }
    };

    mediaRecorder.start();
    document.getElementById("recordingIndicator").classList.remove("hidden");
  } catch (err) {
    showToast(`Microphone error: ${err.message}`);
  }
}

function transcribeWithSpeechAPI() {
  const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
  const rec = new SpeechRec();
  rec.lang = "en-US";
  rec.onresult = (e) => {
    const transcript = e.results[0][0].transcript;
    document.getElementById("chatInput").value = transcript;
  };
  rec.onerror = () => appendUserFromVoice("[Voice recognition failed]");
  rec.start();
}

function appendUserFromVoice(text) {
  document.getElementById("chatInput").value = text;
}

/* ─────────────────────────────────────────────
   Video capture
───────────────────────────────────────────── */
async function toggleVideo() {
  if (videoStream) { stopVideo(); return; }
  try {
    videoStream = await navigator.mediaDevices.getUserMedia({ video: true });
    document.getElementById("videoPreview").srcObject = videoStream;
    document.getElementById("videoSection").classList.remove("hidden");
  } catch (err) {
    showToast(`Camera error: ${err.message}`);
  }
}

function stopVideo() {
  if (videoStream) {
    videoStream.getTracks().forEach((t) => t.stop());
    videoStream = null;
  }
  document.getElementById("videoSection").classList.add("hidden");
  document.getElementById("videoPreview").srcObject = null;
}

function captureFrame() {
  const video = document.getElementById("videoPreview");
  const canvas = document.createElement("canvas");
  canvas.width  = video.videoWidth;
  canvas.height = video.videoHeight;
  canvas.getContext("2d").drawImage(video, 0, 0);
  const dataUrl = canvas.toDataURL("image/jpeg", 0.85);

  // Insert as context note (image data stays local)
  const input = document.getElementById("chatInput");
  input.value = (input.value ? input.value + "\n" : "") +
    "[Screenshot captured from camera – describe what you see in this image and how it relates to my data]";
  showToast("Frame captured. Edit the message and send.");
  stopVideo();
}

/* ─────────────────────────────────────────────
   File upload
───────────────────────────────────────────── */
async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  e.target.value = "";

  try {
    const text = await file.text();
    const preview = text.slice(0, 1500);
    const input = document.getElementById("chatInput");
    input.value = `Analyse this file (${file.name}):\n\`\`\`\n${preview}${text.length > 1500 ? "\n… (truncated)" : ""}\n\`\`\``;
    showToast(`File "${file.name}" attached.`);
  } catch (err) {
    showToast(`Error reading file: ${err.message}`);
  }
}

/* ─────────────────────────────────────────────
   Analyse Tab
───────────────────────────────────────────── */
async function analyseSelection() {
  showOutput("analysisOutput", "⏳ Analysing selection…");
  try {
    const data = await readSelection();
    if (!data) { showOutput("analysisOutput", "⚠️ No data in selection.", "error"); return; }
    const summaries = analyseDataset(data);
    showOutput("analysisOutput", formatAnalysisReport(summaries), "success");
  } catch (err) {
    showOutput("analysisOutput", `Error: ${err.message}`, "error");
  }
}

async function analyseSheet() {
  showOutput("analysisOutput", "⏳ Analysing sheet…");
  try {
    const data = await readUsedRange();
    if (!data) { showOutput("analysisOutput", "⚠️ Sheet appears empty.", "error"); return; }
    const summaries = analyseDataset(data);
    showOutput("analysisOutput", formatAnalysisReport(summaries), "success");
  } catch (err) {
    showOutput("analysisOutput", `Error: ${err.message}`, "error");
  }
}

/* ─────────────────────────────────────────────
   Charts Tab
───────────────────────────────────────────── */
async function doAutoChart() {
  showOutput("chartOutput", "⏳ Creating chart…");
  try {
    const data = await readSelection();
    if (!data) { showOutput("chartOutput", "⚠️ Select a range first.", "error"); return; }
    const result = await autoChart(data, data.address);
    showOutput(
      "chartOutput",
      `✅ Created ${result.type} chart: "${result.title}"\n\n📌 Reason: ${result.reason}`,
      "success"
    );
  } catch (err) {
    showOutput("chartOutput", `Error: ${err.message}`, "error");
  }
}

async function doInsertChart(chartType) {
  showOutput("chartOutput", `⏳ Creating ${chartType} chart…`);
  try {
    const data = await readSelection();
    if (!data) { showOutput("chartOutput", "⚠️ Select a range first.", "error"); return; }
    const result = await insertChart(data.address, chartType, data);
    showOutput("chartOutput", `✅ Created "${result.title}" chart.`, "success");
  } catch (err) {
    showOutput("chartOutput", `Error: ${err.message}`, "error");
  }
}

/* ─────────────────────────────────────────────
   Detect Tab
───────────────────────────────────────────── */
let lastAnomalies = [];
let lastData      = null;

async function doDetectAnomalies() {
  showOutput("detectionOutput", "⏳ Detecting anomalies…");
  try {
    lastData = await readSelection() || await readUsedRange();
    if (!lastData) { showOutput("detectionOutput", "⚠️ No data available.", "error"); return; }
    lastAnomalies = detectAnomalies(lastData);
    showOutput("detectionOutput", formatAnomalyReport(lastAnomalies), lastAnomalies.length ? "success" : "success");
  } catch (err) {
    showOutput("detectionOutput", `Error: ${err.message}`, "error");
  }
}

async function doDetectTrends() {
  showOutput("detectionOutput", "⏳ Detecting trends…");
  try {
    const data = await readSelection() || await readUsedRange();
    if (!data) { showOutput("detectionOutput", "⚠️ No data available.", "error"); return; }
    const trends = detectTrends(data);
    showOutput("detectionOutput", formatTrendReport(trends), "success");
  } catch (err) {
    showOutput("detectionOutput", `Error: ${err.message}`, "error");
  }
}

async function doHighlightAnomalies() {
  if (!lastAnomalies.length) {
    showToast("Run anomaly detection first.");
    return;
  }
  if (!lastData) return;

  // Determine row offset: +1 if header row present, +1 for Excel 1-based
  const rowOffset = lastData.headers ? 2 : 1;
  const addresses = lastAnomalies.map((a) =>
    toA1(a.rowIndex + rowOffset - 1, a.colIndex)
  );

  const colorMap = { high: "#FF4444", medium: "#FF9800", low: "#FFEB3B" };
  try {
    for (const a of lastAnomalies) {
      const addr = toA1(a.rowIndex + rowOffset - 1, a.colIndex);
      await highlightCells([addr], colorMap[a.severity] || "#FF9800");
    }
    showOutput(
      "detectionOutput",
      `✅ Highlighted ${lastAnomalies.length} anomaly cell(s) in the sheet.\n` +
        "🔴 High  🟠 Medium  🟡 Low",
      "success"
    );
  } catch (err) {
    showOutput("detectionOutput", `Error highlighting: ${err.message}`, "error");
  }
}

/* ─────────────────────────────────────────────
   Settings modal
───────────────────────────────────────────── */
function openSettings() {
  // Pre-fill from current config
  document.getElementById("aiBackendUrl").value  = cfg.baseUrl;
  document.getElementById("aiModelName").value   = cfg.model || "";
  document.getElementById("pyBackendUrl").value  = cfg.pyBackendUrl || "";
  document.getElementById("maxTokens").value     = cfg.maxTokens;
  document.getElementById("temperature").value   = cfg.temperature;
  document.getElementById("tempDisplay").textContent = cfg.temperature;
  document.getElementById("settingsModal").classList.remove("hidden");
}

function closeSettings() {
  document.getElementById("settingsModal").classList.add("hidden");
}

function saveSettingsUI() {
  cfg.baseUrl      = document.getElementById("aiBackendUrl").value.trim() || cfg.baseUrl;
  cfg.model        = document.getElementById("aiModelName").value.trim() || null;
  cfg.pyBackendUrl = document.getElementById("pyBackendUrl").value.trim();
  cfg.maxTokens    = parseInt(document.getElementById("maxTokens").value, 10) || 2048;
  cfg.temperature  = parseFloat(document.getElementById("temperature").value) || 0.7;
  saveConfig(cfg);
  closeSettings();
  checkAIConnection();
  showToast("Settings saved.");
}

async function testConnectionUI() {
  document.getElementById("testConnection").textContent = "Testing…";
  const url = document.getElementById("aiBackendUrl").value.trim() || cfg.baseUrl;
  const model = await detectModel(url);
  document.getElementById("testConnection").textContent = "Test Connection";
  if (model) {
    showToast(`✅ Connected! Model: ${model}`);
  } else {
    showToast("❌ Could not connect to AI backend.");
  }
}

/* ─────────────────────────────────────────────
   Helpers
───────────────────────────────────────────── */
function showOutput(id, text, state = "") {
  const el = document.getElementById(id);
  el.textContent = text;
  el.className   = `output-box ${state}`;
  el.classList.remove("hidden");
}

function showToast(msg) {
  const toast = document.createElement("div");
  toast.style.cssText =
    "position:fixed;bottom:16px;left:50%;transform:translateX(-50%);" +
    "background:#333;color:#fff;padding:8px 16px;border-radius:6px;" +
    "font-size:12px;z-index:9999;box-shadow:0 2px 8px rgba(0,0,0,0.5);" +
    "animation:fadeIn 0.2s ease;max-width:280px;text-align:center;";
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

/**
 * Minimal Markdown→HTML for chat bubbles.
 * Supports: bold, italic, inline code, code blocks, bullet lists, links.
 */
function markdownToHtml(md) {
  if (!md) return "";
  return md
    // Code blocks
    .replace(/```(\w*)\n([\s\S]*?)```/g, "<pre><code>$2</code></pre>")
    // Inline code
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    // Bold
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    // Italic
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    // Headers (h3-h1)
    .replace(/^### (.+)$/gm, "<strong>$1</strong>")
    .replace(/^## (.+)$/gm, "<strong>$1</strong>")
    .replace(/^# (.+)$/gm, "<strong>$1</strong>")
    // Bullet lists
    .replace(/^[-*] (.+)$/gm, "• $1")
    // Numbered lists
    .replace(/^\d+\. (.+)$/gm, "• $1")
    // Links
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>')
    // Newlines → <br> (outside of pre blocks)
    .replace(/\n/g, "<br>");
}
