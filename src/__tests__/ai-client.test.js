/**
 * Tests for ai-client.js configuration helpers.
 * No network calls are made.
 */

import { loadConfig, saveConfig, buildSystemPrompt, DEFAULT_CONFIG } from "../utils/ai-client.js";

// ─────────────────────────────────────────────
//  loadConfig / saveConfig
// ─────────────────────────────────────────────

describe("loadConfig / saveConfig", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  test("loadConfig returns defaults when nothing is stored", () => {
    const cfg = loadConfig();
    expect(cfg.baseUrl).toBe(DEFAULT_CONFIG.baseUrl);
    expect(cfg.maxTokens).toBe(DEFAULT_CONFIG.maxTokens);
    expect(cfg.temperature).toBe(DEFAULT_CONFIG.temperature);
  });

  test("saveConfig persists data, loadConfig reads it back", () => {
    saveConfig({ ...DEFAULT_CONFIG, baseUrl: "http://localhost:11434/v1", model: "llama3" });
    const cfg = loadConfig();
    expect(cfg.baseUrl).toBe("http://localhost:11434/v1");
    expect(cfg.model).toBe("llama3");
  });

  test("loadConfig merges with defaults (missing keys get defaults)", () => {
    localStorage.setItem("excel_ai_config", JSON.stringify({ maxTokens: 512 }));
    const cfg = loadConfig();
    expect(cfg.maxTokens).toBe(512);
    expect(cfg.baseUrl).toBe(DEFAULT_CONFIG.baseUrl); // default filled in
  });

  test("loadConfig handles corrupt JSON gracefully", () => {
    localStorage.setItem("excel_ai_config", "NOT_JSON{{{");
    const cfg = loadConfig();
    expect(cfg).toBeDefined();
    expect(cfg.baseUrl).toBe(DEFAULT_CONFIG.baseUrl);
  });
});

// ─────────────────────────────────────────────
//  buildSystemPrompt
// ─────────────────────────────────────────────

describe("buildSystemPrompt", () => {
  test("returns a non-empty string", () => {
    const prompt = buildSystemPrompt();
    expect(typeof prompt).toBe("string");
    expect(prompt.length).toBeGreaterThan(50);
  });

  test("includes data context when provided", () => {
    const ctx = "Range: A1:B5\n2 rows × 2 cols";
    const prompt = buildSystemPrompt(ctx);
    expect(prompt).toContain(ctx);
  });

  test("mentions privacy policy", () => {
    const prompt = buildSystemPrompt();
    expect(prompt.toLowerCase()).toContain("privac");
  });

  test("does not include empty data context section when not provided", () => {
    const prompt = buildSystemPrompt("");
    // Should not have a blank "Current Excel data context:" section
    expect(prompt).not.toContain("Current Excel data context:\n\n");
  });
});
