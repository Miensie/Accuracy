/**
 * ================================================================
 * modules/apiClient.js — Client REST vers le backend Python
 * Encapsule tous les appels à l'API Accuracy Profile v2
 *
 * FIX : export default → window.ApiClient (non-ES-module context)
 * ================================================================
 */
"use strict";

const ApiClient = (() => {

  // ─── Configuration ──────────────────────────────────────────────────────────

  const _DEFAULT_BASE = "https://accuracy.onrender.com";

  let _baseUrl  = _DEFAULT_BASE;
  let _apiKey   = "";
  let _provider = "auto";   // "gemini" | "claude" | "auto"
  let _timeout  = 120_000;  // ms

  function setBaseUrl(url) {
    // URL fixe — modification désactivée pour la production
    _baseUrl = _DEFAULT_BASE;
  }

  function getBaseUrl()  { return _baseUrl; }
  function setApiKey(k)  { _apiKey = k.trim(); }
  function setProvider(p){ _provider = p; }
  function getApiKey()   { return _apiKey; }
  function hasApiKey()   { return !!_apiKey; }


  // ─── Fetch générique avec timeout ──────────────────────────────────────────

  async function _fetch(path, options = {}) {
    const url        = `${_baseUrl}${path}`;
    const controller = new AbortController();
    const timer      = setTimeout(() => controller.abort(), _timeout);

    try {
      const resp = await fetch(url, {
        signal:  controller.signal,
        headers: { "Content-Type": "application/json", ...(options.headers || {}) },
        ...options,
      });
      clearTimeout(timer);

      if (!resp.ok) {
        let detail = `HTTP ${resp.status}`;
        try {
          const err = await resp.json();
          // FastAPI renvoie les erreurs Pydantic sous forme de tableau
          // ex : [{ loc: ["body","planValidation",0,"xRef"], msg: "...", type: "..." }]
          // Sans ce traitement → "Erreur : [object Object],[object Object]"
          if (Array.isArray(err.detail)) {
            detail = err.detail
              .map(e => {
                const loc = Array.isArray(e.loc)
                  ? e.loc.filter(p => typeof p !== "number").slice(1).join(" → ")
                  : "";
                return loc ? `[${loc}] ${e.msg}` : (e.msg || JSON.stringify(e));
              })
              .join(" | ");
          } else {
            detail = err.detail || err.message || JSON.stringify(err) || detail;
          }
        } catch { /* ignore */ }
        throw new Error(detail);
      }

      return await resp.json();
    } catch (e) {
      clearTimeout(timer);
      if (e.name === "AbortError") {
        throw new Error(`Délai dépassé (${_timeout / 1000}s) — le backend est-il démarré ?`);
      }
      throw e;
    }
  }

  async function _fetchBlob(path, body) {
    const url = `${_baseUrl}${path}`;
    const resp = await fetch(url, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(body),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    return await resp.blob();
  }


  // ─── Santé du service ───────────────────────────────────────────────────────

  /**
   * BUG FIX : méthode renommée "health" (était "healthCheck" dans taskpane.js)
   */
  async function health() {
    return _fetch("/api/health");
  }


  // ─── Analyse principale ─────────────────────────────────────────────────────

  async function analyze({
    planValidation,
    planEtalonnage = [],
    config         = {},
    charts         = true,
    chartFormat    = "png_base64",
    normative      = true,
    interpret      = true,
    useLLM         = false,
  }) {
    const backendConfig = {
      methode:    config.methode    || "",
      materiau:   config.materiau   || "",
      unite:      config.unite      || "",
      methodType: config.methodType || "indirect",
      modelType:  config.modelType  || "linear",
      beta:       config.beta       ?? 0.80,
      lambdaVal:  config.lambda     ?? config.lambdaVal ?? 0.10,
      alpha:      config.alpha      ?? 0.05,
      framework:  config.framework  || "iso5725",
      laboratoire: config.laboratoire || "",
      analyste:    config.analyste   || "",
    };

    const qs = new URLSearchParams({
      charts:       charts,
      chart_format: chartFormat,
      normative:    normative,
      interpret:    interpret,
    });
    if (useLLM && _apiKey) {
      qs.set("api_key",  _apiKey);
      qs.set("provider", _provider);
    }

    return _fetch(`/accuracy-profile?${qs}`, {
      method: "POST",
      body:   JSON.stringify({ planValidation, planEtalonnage, config: backendConfig }),
    });
  }


  // ─── Format simplifié ───────────────────────────────────────────────────────

  async function analyzeSimple(data, config = {}) {
    const backendConfig = {
      beta:      config.beta      ?? 0.80,
      lambdaVal: config.lambda    ?? 0.10,
      framework: config.framework || "iso5725",
    };
    return _fetch("/api/simple", {
      method: "POST",
      body:   JSON.stringify({ data, config: backendConfig }),
    });
  }


  // ─── Test de Grubbs ─────────────────────────────────────────────────────────

  async function grubbs(data, alpha = 0.05) {
    return _fetch("/api/grubbs", {
      method: "POST",
      body:   JSON.stringify({ data, alpha }),
    });
  }


  // ─── Modèles d'étalonnage ───────────────────────────────────────────────────

  async function calibration(etalonnage, modelType = "linear") {
    return _fetch(`/api/calibration?model_type=${modelType}&include_charts=false`, {
      method: "POST",
      body:   JSON.stringify(etalonnage),
    });
  }


  // ─── Interprétation IA ──────────────────────────────────────────────────────

  async function interpret(analysisData, config, promptType = "full", useLLM = false) {
    const qs = new URLSearchParams({ prompt_type: promptType });
    if (useLLM && _apiKey) {
      qs.set("api_key",  _apiKey);
      qs.set("provider", _provider);
    }
    return _fetch(`/api/interpret?${qs}`, {
      method: "POST",
      body:   JSON.stringify({ analysis_data: analysisData, config }),
    });
  }


  // ─── Chat IA ────────────────────────────────────────────────────────────────

  async function chat(message, context = {}, history = [], useLLM = false) {
    if (useLLM && !_apiKey) throw new Error("Clé API requise pour le chat LLM");
    return _fetch("/api/chat", {
      method: "POST",
      body:   JSON.stringify({
        message,
        context,
        history,
        ...(useLLM && _apiKey ? { api_key: _apiKey, provider: _provider } : {}),
      }),
    });
  }


  // ─── Rapport PDF ────────────────────────────────────────────────────────────

  async function downloadPDF(analysisData, config) {
    const blob = await _fetchBlob("/api/report/pdf", {
      analysis_data: analysisData,
      config,
    });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = `Rapport_Validation_${new Date().toISOString().slice(0, 10)}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }


  // ─── Critères normatifs ─────────────────────────────────────────────────────

  async function getNorms(framework = "all") {
    return _fetch(`/api/norms?framework=${framework}`);
  }


  // ─── API publique ───────────────────────────────────────────────────────────

  return {
    setBaseUrl, getBaseUrl,
    setApiKey,  getApiKey, hasApiKey,
    setProvider,
    health,          // ← méthode correcte (était "healthCheck" dans l'ancienne version)
    analyze,
    analyzeSimple,
    grubbs,
    calibration,
    interpret,
    chat,
    downloadPDF,
    getNorms,
  };
})();

// ⚠️ FIX CRITIQUE : export window global pour chargement <script> classique
// (export default ne fonctionne PAS sans type="module")
window.ApiClient = ApiClient;