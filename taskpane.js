/**
 * ================================================================
 * taskpane.js — Orchestrateur principal Accuracy Profile Add-in v2
 * Connecté au backend Python via ApiClient
 *
 * FIXES APPLIQUÉS :
 *  1. Suppression des import ES6 (incompatibles avec <script> classique)
 *  2. ApiClient.healthCheck() → ApiClient.health() [méthode correcte]
 *  3. Ajout de blocs finally sur tous les handlers async (boutons toujours réactivés)
 *  4. Démonstration séparée indirect / direct
 *  5. Gestion de l'état backend (bannière connexion)
 *  6. Meilleure gestion des erreurs + messages utilisateur
 * ================================================================
 */
"use strict";

// ─── Modules chargés via <script> — disponibles sur window ──────────────────
// ApiClient, ExcelBridge, DemoData, ReportGenerator sont injectés globalement
// par leurs fichiers respectifs. Pas besoin d'import ES6.

// ─── État global ──────────────────────────────────────────────────────────────
const APP = {
  planValidation: null,
  planEtalonnage: null,
  results:        null,   // Réponse complète du backend v2
  config:         {},
  aiContent:      "",
  chatHistory:    [],
  profileChart:   null,
  backendOnline:  false,
};

// ─── Démarrage ─────────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("⚠ Excel requis");
    _showBanner("⚠ Hôte non supporté — ouvrez dans Excel", false);
    return;
  }
  _initApp();
});

function _initApp() {
  _setupNavigation();
  _setupDataHandlers();
  _setupCalcHandlers();
  _setupProfileHandlers();
  _setupAIHandlers();
  _setupReportHandlers();
  _checkBackendConnection();
  setStatus("Accuracy Profile v2 ✓");
  log("Prêt. Chargez vos données ou utilisez les données de démonstration.", "info");
}


// ═══════════════════════════════════════════════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════════════════════════════════════════════

function _setupNavigation() {
  document.querySelectorAll(".nav-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(tab.dataset.panel)?.classList.add("active");
    });
  });
}


// ═══════════════════════════════════════════════════════════════════════════════
// VÉRIFICATION BACKEND
// ═══════════════════════════════════════════════════════════════════════════════

async function _checkBackendConnection() {
  _showBanner("Connexion en cours…", null);
  try {
    // FIX : méthode correcte est health(), pas healthCheck()
    await ApiClient.health();
    APP.backendOnline = true;
    _showBanner("✓ Connecté", true);
    setStatus("Backend ✓");
    log("Connecté", "ok");
  } catch (e) {
    APP.backendOnline = false;
    _showBanner("⚠ Backend inaccessible — vérifiez la connexion", false);
    setStatus("⚠ Backend hors ligne");
    log("Backend inaccessible : " + _errMsg(e), "warn");
  }
}

function _showBanner(msg, online) {
  const banner = document.getElementById("backend-banner");
  const dot    = document.getElementById("backend-dot");
  const txt    = document.getElementById("backend-msg");
  if (!banner) return;
  if (txt) txt.textContent = msg;
  if (dot) {
    dot.className = "backend-dot" + (online === true ? " online" : online === false ? " offline" : " pending");
  }
  banner.className = "backend-banner" + (online === true ? " ok" : online === false ? " err" : "");
}


// ═══════════════════════════════════════════════════════════════════════════════
// DONNÉES
// ═══════════════════════════════════════════════════════════════════════════════

function _setupDataHandlers() {
  // Afficher/masquer étalonnage selon le type
  document.getElementById("cfg-type").addEventListener("change", function () {
    document.getElementById("etalon-range-row").style.display =
      this.value === "indirect" ? "block" : "none";
  });

  // Détection auto des plages
  document.getElementById("btn-detect-validation").addEventListener("click", async () => {
    try {
      const r = await ExcelBridge.detectUsedRange();
      document.getElementById("range-validation").value = r;
      toast("Plage détectée : " + r, "info");
    } catch (e) { toast("Erreur : " + _errMsg(e), "err"); }
  });

  document.getElementById("btn-detect-etalonnage").addEventListener("click", async () => {
    try {
      const r = await ExcelBridge.detectUsedRange();
      document.getElementById("range-etalonnage").value = r;
      toast("Plage détectée : " + r, "info");
    } catch (e) { toast("Erreur : " + _errMsg(e), "err"); }
  });

  // Génération du plan
  document.getElementById("btn-generate-plan").addEventListener("click", async () => {
    const K = parseInt(document.getElementById("cfg-K").value);
    const I = parseInt(document.getElementById("cfg-I").value);
    const J = parseInt(document.getElementById("cfg-J").value);
    const u = document.getElementById("cfg-unite").value;
    const methodType = document.getElementById("cfg-type").value;
    setBtnLoading("btn-generate-plan", true, "Génération…");
    try {
      const resV = await ExcelBridge.generatePlanValidation(K, I, J, u, methodType);
      // Auto-remplir la plage validation avec l'adresse qualifiée (ex: "Plan_Validation!A1:F28")
      // Si generatePlanValidation retourne qualifiedRange, on l'utilise, sinon adresse simple
      const rangeV = resV.qualifiedRange || `Plan_Validation!A1:F${resV.rows + 1}`;
      document.getElementById("range-validation").value = rangeV;
      toast(`✅ Plan généré : ${resV.rows} lignes → onglet "${resV.sheetName}"`, "info");
      log(`Plan ${methodType} : K=${K}, I=${I}, J=${J} → ${resV.rows} lignes | plage : ${rangeV}`, "ok");

      if (methodType === "indirect") {
        const resE = await ExcelBridge.generatePlanEtalonnage(I, 2, 2, u);
        // Auto-remplir la plage étalonnage
        const rangeE = resE.qualifiedRange || `Plan_Etalonnage!A1:E${resE.rows + 1}`;
        document.getElementById("range-etalonnage").value = rangeE;
        log(`Plan étalonnage créé → ${resE.rows} lignes | plage : ${rangeE}`, "ok");
      }

      toast("💡 Plages pré-remplies. Saisissez vos données puis cliquez « Importer ».", "info", 5000);
    } catch (e) {
      toast("Erreur : " + _errMsg(e), "err");
      log("Erreur génération plan : " + _errMsg(e), "err");
    } finally {
      setBtnLoading("btn-generate-plan", false, "⊞ Générer le plan");
    }
  });

  // Templates
  document.getElementById("btn-generate-templates").addEventListener("click", async () => {
    _readConfigFromUI();
    setBtnLoading("btn-generate-templates", true, "Génération…");
    try {
      const res = await ExcelBridge.generateBlankTemplates({
        K: parseInt(document.getElementById("cfg-K").value) || 3,
        I: parseInt(document.getElementById("cfg-I").value) || 3,
        J: parseInt(document.getElementById("cfg-J").value) || 3,
        unite:      document.getElementById("cfg-unite").value,
        methodType: document.getElementById("cfg-type").value,
      });
      toast("✅ Feuilles de saisie créées", "info");
      log("Templates : " + [res.sheetP, res.sheetV, res.sheetE].filter(Boolean).join(", "), "ok");
    } catch (e) {
      toast("Erreur : " + _errMsg(e), "err");
      log("Erreur templates : " + _errMsg(e), "err");
    } finally {
      setBtnLoading("btn-generate-templates", false, "▦ Feuilles de saisie");
    }
  });

  // Import + calcul
  document.getElementById("btn-import").addEventListener("click", _handleImportAndCalc);

  // Démo — deux boutons séparés (indirect / direct)
  document.getElementById("btn-demo-indirect").addEventListener("click", () => _handleDemo("indirect"));
  document.getElementById("btn-demo-direct").addEventListener("click",   () => _handleDemo("direct"));
}


async function _handleImportAndCalc() {
  const rangeVal   = document.getElementById("range-validation").value.trim();
  const rangeEta   = document.getElementById("range-etalonnage").value.trim();
  const methodType = document.getElementById("cfg-type").value;

  if (!rangeVal) { toast("Saisissez la plage du plan de validation", "warn"); return; }

  // Pour une méthode indirecte la plage d'étalonnage est obligatoire
  if (methodType === "indirect" && !rangeEta) {
    toast("Méthode indirecte : saisissez aussi la plage du plan d'étalonnage", "warn");
    return;
  }

  setBtnLoading("btn-import", true, "Import et calcul en cours…");
  try {
    log(`Lecture plan de validation : ${rangeVal}`, "info");
    APP.planValidation = await ExcelBridge.readPlanValidation(rangeVal);
    log(`✓ ${APP.planValidation.length} lignes lues dans Plan_Validation`, "ok");

    APP.planEtalonnage = [];
    if (methodType === "indirect") {
      log(`Lecture plan d'étalonnage : ${rangeEta}`, "info");
      APP.planEtalonnage = await ExcelBridge.readPlanEtalonnage(rangeEta);
      log(`✓ ${APP.planEtalonnage.length} lignes lues dans Plan_Etalonnage`, "ok");
    }

    _readConfigFromUI();
    await _runAnalysis();
    _renderPreview();
    toast(`✅ ${APP.planValidation.length} mesures importées et analysées`, "ok");
  } catch (e) {
    toast("Erreur import : " + _errMsg(e), "err");
    log("❌ Erreur import : " + _errMsg(e), "err");
  } finally {
    // FIX : finally garantit que le bouton est toujours réactivé
    setBtnLoading("btn-import", false, "⊞ Importer et calculer");
  }
}


function _handleDemo(type) {
  // Forcer le sélecteur de type au type demandé
  document.getElementById("cfg-type").value = type;
  document.getElementById("etalon-range-row").style.display =
    type === "indirect" ? "block" : "none";

  if (type === "direct") {
    APP.planValidation = DemoData.DIRECT_VALIDATION;
    APP.planEtalonnage = [];
    _applyDemoConfig(DemoData.DIRECT_CONFIG);
    log(`Démo directe : ${APP.planValidation.length} mesures chargées`, "ok");
  } else {
    APP.planValidation = DemoData.FEINBERG_VALIDATION;
    APP.planEtalonnage = DemoData.FEINBERG_ETALONNAGE;
    _applyDemoConfig(DemoData.FEINBERG_CONFIG);
    log(`Démo Feinberg : ${APP.planValidation.length} mesures · ${APP.planEtalonnage.length} étalons`, "ok");
  }

  _readConfigFromUI();
  _runAnalysis().catch(e => {
    toast("Erreur analyse : " + _errMsg(e), "err");
    log("Erreur : " + _errMsg(e), "err");
  });
  _renderPreview();
  toast(`✅ Données de démonstration (${type}) chargées`, "info");
}

function _applyDemoConfig(cfg) {
  document.getElementById("cfg-methode").value  = cfg.methode    || "";
  document.getElementById("cfg-materiau").value = cfg.materiau   || "";
  document.getElementById("cfg-unite").value    = cfg.unite      || "";
  document.getElementById("cfg-type").value     = cfg.methodType || "indirect";
  document.getElementById("cfg-lambda").value   = ((cfg.lambda ?? 0.10) * 100).toFixed(1);
  document.getElementById("cfg-beta").value     = ((cfg.beta   ?? 0.80) * 100).toFixed(0);
  document.getElementById("cfg-K").value        = cfg.K || 3;
  document.getElementById("cfg-I").value        = cfg.I || 3;
  document.getElementById("cfg-J").value        = cfg.J || 3;
  document.getElementById("etalon-range-row").style.display =
    cfg.methodType === "indirect" ? "block" : "none";
}

function _readConfigFromUI() {
  APP.config = {
    methode:    document.getElementById("cfg-methode").value   || "Méthode analytique",
    materiau:   document.getElementById("cfg-materiau").value  || "—",
    unite:      document.getElementById("cfg-unite").value     || "",
    methodType: document.getElementById("cfg-type").value      || "indirect",
    modelType:  document.getElementById("cfg-model").value     || "linear",
    framework:  document.getElementById("cfg-framework").value || "iso5725",
    lambda:     parseFloat(document.getElementById("cfg-lambda").value) / 100 || 0.10,
    beta:       parseFloat(document.getElementById("cfg-beta").value)   / 100 || 0.80,
    alpha:      0.05,
  };
}

function _renderPreview() {
  if (!APP.planValidation?.length) return;
  document.getElementById("preview-thead").innerHTML = `
    <tr><th>Niveau</th><th>Série</th><th>Rép.</th><th>X réf.</th><th>Y réponse</th></tr>`;
  const rows = APP.planValidation.slice(0, 12);
  document.getElementById("preview-tbody").innerHTML =
    rows.map(r => `<tr>
      <td>${r.niveau}</td><td>${r.serie}</td><td>${r.rep}</td>
      <td>${r.xRef}</td><td>${r.yResponse}</td>
    </tr>`).join("") +
    (APP.planValidation.length > 12
      ? `<tr><td colspan="5" style="text-align:center;color:var(--text-muted);font-style:italic">
           … et ${APP.planValidation.length - 12} autres lignes</td></tr>`
      : "");
  document.getElementById("preview-n").textContent = APP.planValidation.length;
  document.getElementById("preview-card").style.display = "block";
}


// ═══════════════════════════════════════════════════════════════════════════════
// ANALYSE — APPEL BACKEND
// ═══════════════════════════════════════════════════════════════════════════════

async function _runAnalysis() {
  if (!APP.planValidation?.length) return;

  setStatus("Analyse en cours…");
  log("Envoi des données au backend Python…", "info");

  try {
    APP.results = await ApiClient.analyze({
      planValidation: APP.planValidation,
      planEtalonnage: APP.planEtalonnage || [],
      config:         APP.config,
      charts:         true,
      chartFormat:    "png_base64",
      normative:      true,
      interpret:      true,
      useLLM:         false,
    });

    if (APP.results.status === "error") {
      throw new Error(APP.results.detail || "Erreur de calcul backend");
    }

    const dur   = APP.results.meta?.duration_s ?? "?";
    const score = APP.results.qualityScore?.overall;
    const label = APP.results.qualityScore?.label || "";
    log(`Calculs terminés en ${dur}s — Score qualité : ${score}/100 (${label})`, "ok");
    setStatus(`Score : ${score}/100 (${label})`);

    _renderCalcResults();
    _renderProfileChart();

  } catch (e) {
    toast("Erreur backend : " + _errMsg(e), "err");
    log("Erreur analyse : " + _errMsg(e), "err");
    setStatus("⚠ Erreur backend");
    throw e;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// CALCULS — RENDU
// ═══════════════════════════════════════════════════════════════════════════════

function _setupCalcHandlers() {
  document.getElementById("btn-write-results").addEventListener("click", async () => {
    if (!APP.results) { toast("Aucun résultat à écrire", "warn"); return; }
    setBtnLoading("btn-write-results", true, "Écriture…");
    try {
      const sheet = await ExcelBridge.writeAnalysisResults(APP.results, APP.config);
      toast(`✅ Résultats écrits → onglet "${sheet}"`, "ok");
    } catch (e) {
      toast("Erreur : " + _errMsg(e), "err");
    } finally {
      setBtnLoading("btn-write-results", false, "📊 Écrire les résultats dans Excel");
    }
  });
}

function _renderCalcResults() {
  const { models = {}, criteria = [], tolerances = [], outliers = [],
          normativeChecks = [], normality = [], homogeneity = {},
          qualityScore = {}, interpretation = [] } = APP.results;

  // ── Score qualité ────────────────────────────────────────────────────────────
  if (qualityScore.overall !== undefined) {
    const scoreColor = qualityScore.overall >= 75 ? "var(--valid)"
                     : qualityScore.overall >= 55 ? "var(--warning)" : "var(--invalid)";
    document.getElementById("quality-content").innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:8px">
        ${_scoreCard("Global",       qualityScore.overall,    qualityScore.label, scoreColor)}
        ${_scoreCard("Justesse",     qualityScore.justesse)}
        ${_scoreCard("Fidélité",     qualityScore.fidelite)}
        ${_scoreCard("Profil",       qualityScore.profil)}
        ${_scoreCard("Normalité",    qualityScore.normalite)}
        ${_scoreCard("Homogénéité",  qualityScore.homogeneite)}
      </div>
      ${(qualityScore.details || []).map(d => `<div class="api-note">⚠ ${d}</div>`).join("")}`;
  }

  // ── Modèles d'étalonnage ─────────────────────────────────────────────────────
  const etaCard = document.getElementById("etalonnage-card");
  if (APP.config.methodType === "indirect" && Object.keys(models).length > 0) {
    etaCard.style.display = "block";
    document.getElementById("etalonnage-results").innerHTML =
      Object.entries(models).map(([serie, m]) => `
        <div class="mono-block" style="margin-bottom:6px">
          <div class="model-row">
            <span class="model-label">${serie}</span>
            <span>Y = <strong>${(m.a1||0).toFixed(4)}</strong>·X + ${(m.a0||0).toFixed(4)}
              <span style="color:var(--text-muted);font-size:9px;margin-left:8px">
                R²=${(m.r2||0).toFixed(6)} · r=${(m.r||0).toFixed(6)} · n=${m.n} · RMSE=${(m.rmse||0).toFixed(4)}
              </span>
            </span>
          </div>
        </div>`).join("");
  } else {
    etaCard.style.display = "none";
  }

  // ── Critères ISO 5725-2 ───────────────────────────────────────────────────────
  const lambda = APP.config.lambda ?? 0.10;
  document.getElementById("criteria-tbody").innerHTML = criteria.map(c => {
    const biasFlag = Math.abs(c.bRel) > lambda * 100;
    return `<tr>
      <td><strong>${c.niveau}</strong></td>
      <td>${(c.xMean||0).toFixed(4)}</td>
      <td>${(c.zMean||0).toFixed(4)}</td>
      <td>${(c.sr||0).toFixed(4)}</td>
      <td>${(c.sB||0).toFixed(4)}</td>
      <td>${(c.sFI||0).toFixed(4)}</td>
      <td>${(c.cv||0).toFixed(2)}</td>
      <td>${(c.cvR||0).toFixed(2)}</td>
      <td style="${biasFlag ? "color:var(--invalid);font-weight:700" : ""}">${(c.bRel||0).toFixed(3)}</td>
      <td>${(c.recouvMoy||0).toFixed(3)}</td>
      <td>${c.shapiro_p != null ? (c.shapiro_p).toFixed(4) + (c.shapiro_normal ? " ✓" : " ✗") : "—"}</td>
    </tr>`;
  }).join("");

  // ── Intervalles de tolérance ─────────────────────────────────────────────────
  document.getElementById("tolerance-tbody").innerHTML = tolerances.map(t => `<tr>
    <td><strong>${t.niveau}</strong></td>
    <td>${(t.xMean||0).toFixed(4)}</td>
    <td>${(t.sIT||0).toFixed(4)}</td>
    <td>${(t.ktol||0).toFixed(4)}</td>
    <td>${t.nu ?? "—"}</td>
    <td>${(t.ltbRel||0).toFixed(3)}</td>
    <td>${(t.lthRel||0).toFixed(3)}</td>
    <td>${(t.laBasse||90).toFixed(1)}</td>
    <td>${(t.laHaute||110).toFixed(1)}</td>
    <td>${(t.errorTotal||0).toFixed(3)}</td>
    <td><span class="status-${t.accept ? "valid" : "invalid"}">${t.accept ? "VALIDE" : "NON VALIDE"}</span></td>
  </tr>`).join("");

  // ── Tests statistiques ────────────────────────────────────────────────────────
  const levene = homogeneity?.levene;
  document.getElementById("stat-tests-content").innerHTML = `
    <div style="font-family:var(--font-data);font-size:10px;line-height:2">
      <div><strong>Normalité (Shapiro-Wilk) :</strong>
        ${normality.map(n =>
          `Niv.${n.niveau} : W=${(n.stat||0).toFixed(4)} p=${(n.p_value||0).toFixed(4)} ${n.normal ? "✓" : "✗"}`
        ).join(" · ") || "—"}
      </div>
      ${levene ? `<div><strong>Homogénéité (Levene) :</strong>
        stat=${(levene.stat||0).toFixed(4)} p=${(levene.p_value||0).toFixed(4)}
        → ${levene.homogeneous ? "✓ Variances homogènes" : "⚠ Hétérogénéité détectée"}
      </div>` : ""}
    </div>`;

  // ── Aberrants (Grubbs) ────────────────────────────────────────────────────────
  document.getElementById("outlier-results").innerHTML = outliers.map(o => {
    const cls = o.suspect ? (o.classification === "aberrant" ? "aberrant" : "suspect") : "ok";
    return `<div class="outlier-row ${cls}">
      <strong>Niveau ${o.niveau}</strong> (X̄=${(o.xMean||0).toFixed(3)}, n=${o.n}) —
      G=${(o.G||0).toFixed(4)} / G<sub>crit</sub>=${(o.Gcrit||0).toFixed(4)}
      → ${o.suspect
        ? `⚠ <strong>${(o.classification||"SUSPECT").toUpperCase()}</strong> : ${(o.suspectVal||0).toFixed(6)}`
        : "✅ Aucun aberrant"}
    </div>`;
  }).join("");

  // ── Vérifications normatives ──────────────────────────────────────────────────
  const normCard = document.getElementById("normative-card");
  if (normativeChecks?.length > 0) {
    normCard.style.display = "block";
    document.getElementById("normative-content").innerHTML = normativeChecks.map(item =>
      `<div class="outlier-row ${_sevToClass(item.severity)}" style="flex-direction:column;gap:2px">
        <div><strong>[${item.category}]</strong></div>
        <div>${item.message}</div>
      </div>`
    ).join("");
  } else {
    normCard.style.display = "none";
  }

  document.getElementById("calc-empty").style.display   = "none";
  document.getElementById("calc-results").style.display = "block";
}

function _scoreCard(label, val, sub = "", color = "var(--navy-800)") {
  return `<div style="background:var(--bg-elevated);border:1px solid var(--border);border-radius:6px;padding:8px;text-align:center">
    <div style="font-size:18px;font-weight:700;color:${color}">${val != null ? (+val).toFixed(0) : "—"}</div>
    <div style="font-size:9px;color:var(--text-muted);text-transform:uppercase;letter-spacing:.06em">${label}</div>
    ${sub ? `<div style="font-size:9px;color:${color};font-weight:600">${sub}</div>` : ""}
  </div>`;
}

function _sevToClass(sev) {
  return { success: "ok", info: "ok", warning: "suspect", critical: "aberrant" }[sev] || "ok";
}


// ═══════════════════════════════════════════════════════════════════════════════
// PROFIL D'EXACTITUDE — RENDU
// ═══════════════════════════════════════════════════════════════════════════════

function _setupProfileHandlers() {
  document.getElementById("btn-insert-chart").addEventListener("click", async () => {
    if (!APP.results) { toast("Aucun profil à insérer", "warn"); return; }
    setBtnLoading("btn-insert-chart", true, "Insertion…");
    try {
      const chartBase64 = APP.results.charts?.profile || null;
      const sheet = await ExcelBridge.insertProfileChart(
        APP.results.tolerances, APP.config, chartBase64
      );
      toast(`✅ Graphique inséré → onglet "${sheet}"`, "ok");
    } catch (e) {
      toast("Erreur : " + _errMsg(e), "err");
    } finally {
      setBtnLoading("btn-insert-chart", false, "📊 Insérer le graphique dans Excel");
    }
  });
}

function _renderProfileChart() {
  if (!APP.results?.tolerances?.length) return;

  const { tolerances, validity = {}, charts = {} } = APP.results;
  const lambda  = APP.config.lambda ?? 0.10;
  const beta    = APP.config.beta   ?? 0.80;
  const laBasse = tolerances[0]?.laBasse ?? (1 - lambda) * 100;
  const laHaute = tolerances[0]?.laHaute ?? (1 + lambda) * 100;

  // ── Image backend (PNG) ─────────────────────────────────────────────────────
  if (charts.profile) {
    const imgCard = document.getElementById("profile-img-card");
    const img     = document.getElementById("profile-img");
    img.src       = charts.profile;
    imgCard.style.display = "block";
  }
  if (charts.anova) {
    const anovaCard = document.getElementById("anova-img-card");
    const anovaImg  = document.getElementById("anova-img");
    anovaImg.src    = charts.anova;
    anovaCard.style.display = "block";
  }

  // ── Légende graphique ────────────────────────────────────────────────────────
  document.getElementById("profile-legend").innerHTML = `
    <div class="legend-item">
      <div class="legend-line" style="background:#F5A623"></div>
      <span>Taux de recouvrement</span>
    </div>
    <div class="legend-item">
      <div class="legend-line" style="border-top:2px solid #1A3050;width:18px;background:none"></div>
      <span>LTB / LTH (${(beta*100).toFixed(0)}%-expectation)</span>
    </div>
    <div class="legend-item">
      <div class="legend-line" style="border-top:2px dashed #EF4444;width:18px;background:none"></div>
      <span>Limites d'acceptabilité (±${(lambda*100).toFixed(0)}%)</span>
    </div>`;

  // ── Chart.js interactif ─────────────────────────────────────────────────────
  const labels      = tolerances.map(t => `${(t.xMean||0).toFixed(3)} ${APP.config.unite || ""}`);
  const recouv      = tolerances.map(t => +((t.recouvRel || 100).toFixed(3)));
  const ltb         = tolerances.map(t => +((t.ltbRel || 0).toFixed(3)));
  const lth         = tolerances.map(t => +((t.lthRel || 0).toFixed(3)));
  const pointColors = tolerances.map(t => t.accept ? "#22C55E" : "#EF4444");

  if (APP.profileChart) { APP.profileChart.destroy(); APP.profileChart = null; }

  APP.profileChart = new Chart(document.getElementById("chart-profile"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Taux de recouvrement (%)", data: recouv,
          borderColor: "#F5A623", backgroundColor: "rgba(245,166,35,0.08)",
          pointBackgroundColor: pointColors, pointRadius: 7, pointHoverRadius: 9,
          borderWidth: 2.5, tension: 0.3, fill: false, order: 1,
        },
        {
          label: "LTB (%)", data: ltb,
          borderColor: "#1A3050", backgroundColor: "rgba(26,48,80,0.06)",
          pointRadius: 4, borderWidth: 1.5, tension: 0.3, fill: "+1", order: 2,
        },
        {
          label: "LTH (%)", data: lth,
          borderColor: "#1A3050", pointRadius: 4, borderWidth: 1.5, tension: 0.3, fill: false, order: 3,
        },
        {
          label: `L.Accept. basse (${laBasse.toFixed(0)}%)`,
          data: tolerances.map(() => laBasse),
          borderColor: "#EF4444", borderDash: [7, 4], pointRadius: 0, borderWidth: 1.5, fill: false, order: 4,
        },
        {
          label: `L.Accept. haute (${laHaute.toFixed(0)}%)`,
          data: tolerances.map(() => laHaute),
          borderColor: "#EF4444", borderDash: [7, 4], pointRadius: 0, borderWidth: 1.5, fill: false, order: 5,
        },
        {
          label: "Référence 100%",
          data: tolerances.map(() => 100),
          borderColor: "rgba(140,160,185,0.4)", borderDash: [3, 3],
          pointRadius: 0, borderWidth: 1, fill: false, order: 6,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { labels: { color: "#4A6080", font: { size: 9, family: "'IBM Plex Mono'" }, boxWidth: 14 } },
        tooltip: {
          callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.raw?.toFixed(3)}%` },
        },
      },
      scales: {
        x: {
          ticks: { color: "#4A6080", font: { size: 9 } },
          grid:  { color: "rgba(180,200,220,0.3)" },
          title: { display: true, text: `Concentration (${APP.config.unite || ""})`, color: "#4A6080", font: { size: 9 } },
        },
        y: {
          ticks: { color: "#4A6080", font: { size: 9 }, callback: v => v + "%" },
          grid:  { color: "rgba(180,200,220,0.3)" },
          title: { display: true, text: "Recouvrement / Intervalle (%)", color: "#4A6080", font: { size: 9 } },
          suggestedMin: Math.min(laBasse - 8, ...ltb) - 2,
          suggestedMax: Math.max(laHaute + 8, ...lth) + 2,
        },
      },
    },
  });

  // ── Verdict ─────────────────────────────────────────────────────────────────
  const { valid, partial, nValid, nTotal, pct, validDomain } = validity;
  const domainStr = validDomain
    ? `${(validDomain.xMin||0).toFixed(3)} – ${(validDomain.xMax||0).toFixed(3)} ${APP.config.unite || ""}`
    : "—";
  const verdictEl = document.getElementById("verdict-box");

  if (valid) {
    verdictEl.className = "verdict-box verdict-valid";
    verdictEl.innerHTML = `<strong>✅ MÉTHODE VALIDE</strong> — ${nTotal} niveaux respectent les critères
      β-expectation.<br>Résultats dans ±${(lambda*100).toFixed(0)}% avec probabilité ${(beta*100).toFixed(0)}%.
      Domaine : <strong>${domainStr}</strong>`;
  } else if (partial) {
    verdictEl.className = "verdict-box verdict-partial";
    verdictEl.innerHTML = `<strong>⚠ MÉTHODE PARTIELLEMENT VALIDE</strong> — ${nValid}/${nTotal} niveaux (${pct}%).<br>
      Domaine validé : <strong>${domainStr}</strong>`;
  } else {
    verdictEl.className = "verdict-box verdict-invalid";
    verdictEl.innerHTML = `<strong>❌ MÉTHODE NON VALIDE</strong> — Aucun niveau ne respecte λ=±${(lambda*100).toFixed(0)}%.
      Révision du protocole nécessaire.`;
  }

  // ── Domaine de validité ───────────────────────────────────────────────────────
  document.getElementById("validity-domain").innerHTML = tolerances.map(t => `
    <div class="validity-row">
      <div class="validity-dot" style="background:${t.accept ? "var(--valid)" : "var(--invalid)"}"></div>
      <span>Niveau ${t.niveau} — <strong>${(t.xMean||0).toFixed(3)} ${APP.config.unite||""}</strong> :
        Récouv.=${(t.recouvRel||0).toFixed(2)}% | LTB=${(t.ltbRel||0).toFixed(2)}% / LTH=${(t.lthRel||0).toFixed(2)}%
        | Err.tot.=${(t.errorTotal||0).toFixed(2)}%
        → <strong>${t.accept ? "VALIDE" : "NON VALIDE"}</strong>
      </span>
    </div>`).join("");

  document.getElementById("profile-empty").style.display   = "none";
  document.getElementById("profile-content").style.display = "block";
}


// ═══════════════════════════════════════════════════════════════════════════════
// IA
// ═══════════════════════════════════════════════════════════════════════════════

function _setupAIHandlers() {
  // Clé API
  document.getElementById("btn-save-key")?.addEventListener("click", () => {
    const key = document.getElementById("ai-key")?.value?.trim();
    if (!key) { toast("Saisissez votre clé API", "warn"); return; }
    ApiClient.setApiKey(key);
    toast("✅ Clé API sauvegardée (session courante)", "info");
  });

  // Interprétation par règles (sans LLM)
  document.getElementById("btn-ai-rules").addEventListener("click", async () => {
    if (!APP.results) { toast("Calculez d'abord le profil d'exactitude", "warn"); return; }
    setBtnLoading("btn-ai-rules", true, "Analyse en cours…");
    _showAIResult('<span class="spinner"></span> Moteur de règles…', "RÈGLES");
    try {
      const res   = await ApiClient.interpret(APP.results, APP.config, "full", false);
      const items = res.items || [];
      APP.aiContent = items.map(i => `[${i.category}] ${i.message}`).join("\n");
      _showAIResult(
        items.map(i => `
          <div style="padding:5px 8px;border-radius:4px;margin-bottom:4px;
            background:var(--${_sevBg(i.severity)});border-left:2px solid var(--${_sevColor(i.severity)});
            font-size:11px">
            <strong>[${i.category}]</strong> ${i.message}
          </div>`).join(""),
        "RÈGLES EXPERTES"
      );
      toast("✅ Interprétation terminée", "info");
    } catch (e) {
      _showAIResult(`<span style="color:var(--invalid)">❌ ${_errMsg(e)}</span>`, "ERREUR");
      toast(_errMsg(e), "err");
    } finally {
      setBtnLoading("btn-ai-rules", false, "⚙ Interprétation par règles (sans clé API)");
    }
  });

  // Actions LLM
  const llmActions = {
    "btn-ai-full":     { type: "full",            label: "✦ Diagnostic complet (LLM)" },
    "btn-ai-profile":  { type: "profile",         label: "◈ Interpréter le profil (LLM)" },
    "btn-ai-outliers": { type: "outliers",        label: "⚠ Analyser les aberrants (LLM)" },
    "btn-ai-reco":     { type: "recommendations", label: "📋 Recommandations (LLM)" },
  };

  Object.entries(llmActions).forEach(([btnId, { type, label }]) => {
    document.getElementById(btnId).addEventListener("click", async () => {
      if (!APP.results) { toast("Calculez d'abord le profil d'exactitude", "warn"); return; }
      setBtnLoading(btnId, true, "Analyse LLM…");
      _showAIResult('<span class="spinner"></span> Analyse par IA en cours…', "LLM");
      try {
        const useLLM = ApiClient.hasApiKey();
        const res    = await ApiClient.interpret(APP.results, APP.config, type, useLLM);
        const text   = res.text || res.items?.map(i => `[${i.category}] ${i.message}`).join("\n") || "—";
        APP.aiContent = text;
        _showAIResult(_formatAIText(text), `LLM — ${useLLM ? "LLM activé" : "Règles"}`);
        toast("✅ Analyse IA terminée", "info");
      } catch (e) {
        _showAIResult(`<span style="color:var(--invalid)">❌ ${_errMsg(e)}</span>`, "ERREUR");
        toast(_errMsg(e), "err");
      } finally {
        setBtnLoading(btnId, false, label);
      }
    });
  });

  // Chat
  document.getElementById("btn-chat").addEventListener("click", _handleChat);
  document.getElementById("chat-input").addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); _handleChat(); }
  });
  document.querySelectorAll(".chip").forEach(chip => {
    chip.addEventListener("click", () => {
      document.getElementById("chat-input").value = chip.dataset.p;
      _handleChat();
    });
  });
}

async function _handleChat() {
  const input = document.getElementById("chat-input");
  const msg   = input.value.trim();
  if (!msg) return;

  input.value = "";
  _appendChat("user", msg);
  const typingId = _appendChat("assistant", '<span class="spinner"></span>');

  const context = APP.results ? {
    validity:     APP.results.validity,
    tolerances:   (APP.results.tolerances || []).map(t => ({
      niveau: t.niveau, xMean: t.xMean, recouvRel: t.recouvRel,
      ltbRel: t.ltbRel, lthRel: t.lthRel, accept: t.accept,
    })),
    qualityScore: APP.results.qualityScore,
    config:       { lambda: APP.config.lambda, beta: APP.config.beta, methode: APP.config.methode },
  } : {};

  try {
    const useLLM = ApiClient.hasApiKey();
    const res    = await ApiClient.chat(msg, context, APP.chatHistory, useLLM);
    const text   = res.response || res.items?.join("\n") || "—";
    APP.chatHistory.push({ role: "user",      content: msg });
    APP.chatHistory.push({ role: "assistant", content: text });
    if (APP.chatHistory.length > 20) APP.chatHistory = APP.chatHistory.slice(-20);
    const bubble = document.getElementById(typingId);
    if (bubble) bubble.innerHTML = _formatAIText(text);
  } catch (e) {
    const bubble = document.getElementById(typingId);
    if (bubble) bubble.innerHTML = `<span style="color:var(--invalid)">❌ ${_errMsg(e)}</span>`;
  }
}

let _chatN = 0;
function _appendChat(role, html) {
  const id  = `chat-${++_chatN}`;
  const box = document.getElementById("chat-messages");
  box.insertAdjacentHTML("beforeend", `
    <div class="chat-msg ${role}">
      <div class="chat-bubble" id="${id}">${html}</div>
    </div>`);
  box.scrollTop = box.scrollHeight;
  return id;
}

function _showAIResult(html, sourceLabel = "RÉSULTAT") {
  document.getElementById("ai-result-card").style.display = "block";
  document.getElementById("ai-source-label").textContent  = sourceLabel;
  document.getElementById("ai-content").innerHTML = html;
}

function _formatAIText(text) {
  return String(text)
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\n\n/g, "<br><br>")
    .replace(/\n/g, "<br>")
    .replace(/^(\d+)\.\s+/gm, "<br><b>$1.</b> ");
}

function _sevBg(sev)    { return { success:"valid-bg", info:"amber-bg", warning:"warning-bg", critical:"invalid-bg" }[sev] || "bg-elevated"; }
function _sevColor(sev) { return { success:"valid", info:"amber", warning:"warning", critical:"invalid" }[sev] || "navy-400"; }


// ═══════════════════════════════════════════════════════════════════════════════
// RAPPORT
// ═══════════════════════════════════════════════════════════════════════════════

function _setupReportHandlers() {
  const _getOpts = () => ({
    labo:       document.getElementById("rpt-labo").value,
    analyste:   document.getElementById("rpt-analyste").value,
    ref:        document.getElementById("rpt-ref").value,
    version:    document.getElementById("rpt-version").value,
    params:     document.getElementById("rpt-params").checked,
    etalonnage: document.getElementById("rpt-etalonnage").checked,
    criteria:   document.getElementById("rpt-criteria").checked,
    tolerance:  document.getElementById("rpt-tolerance").checked,
    outliers:   document.getElementById("rpt-outliers").checked,
    normative:  document.getElementById("rpt-normative").checked,
    ai:         document.getElementById("rpt-ai").checked,
  });

  // HTML
  document.getElementById("btn-report-html").addEventListener("click", () => {
    if (!APP.results) { toast("Calculez d'abord le profil", "warn"); return; }
    try {
      const html  = ReportGenerator.generateHTMLReport(
        { results: APP.results, config: APP.config, aiContent: APP.aiContent },
        _getOpts()
      );
      const fname = `Rapport_${(APP.config.methode||"methode").replace(/\s+/g,"_").slice(0,30)}_${new Date().toISOString().slice(0,10)}.html`;
      ReportGenerator.downloadHTMLReport(html, fname);
      toast("✅ Rapport HTML téléchargé", "ok");
      logReport("Rapport HTML généré", "ok");
    } catch (e) {
      toast("Erreur : " + _errMsg(e), "err");
      logReport("Erreur HTML : " + _errMsg(e), "err");
    }
  });

  // PDF via backend
  document.getElementById("btn-report-pdf").addEventListener("click", async () => {
    if (!APP.results) { toast("Calculez d'abord le profil", "warn"); return; }
    setBtnLoading("btn-report-pdf", true, "Génération PDF…");
    try {
      const opts = _getOpts();
      const cfg  = { ...APP.config, laboratoire: opts.labo, analyste: opts.analyste };
      await ReportGenerator.downloadPDFReport(APP.results, cfg);
      toast("✅ Rapport PDF téléchargé", "ok");
      logReport("Rapport PDF généré via backend", "ok");
    } catch (e) {
      toast("Erreur PDF : " + _errMsg(e), "err");
      logReport("Erreur PDF : " + _errMsg(e), "err");
    } finally {
      setBtnLoading("btn-report-pdf", false, "📄 Télécharger le rapport PDF (via backend)");
    }
  });
}


// ═══════════════════════════════════════════════════════════════════════════════
// UTILITAIRES UI
// ═══════════════════════════════════════════════════════════════════════════════

function toast(msg, type = "info", duration = 3200) {
  const icons = { ok: "✅", err: "❌", info: "ℹ", warn: "⚠" };
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${icons[type] || "ℹ"}</span><span>${msg}</span>`;
  document.getElementById("toast-container").appendChild(el);
  setTimeout(() => {
    el.style.transition = "all 0.25s ease";
    el.style.opacity    = "0";
    el.style.transform  = "translateX(16px)";
    setTimeout(() => el.remove(), 250);
  }, duration);
}

function log(msg, type = "info") {
  const el = document.getElementById("log-data");
  if (!el) return;
  const e  = document.createElement("div");
  e.className = `log-entry ${type}`;
  e.innerHTML = `<span class="log-ts">${new Date().toLocaleTimeString("fr-FR")}</span>${msg}`;
  el.appendChild(e);
  el.scrollTop = el.scrollHeight;
}

function logReport(msg, type = "info") {
  const el = document.getElementById("log-report");
  if (!el) return;
  const e  = document.createElement("div");
  e.className = `log-entry ${type}`;
  e.innerHTML = `<span class="log-ts">${new Date().toLocaleTimeString("fr-FR")}</span>${msg}`;
  el.appendChild(e);
  el.scrollTop = el.scrollHeight;
}

function setBtnLoading(id, loading, label) {
  const btn = document.getElementById(id);
  if (!btn) return;
  btn.disabled  = loading;
  btn.innerHTML = loading ? `<span class="spinner"></span> ${label}` : label;
}

function setStatus(msg) {
  const el = document.getElementById("footer-status");
  if (el) el.textContent = msg;
}

/**
 * Extrait un message lisible depuis n'importe quel type d'erreur.
 * Résout le problème "Erreur : [object Object],[object Object]"
 * qui survient quand FastAPI renvoie un tableau d'erreurs Pydantic.
 *
 *   e instanceof Error  → e.message
 *   Array Pydantic      → "[champ] message | [champ2] message2"
 *   Objet quelconque    → JSON.stringify
 */
function _errMsg(e) {
  if (!e) return "Erreur inconnue";
  if (typeof e === "string") return e;
  if (e instanceof Error)   return e.message || String(e);
  if (Array.isArray(e)) {
    return e.map(item => {
      if (item && typeof item === "object") {
        const loc = Array.isArray(item.loc)
          ? item.loc.filter(p => typeof p !== "number").slice(1).join("→")
          : "";
        return loc ? `[${loc}] ${item.msg}` : (item.msg || JSON.stringify(item));
      }
      return String(item);
    }).join(" | ");
  }
  if (typeof e === "object") {
    return e.message || e.detail || e.msg || JSON.stringify(e);
  }
  return String(e);
}