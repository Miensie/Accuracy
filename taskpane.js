/**
 * ================================================================
 * taskpane.js — Orchestrateur principal Accuracy Profile Add-in
 * ================================================================
 */
import "./taskpane.css";


// ─── Configuration Backend ────────────────────────────────────────────────────
const BACKEND_URL = 'https://accuracy.onrender.com'; // Adjust if backend is on different port/host

// ─── Vérification Backend ─────────────────────────────────────────────────────
async function checkBackendHealth() {
  try {
    const response = await fetch(`${BACKEND_URL}/api/health`);
    if (!response.ok) {
      throw new Error(`Backend non accessible: ${response.status}`);
    }
    const data = await response.json();
    return data.status === 'ok';
  } catch (error) {
    console.error('Erreur vérification backend:', error);
    return false;
  }
}

async function getNormativeCriteria() {
  try {
    const response = await fetch(`${BACKEND_URL}/api/norms`);
    if (!response.ok) {
      throw new Error(`Erreur récupération normes: ${response.status}`);
    }
    return await response.json();
  } catch (error) {
    console.error('Erreur récupération normes:', error);
    return null;
  }
}

// ─── État global ──────────────────────────────────────────────────────────────
const APP = {
  planValidation:  null,
  planEtalonnage:  null,
  results:         null,
  config:          {},
  aiContent:       "",
  profileChart:    null,
};

// ─── Démarrage ────────────────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("⚠ Excel requis");
    return;
  }
  initApp();
});

function initApp() {
  // Vérifier la connexion au backend
  checkBackendHealth().then(healthy => {
    if (!healthy) {
      toast("⚠ Backend non accessible. Vérifiez que le serveur Python est démarré.", "warn");
      setStatus("Backend hors ligne");
    } else {
      setStatus("Backend connecté ✓");
    }
  });

  setupNavigation();
  setupDataHandlers();
  setupCalcHandlers();
  setupProfileHandlers();
  setupAIHandlers();
  setupReportHandlers();
  setStatus("Accuracy Profile v1.0 ✓");
  log("Prêt. Chargez vos données ou utilisez les données Feinberg (2010).", "info");
}

// ─── Navigation ───────────────────────────────────────────────────────────────
function setupNavigation() {
  document.querySelectorAll(".nav-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(tab.dataset.panel)?.classList.add("active");
    });
  });
}

// ─── DONNÉES ──────────────────────────────────────────────────────────────────
function setupDataHandlers() {
  // Afficher/masquer le champ étalonnage selon le type de méthode
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
    } catch (e) { toast("Erreur : " + e.message, "err"); }
  });

  document.getElementById("btn-detect-etalonnage").addEventListener("click", async () => {
    try {
      const r = await ExcelBridge.detectUsedRange();
      document.getElementById("range-etalonnage").value = r;
      toast("Plage détectée : " + r, "info");
    } catch (e) { toast("Erreur : " + e.message, "err"); }
  });

  // Génération du plan
  document.getElementById("btn-generate-plan").addEventListener("click", async () => {
    const K = parseInt(document.getElementById("cfg-K").value);
    const I = parseInt(document.getElementById("cfg-I").value);
    const J = parseInt(document.getElementById("cfg-J").value);
    const u = document.getElementById("cfg-unite").value;

    setBtnLoading("btn-generate-plan", true, "Génération…");
    try {
      const methodType = document.getElementById("cfg-type").value;
      const { sheetName, rows } = await ExcelBridge.generatePlanValidation(K, I, J, u, methodType);
      toast(`✅ Plan généré : ${rows} lignes → onglet "${sheetName}"`, "info");
      log(`Plan de validation ${methodType} : K=${K}, I=${I}, J=${J}`, "ok");

      if (methodType === "indirect") {
        await ExcelBridge.generatePlanEtalonnage(I, 2, 2, u);
        log("Plan d'étalonnage créé (2 niveaux, 2 répétitions)", "ok");
      } else {
        log("Méthode directe : pas de plan d'étalonnage requis", "info");
      }
    } catch (e) {
      toast("Erreur génération plan : " + e.message, "err");
      log("Erreur : " + e.message, "err");
    }
    setBtnLoading("btn-generate-plan", false, "⊞ Générer le plan dans Excel");
  });

  // Templates vierges
  document.getElementById("btn-generate-templates").addEventListener("click", async () => {
    readConfigFromUI();
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
      log("Templates créés : " + [res.sheetP, res.sheetV, res.sheetE].filter(Boolean).join(", "), "ok");
    } catch (e) {
      toast("Erreur : " + e.message, "err");
      log("Erreur : " + e.message, "err");
    }
    setBtnLoading("btn-generate-templates", false, "▦ Générer les feuilles de saisie (templates)");
  });

  // Import et calcul
  document.getElementById("btn-import").addEventListener("click", handleImportAndCalc);

  // Démo Feinberg
  document.getElementById("btn-demo").addEventListener("click", handleDemo);
}

async function handleImportAndCalc() {
  const rangeVal = document.getElementById("range-validation").value.trim();
  const rangeEta = document.getElementById("range-etalonnage").value.trim();
  const methodType = document.getElementById("cfg-type").value;

  if (!rangeVal) { toast("Saisissez la plage du plan de validation", "warn"); return; }

  setBtnLoading("btn-import", true, "Import en cours…");
  try {
    // Lecture depuis Excel
    APP.planValidation = await ExcelBridge.readPlanValidation(rangeVal);

    if (methodType === "indirect" && rangeEta) {
      APP.planEtalonnage = await ExcelBridge.readPlanEtalonnage(rangeEta);
    }

    readConfigFromUI();
    runAnalysis();

    // Afficher l'aperçu
    renderPreview();
    toast(`✅ ${APP.planValidation.length} mesures importées`, "info");
    log(`${APP.planValidation.length} mesures lues depuis Excel`, "ok");
  } catch (e) {
    toast("Erreur import : " + e.message, "err");
    log("Erreur : " + e.message, "err");
  }
  setBtnLoading("btn-import", false, "⊞ Importer et calculer");
}

function handleDemo() {
  const demoType = document.getElementById("cfg-type").value;

  if (demoType === "direct") {
    // Démo méthode directe : dosage gravimétrique NaCl dans solution
    APP.planValidation = DemoData.DIRECT_VALIDATION;
    APP.planEtalonnage = [];

    const cfg = DemoData.DIRECT_CONFIG;
    document.getElementById("cfg-methode").value  = cfg.methode;
    document.getElementById("cfg-materiau").value = cfg.materiau;
    document.getElementById("cfg-unite").value    = cfg.unite;
    document.getElementById("cfg-type").value     = cfg.methodType;
    document.getElementById("cfg-lambda").value   = cfg.lambda * 100;
    document.getElementById("cfg-beta").value     = cfg.beta * 100;
    document.getElementById("cfg-K").value        = cfg.K;
    document.getElementById("cfg-I").value        = cfg.I;
    document.getElementById("cfg-J").value        = cfg.J;
    document.getElementById("etalon-range-row").style.display = "none";

    readConfigFromUI();
    runAnalysis();
    renderPreview();
    toast("✅ Démo méthode directe chargée", "info");
    log(`Cas direct : ${APP.planValidation.length} mesures`, "ok");
  } else {
    // Démo méthode indirecte : Feinberg (2010)
    APP.planValidation = DemoData.FEINBERG_VALIDATION;
    APP.planEtalonnage = DemoData.FEINBERG_ETALONNAGE;

    const cfg = DemoData.FEINBERG_CONFIG;
    document.getElementById("cfg-methode").value  = cfg.methode;
    document.getElementById("cfg-materiau").value = cfg.materiau;
    document.getElementById("cfg-unite").value    = cfg.unite;
    document.getElementById("cfg-type").value     = cfg.methodType;
    document.getElementById("cfg-lambda").value   = cfg.lambda * 100;
    document.getElementById("cfg-beta").value     = cfg.beta * 100;
    document.getElementById("cfg-K").value        = cfg.K;
    document.getElementById("cfg-I").value        = cfg.I;
    document.getElementById("cfg-J").value        = cfg.J;
    document.getElementById("etalon-range-row").style.display = "block";

    readConfigFromUI();
    runAnalysis();
    renderPreview();
    toast("✅ Données Feinberg (2010) chargées — méthode indirecte", "info");
    log(`Cas Feinberg : ${APP.planValidation.length} mesures · ${APP.planEtalonnage.length} étalons`, "ok");
  }
}

function readConfigFromUI() {
  APP.config = {
    methode:    document.getElementById("cfg-methode").value  || "Méthode analytique",
    materiau:   document.getElementById("cfg-materiau").value || "—",
    unite:      document.getElementById("cfg-unite").value    || "",
    methodType: document.getElementById("cfg-type").value     || "indirect",
    lambda:     parseFloat(document.getElementById("cfg-lambda").value) / 100 || 0.10,
    beta:       parseFloat(document.getElementById("cfg-beta").value)   / 100 || 0.80,
    modelType:  "linear",
  };
}

function runAnalysis() {
  if (!APP.planValidation?.length) return;

  try {
    // Vérifier si on peut utiliser le format simple
    const canUseSimple = APP.config.methodType === 'direct' && (!APP.planEtalonnage || APP.planEtalonnage.length === 0);

    if (canUseSimple) {
      // Utiliser l'endpoint simplifié
      const simpleData = APP.planValidation.map(row => ({
        concentration: row.xRef,
        replicate: parseInt(row.rep),
        measured: row.yResponse,
        reference: row.xRef
      }));

      const requestData = {
        data: simpleData,
        config: {
          methode: APP.config.methode,
          materiau: APP.config.materiau,
          unite: APP.config.unite,
          methodType: APP.config.methodType,
          modelType: APP.config.modelType,
          beta: APP.config.beta,
          lambdaVal: APP.config.lambda,
          alpha: 0.05,
          framework: "iso5725"
        }
      };

      fetch(`${BACKEND_URL}/api/simple`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestData)
      })
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        if (data.status === 'error') {
          throw new Error(data.detail || 'Erreur de calcul');
        }
        APP.results = data;
        renderCalcResults();
        renderProfileChart();
        log("Calculs simplifiés terminés avec succès", "ok");
      })
      .catch(error => {
        toast("Erreur calcul simplifié : " + error.message, "err");
        log("Erreur calcul simplifié : " + error.message, "err");
        console.error(error);
      });

    } else {
      // Utiliser l'endpoint complet
      const requestData = {
        planValidation: APP.planValidation.map(row => ({
          niveau: row.niveau,
          serie: row.serie,
          rep: row.rep,
          xRef: row.xRef,
          yResponse: row.yResponse
        })),
        planEtalonnage: (APP.planEtalonnage || []).map(row => ({
          serie: row.serie,
          niveau: row.niveau,
          rep: row.rep,
          xEtalon: row.xEtalon,
          yResponse: row.yResponse
        })),
        config: {
          methode: APP.config.methode,
          materiau: APP.config.materiau,
          unite: APP.config.unite,
          methodType: APP.config.methodType,
          modelType: APP.config.modelType,
          beta: APP.config.beta,
          lambdaVal: APP.config.lambda,
          alpha: 0.05,
          framework: "iso5725"
        }
      };

      fetch(`${BACKEND_URL}/accuracy-profile`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestData)
      })
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        if (data.status === 'error') {
          throw new Error(data.detail || 'Erreur de calcul');
        }
        APP.results = data;
        renderCalcResults();
        renderProfileChart();
        log("Calculs complets terminés avec succès", "ok");
      })
      .catch(error => {
        toast("Erreur calcul complet : " + error.message, "err");
        log("Erreur calcul complet : " + error.message, "err");
        console.error(error);
      });
    }

  } catch (e) {
    toast("Erreur préparation données : " + e.message, "err");
    log("Erreur préparation : " + e.message, "err");
    console.error(e);
  }
}

function renderPreview() {
  if (!APP.planValidation?.length) return;

  // Headers
  const thead = document.getElementById("preview-thead");
  thead.innerHTML = `<tr>
    <th>Niveau</th><th>Série</th><th>Rép.</th>
    <th>X réf.</th><th>Y réponse</th>
  </tr>`;

  const tbody = document.getElementById("preview-tbody");
  tbody.innerHTML = APP.planValidation.slice(0, 12).map(r => `<tr>
    <td>${r.niveau}</td><td>${r.serie}</td><td>${r.rep}</td>
    <td>${r.xRef}</td><td>${r.yResponse}</td>
  </tr>`).join("");

  if (APP.planValidation.length > 12) {
    tbody.innerHTML += `<tr><td colspan="5" style="text-align:center;color:var(--text-muted);font-style:italic">… et ${APP.planValidation.length - 12} autres lignes</td></tr>`;
  }

  document.getElementById("preview-n").textContent = APP.planValidation.length;
  document.getElementById("preview-card").style.display = "block";
}

// ─── CALCULS ──────────────────────────────────────────────────────────────────
function setupCalcHandlers() {
  document.getElementById("btn-write-results").addEventListener("click", async () => {
    if (!APP.results) { toast("Aucun résultat à écrire", "warn"); return; }
    setBtnLoading("btn-write-results", true, "Écriture…");
    try {
      await ExcelBridge.writeAnalysisResults(APP.results, APP.config);
      toast("✅ Résultats écrits dans Excel", "info");
    } catch (e) { toast("Erreur : " + e.message, "err"); }
    setBtnLoading("btn-write-results", false, "📊 Écrire les résultats dans Excel");
  });
}

function renderCalcResults() {
  const { models, criteria, tolerances, outliers } = APP.results;

  // Modèles d'étalonnage
  const etaCard = document.getElementById("etalonnage-card");
  if (APP.config.methodType === "indirect" && Object.keys(models).length > 0) {
    etaCard.style.display = "block";
    document.getElementById("etalonnage-results").innerHTML = Object.entries(models)
      .map(([serie, m]) => `
        <div class="mono-block">
          <div style="font-weight:600;color:var(--amber-dim);margin-bottom:4px">${serie}</div>
          <div>Y = <span style="color:var(--navy-800)">${m.a1.toFixed(4)}</span>·X + <span style="color:var(--navy-800)">${m.a0.toFixed(4)}</span></div>
          <div style="color:var(--text-muted);font-size:10px">R² = ${m.r2.toFixed(6)} · r = ${m.r.toFixed(6)} · N = ${m.n}</div>
        </div>`).join("");
  } else {
    etaCard.style.display = "none";
  }

  // Critères
  const tbody = document.getElementById("criteria-tbody");
  tbody.innerHTML = criteria.map(c => {
    const biasClass = Math.abs(c.bRel) > APP.config.lambda * 100 ? "color:var(--invalid);font-weight:600" : "";
    return `<tr>
      <td>${c.niveau}</td>
      <td>${c.xMean.toFixed(4)}</td>
      <td>${c.zMean.toFixed(4)}</td>
      <td>${c.sr.toFixed(4)}</td>
      <td>${c.sB.toFixed(4)}</td>
      <td>${c.sFI.toFixed(4)}</td>
      <td>${c.cv.toFixed(2)}</td>
      <td style="${biasClass}">${c.bRel.toFixed(3)}</td>
      <td>${c.recouvMoy.toFixed(3)}</td>
    </tr>`;
  }).join("");

  // Intervalles de tolérance
  const tbody2 = document.getElementById("tolerance-tbody");
  tbody2.innerHTML = tolerances.map(t => `<tr>
    <td>${t.niveau}</td>
    <td>${t.xMean.toFixed(4)}</td>
    <td>${t.sIT.toFixed(4)}</td>
    <td>${t.ktol.toFixed(4)}</td>
    <td>${t.nu}</td>
    <td>${(t.ltbRel || 0).toFixed(3)}</td>
    <td>${(t.lthRel || 0).toFixed(3)}</td>
    <td>${t.laBasse.toFixed(1)}</td>
    <td>${t.laHaute.toFixed(1)}</td>
    <td><span class="status-${t.accept ? "valid" : "invalid"}">${t.accept ? "VALIDE" : "NON VALIDE"}</span></td>
  </tr>`).join("");

  // Aberrants
  const outlierEl = document.getElementById("outlier-results");
  outlierEl.innerHTML = outliers.map(o => {
    const cls = o.grubbs.suspect ? "aberrant" : "ok";
    return `<div class="outlier-row ${cls}">
      <strong>Niveau ${o.niveau}</strong> (X̄=${o.xMean.toFixed(3)}, n=${o.n}) —
      G=${o.grubbs.G} / G<sub>crit</sub>=${o.grubbs.Gcrit} →
      ${o.grubbs.suspect
        ? `⚠ <strong>ABERRANT détecté</strong> : ${o.grubbs.suspectVal?.toFixed(4)}`
        : "✅ Aucun aberrant"}
    </div>`;
  }).join("");

  document.getElementById("calc-empty").style.display   = "none";
  document.getElementById("calc-results").style.display = "block";
}

// ─── PROFIL D'EXACTITUDE ──────────────────────────────────────────────────────
function setupProfileHandlers() {
  document.getElementById("btn-insert-chart").addEventListener("click", async () => {
    if (!APP.results) { toast("Aucun profil à insérer", "warn"); return; }
    setBtnLoading("btn-insert-chart", true, "Insertion…");
    try {
      const sheet = await ExcelBridge.insertProfileChart(APP.results.tolerances, APP.config);
      toast(`✅ Graphique inséré → onglet "${sheet}"`, "info");
    } catch (e) { toast("Erreur : " + e.message, "err"); }
    setBtnLoading("btn-insert-chart", false, "📊 Insérer le graphique dans Excel");
  });
}

function renderProfileChart() {
  if (!APP.results?.tolerances?.length) return;

  const tolerances = APP.results.tolerances;
  const validity   = APP.results.validity;
  const laBasse    = tolerances[0].laBasse;
  const laHaute    = tolerances[0].laHaute;

  // Légende
  document.getElementById("profile-legend").innerHTML = `
    <div class="legend-item">
      <div class="legend-line" style="background:#F5A623;height:2px"></div>
      <span>Taux de recouvrement</span>
    </div>
    <div class="legend-item">
      <div class="legend-line" style="border-top:2px solid #1A3050;width:18px"></div>
      <span>LTB / LTH (β-expectation)</span>
    </div>
    <div class="legend-item">
      <div class="legend-line dashed" style="color:#EF4444;border-top-width:2px;width:18px"></div>
      <span>Limites d'acceptabilité (±${(APP.config.lambda*100).toFixed(0)}%)</span>
    </div>
    <div class="legend-item">
      <div class="legend-line dashed" style="color:#9ca3af;border-top-width:1px;width:18px"></div>
      <span>Référence 100%</span>
    </div>`;

  // Données
  const labels  = tolerances.map(t => `${t.xMean.toFixed(3)} ${APP.config.unite || ""}`);
  const recouv  = tolerances.map(t => +(t.recouvRel || 100).toFixed(3));
  const ltb     = tolerances.map(t => +(t.ltbRel || 0).toFixed(3));
  const lth     = tolerances.map(t => +(t.lthRel || 0).toFixed(3));
  const ref100  = tolerances.map(() => 100);
  const laLow   = tolerances.map(() => laBasse);
  const laHigh  = tolerances.map(() => laHaute);

  // Couleurs des points selon validité
  const recouvColors = tolerances.map(t => t.accept ? "#F5A623" : "#EF4444");

  // Détruire l'ancien graphique
  if (APP.profileChart) { APP.profileChart.destroy(); APP.profileChart = null; }

  const canvas = document.getElementById("chart-profile");
  APP.profileChart = new Chart(canvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Taux de recouvrement (%)",
          data: recouv,
          borderColor: "#F5A623",
          backgroundColor: "rgba(245,166,35,0.08)",
          pointBackgroundColor: recouvColors,
          pointRadius: 6, pointHoverRadius: 8,
          borderWidth: 2, tension: 0.3, fill: false,
        },
        {
          label: "LTB (%)",
          data: ltb,
          borderColor: "#1A3050",
          backgroundColor: "rgba(26,48,80,0.06)",
          pointRadius: 3, borderWidth: 1.5, tension: 0.3,
          fill: "+1",
        },
        {
          label: "LTH (%)",
          data: lth,
          borderColor: "#1A3050",
          pointRadius: 3, borderWidth: 1.5, tension: 0.3,
          fill: false,
        },
        {
          label: `L.Accept. basse (${laBasse.toFixed(0)}%)`,
          data: laLow,
          borderColor: "#EF4444",
          borderDash: [7, 4], pointRadius: 0, borderWidth: 1.5, fill: false,
        },
        {
          label: `L.Accept. haute (${laHaute.toFixed(0)}%)`,
          data: laHigh,
          borderColor: "#EF4444",
          borderDash: [7, 4], pointRadius: 0, borderWidth: 1.5, fill: false,
        },
        {
          label: "Référence 100%",
          data: ref100,
          borderColor: "rgba(140,160,185,0.4)",
          borderDash: [3, 3], pointRadius: 0, borderWidth: 1, fill: false,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { labels: { color: "#4A6080", font: { size: 10, family: "'IBM Plex Mono'" }, boxWidth: 14 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${ctx.raw?.toFixed(3)}%`,
          },
        },
      },
      scales: {
        x: {
          ticks: { color: "#4A6080", font: { size: 9, family: "'IBM Plex Mono'" } },
          grid:  { color: "rgba(180,200,220,0.3)" },
          title: { display: true, text: `Concentration de référence (${APP.config.unite || ""})`, color: "#4A6080", font: { size: 9 } },
        },
        y: {
          ticks: { color: "#4A6080", font: { size: 9, family: "'IBM Plex Mono'" }, callback: v => v + "%" },
          grid:  { color: "rgba(180,200,220,0.3)" },
          title: { display: true, text: "Taux de recouvrement (%)", color: "#4A6080", font: { size: 9 } },
          suggestedMin: Math.min(laBasse - 10, ...ltb) - 2,
          suggestedMax: Math.max(laHaute + 10, ...lth) + 2,
        },
      },
    },
  });

  // Verdict
  const { valid, partial, invalid, nValid, nTotal, pct, domain } = validity;
  const verdictEl = document.getElementById("verdict-box");
  if (valid) {
    verdictEl.className = "verdict-box verdict-valid";
    verdictEl.innerHTML = `<strong>✅ MÉTHODE VALIDE</strong> — Les ${nTotal} niveaux de concentration respectent les critères β-expectation.<br>
      Le procédé analytique produit des résultats dans les limites d'acceptabilité ±${(APP.config.lambda*100).toFixed(0)}% avec une probabilité de ${(APP.config.beta*100).toFixed(0)}%.`;
  } else if (partial) {
    verdictEl.className = "verdict-box verdict-partial";
    verdictEl.innerHTML = `<strong>⚠ MÉTHODE PARTIELLEMENT VALIDE</strong> — ${nValid}/${nTotal} niveaux validés (${pct}%).<br>
      La méthode est valide uniquement sur la plage : <strong>${domain?.min.toFixed(3)} – ${domain?.max.toFixed(3)} ${APP.config.unite || ""}</strong>.`;
  } else {
    verdictEl.className = "verdict-box verdict-invalid";
    verdictEl.innerHTML = `<strong>❌ MÉTHODE NON VALIDE</strong> — Aucun niveau ne respecte les critères β-expectation.<br>
      La méthode ne peut pas être utilisée en routine dans les conditions actuelles.`;
  }

  // Domaine de validité
  const domainEl = document.getElementById("validity-domain");
  domainEl.innerHTML = tolerances.map(t => `
    <div class="validity-row">
      <div class="validity-dot" style="background:${t.accept ? "#22C55E" : "#EF4444"}"></div>
      <span>Niveau ${t.niveau} — <strong>${t.xMean.toFixed(3)} ${APP.config.unite || ""}</strong> :
        Récouv.=${t.recouvRel.toFixed(2)}% | LTB=${(t.ltbRel||0).toFixed(2)}% / LTH=${(t.lthRel||0).toFixed(2)}%
        → <strong>${t.accept ? "VALIDE" : "NON VALIDE"}</strong>
      </span>
    </div>`).join("");

  document.getElementById("profile-empty").style.display    = "none";
  document.getElementById("profile-content").style.display  = "block";
}

// ─── IA ───────────────────────────────────────────────────────────────────────
function setupAIHandlers() {
  const aiMap = {
    "btn-ai-full":    async () => await callBackendInterpret("full"),
    "btn-ai-profile": async () => await callBackendInterpret("profile"),
    "btn-ai-outliers":async () => await callBackendInterpret("outliers"),
    "btn-ai-reco":    async () => await callBackendInterpret("recommendations"),
  };

  const btnLabels = {
    "btn-ai-full":     "✦ Diagnostic complet",
    "btn-ai-profile":  "◈ Interpréter le profil",
    "btn-ai-outliers": "⚠ Analyser les aberrants",
    "btn-ai-reco":     "📋 Recommandations",
  };

  Object.entries(aiMap).forEach(([btnId, fn]) => {
    document.getElementById(btnId).addEventListener("click", async () => {
      if (!APP.results) { toast("Calculez d'abord le profil d'exactitude", "warn"); return; }

      setBtnLoading(btnId, true, "Analyse IA…");
      document.getElementById("ai-result-card").style.display = "block";
      document.getElementById("ai-content").innerHTML = '<span class="spinner"></span> Gemini analyse votre profil d\'exactitude…';

      try {
        const result = await fn();
        APP.aiContent = result.text || result.items.join('\n');
        document.getElementById("ai-content").innerHTML = formatAIContent(result);
        toast("✅ Analyse IA terminée", "info");
      } catch (e) {
        document.getElementById("ai-content").innerHTML = `<span style="color:var(--invalid)">❌ ${e.message}</span>`;
        toast(e.message, "err");
      }
      setBtnLoading(btnId, false, btnLabels[btnId]);
    });
  });

  document.getElementById("btn-chat").addEventListener("click", handleChat);
  document.getElementById("chat-input").addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleChat(); }
  });
  document.querySelectorAll(".chip").forEach(chip => {
    chip.addEventListener("click", () => {
      document.getElementById("chat-input").value = chip.dataset.p;
      handleChat();
    });
  });
}

async function callBackendInterpret(promptType) {
  const response = await fetch(`${BACKEND_URL}/api/interpret?provider=gemini&prompt_type=${promptType}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      analysis_data: APP.results,
      config: APP.config
    })
  });

  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }

  return await response.json();
}

function formatAIContent(result) {
  if (result.source === 'llm') {
    return result.text.replace(/\n/g, '<br>');
  } else if (result.items) {
    return result.items.map(item => `<div>${item}</div>`).join('');
  }
  return 'Aucune réponse';
}

async function handleChat() {
  const input = document.getElementById("chat-input");
  const msg   = input.value.trim();
  if (!msg) return;

  input.value = "";
  appendChat("user", msg);
  const typingId = appendChat("assistant", '<span class="spinner"></span>');

  try {
    const context = APP.results ? {
      validity:    APP.results.validity,
      tolerances:  APP.results.tolerances.map(t => ({ niveau: t.niveau, xMean: t.xMean, recouvRel: t.recouvRel, ltbRel: t.ltbRel, lthRel: t.lthRel, accept: t.accept })),
      config:      { lambda: APP.config.lambda, beta: APP.config.beta, methode: APP.config.methode },
    } : {};

    const response = await fetch(`${BACKEND_URL}/api/chat`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        message: msg,
        context: context,
        provider: 'gemini'
      })
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    document.getElementById(typingId).innerHTML = data.response.replace(/\n/g, '<br>');
  } catch (e) {
    document.getElementById(typingId).innerHTML = `❌ ${e.message}`;
  }
}

let _chatN = 0;
function appendChat(role, html) {
  const id  = `chat-msg-${++_chatN}`;
  const box = document.getElementById("chat-messages");
  box.insertAdjacentHTML("beforeend", `
    <div class="chat-msg ${role}">
      <div class="chat-bubble" id="${id}">${html}</div>
    </div>`);
  box.scrollTop = box.scrollHeight;
  return id;
}

// ─── RAPPORT ──────────────────────────────────────────────────────────────────
function setupReportHandlers() {
  document.getElementById("btn-report-html").addEventListener("click", () => {
    if (!APP.results) { toast("Calculez d'abord le profil", "warn"); return; }

    const opts = {
      labo:     document.getElementById("rpt-labo").value,
      analyste: document.getElementById("rpt-analyste").value,
      ref:      document.getElementById("rpt-ref").value,
      version:  document.getElementById("rpt-version").value,
      params:      document.getElementById("rpt-params").checked,
      etalonnage:  document.getElementById("rpt-etalonnage").checked,
      criteria:    document.getElementById("rpt-criteria").checked,
      tolerance:   document.getElementById("rpt-tolerance").checked,
      profile:     document.getElementById("rpt-profile").checked,
      outliers:    document.getElementById("rpt-outliers").checked,
      ai:          document.getElementById("rpt-ai").checked,
    };

    const html = ReportGenerator.generateHTMLReport(
      { results: APP.results, config: APP.config, aiContent: APP.aiContent },
      opts
    );
    const fname = `Rapport_Validation_${(APP.config.methode || "methode").replace(/\s+/g, "_").slice(0, 30)}_${new Date().toISOString().slice(0,10)}.html`;
    ReportGenerator.downloadHTMLReport(html, fname);
    toast("✅ Rapport HTML téléchargé", "info");
    logReport("Rapport HTML généré", "ok");
  });


}

// ─── Utilitaires UI ───────────────────────────────────────────────────────────
function toast(msg, type = "info", duration = 3200) {
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${{ok:"✅",err:"❌",info:"ℹ",warn:"⚠"}[type]||"ℹ"}</span><span>${msg}</span>`;
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