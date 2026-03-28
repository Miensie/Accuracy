/**
 * ================================================================
 * modules/reportGenerator.js — Génération des rapports
 * HTML local + PDF via backend Python (reportlab)
 * Compatible avec la réponse API v2 (criteria, tolerances, qualityScore…)
 * ================================================================
 */
"use strict";

const ReportGenerator = (() => {

  // ─── Rapport PDF via backend ─────────────────────────────────────────────────

  /**
   * Télécharge le rapport PDF généré par le backend.
   * Passe config + résultats au endpoint /api/report/pdf
   */
  async function downloadPDFReport(results, config) {
    if (!results) throw new Error("Résultats manquants");
    await ApiClient.downloadPDF(results, _normalizeConfig(config));
  }


  // ─── Rapport HTML local ──────────────────────────────────────────────────────

  function generateHTMLReport(data, opts = {}) {
    const { results, config, aiContent = "" } = data;
    if (!results) throw new Error("Résultats manquants");

    const { criteria = [], tolerances = [], outliers = [],
            validity = {}, qualityScore = {}, normativeChecks = [],
            interpretation = [], models = {} } = results;

    const lambda    = config.lambda ?? config.lambdaVal ?? 0.10;
    const beta      = config.beta   ?? 0.80;
    const unite     = config.unite  || "";
    const methode   = config.methode || "—";
    const dateStr   = new Date().toLocaleDateString("fr-FR", { day:"2-digit", month:"long", year:"numeric" });

    const statusColor = validity.valid ? "#166534" : validity.partial ? "#92400E" : "#991B1B";
    const statusBg    = validity.valid ? "#F0FDF4" : validity.partial ? "#FFFBEB" : "#FEF2F2";
    const statusText  = validity.valid
      ? `✓ MÉTHODE VALIDE (${validity.nValid}/${validity.nTotal} niveaux)`
      : validity.partial
        ? `⚠ VALIDATION PARTIELLE (${validity.nValid}/${validity.nTotal} niveaux)`
        : `✗ MÉTHODE NON VALIDÉE (0/${validity.nTotal} niveaux)`;

    return `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Rapport de validation — ${methode}</title>
<style>
  :root {
    --navy:#0B1929; --amber:#F5A623; --valid:#166534; --invalid:#991B1B;
    --warn:#92400E; --info:#1E40AF; --border:#D4DDE8; --bg:#F8F9FC;
  }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:"Segoe UI",Arial,sans-serif; font-size:12px; color:#1A1A2E; background:#fff; }
  .page { max-width:960px; margin:0 auto; padding:32px 40px; }

  /* Header */
  .rpt-header { background:var(--navy); color:#fff; padding:20px 24px; border-radius:8px;
                display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; }
  .rpt-title  { font-size:18px; font-weight:700; letter-spacing:.05em; }
  .rpt-title span { color:var(--amber); }
  .rpt-meta   { font-size:10px; color:#9BBDD6; text-align:right; line-height:1.8; }

  /* Statut */
  .status-banner { padding:14px 18px; border-radius:6px; font-size:14px; font-weight:700;
                   margin-bottom:18px; display:flex; align-items:center; gap:12px;
                   background:${statusBg}; color:${statusColor}; border-left:4px solid ${statusColor}; }

  /* Score */
  .score-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:18px; }
  .score-card { background:var(--bg); border:1px solid var(--border); border-radius:6px;
                padding:10px 12px; text-align:center; }
  .score-val  { font-size:22px; font-weight:700; color:var(--navy); }
  .score-lbl  { font-size:9px; color:#4A6080; text-transform:uppercase; letter-spacing:.08em; margin-top:2px; }

  /* Sections */
  .section    { margin-bottom:22px; }
  .section-h  { font-size:13px; font-weight:700; color:var(--navy); border-bottom:2px solid var(--amber);
                padding-bottom:4px; margin-bottom:10px; letter-spacing:.03em; }

  /* Tables */
  table { width:100%; border-collapse:collapse; font-size:10px; }
  th    { background:var(--navy); color:#9BBDD6; font-size:9px; padding:6px 8px; text-align:left;
          font-weight:600; letter-spacing:.06em; }
  td    { padding:5px 8px; border-bottom:1px solid var(--border); color:#2D3748; }
  tr:nth-child(even) td { background:var(--bg); }
  .valid   { color:#166534; font-weight:700; }
  .invalid { color:#991B1B; font-weight:700; }
  .warn    { color:#92400E; font-weight:700; }

  /* Interprétation */
  .interp-item { padding:7px 10px; border-radius:4px; margin-bottom:6px;
                  font-size:11px; line-height:1.55; }
  .interp-success  { background:#F0FDF4; border-left:3px solid #22C55E; color:#166534; }
  .interp-info     { background:#EFF6FF; border-left:3px solid #3B82F6; color:#1E40AF; }
  .interp-warning  { background:#FFFBEB; border-left:3px solid #F59E0B; color:#92400E; }
  .interp-critical { background:#FEF2F2; border-left:3px solid #EF4444; color:#991B1B; }

  /* AI */
  .ai-box { background:#F0F3F8; border:1px solid var(--border); border-radius:6px;
             padding:14px 16px; font-size:11px; line-height:1.8; white-space:pre-wrap; }

  /* Footer */
  .rpt-footer { margin-top:32px; padding-top:12px; border-top:1px solid var(--border);
                font-size:9px; color:#8BA3BE; text-align:center; }

  @media print { .page { padding:16px 20px; } }
</style>
</head>
<body>
<div class="page">

<!-- Header -->
<div class="rpt-header">
  <div>
    <div class="rpt-title">ACCURACY <span>PROFILE</span></div>
    <div style="font-size:10px;color:#9BBDD6;margin-top:4px">Rapport de validation analytique · ISO 5725-2 · Feinberg (2010)</div>
  </div>
  <div class="rpt-meta">
    <div><strong>${opts.labo || "Laboratoire"}</strong></div>
    <div>${opts.analyste || "—"}</div>
    <div>Réf. : ${opts.ref || "—"} · v${opts.version || "1.0"}</div>
    <div>${dateStr}</div>
  </div>
</div>

<!-- Méthode -->
<div class="section">
  <div class="section-h">Méthode analysée</div>
  <table>
    <tr><th>Méthode</th><th>Matériau</th><th>Unité</th><th>Type</th><th>β</th><th>λ</th></tr>
    <tr>
      <td>${methode}</td>
      <td>${config.materiau || "—"}</td>
      <td>${unite || "—"}</td>
      <td>${config.methodType || "—"}</td>
      <td>${(beta*100).toFixed(0)}%</td>
      <td>±${(lambda*100).toFixed(0)}%</td>
    </tr>
  </table>
</div>

<!-- Statut -->
<div class="status-banner">${statusText}</div>

${qualityScore.overall != null ? `
<!-- Score de qualité -->
<div class="section">
  <div class="section-h">Score de qualité</div>
  <div class="score-grid">
    <div class="score-card"><div class="score-val">${qualityScore.overall?.toFixed(1)}</div><div class="score-lbl">Score global / 100<br><strong>${qualityScore.label || ""}</strong></div></div>
    <div class="score-card"><div class="score-val">${qualityScore.justesse?.toFixed(0)}</div><div class="score-lbl">Justesse</div></div>
    <div class="score-card"><div class="score-val">${qualityScore.fidelite?.toFixed(0)}</div><div class="score-lbl">Fidélité</div></div>
    <div class="score-card"><div class="score-val">${qualityScore.profil?.toFixed(0)}</div><div class="score-lbl">Profil</div></div>
    <div class="score-card"><div class="score-val">${qualityScore.normalite?.toFixed(0)}</div><div class="score-lbl">Normalité</div></div>
    <div class="score-card"><div class="score-val">${qualityScore.homogeneite?.toFixed(0)}</div><div class="score-lbl">Homogénéité</div></div>
  </div>
</div>` : ""}

${opts.etalonnage && Object.keys(models).length > 0 ? `
<!-- Modèles d'étalonnage -->
<div class="section">
  <div class="section-h">Modèles d'étalonnage</div>
  <table>
    <tr><th>Série</th><th>Modèle</th><th>a₀</th><th>a₁</th><th>R²</th><th>r</th><th>N</th></tr>
    ${Object.entries(models).map(([s,m]) => `
    <tr>
      <td>${s}</td>
      <td>${m.modelType || "linear"}</td>
      <td>${(m.a0||0).toFixed(6)}</td>
      <td>${(m.a1||0).toFixed(6)}</td>
      <td>${(m.r2||0).toFixed(6)}</td>
      <td>${(m.r||0).toFixed(6)}</td>
      <td>${m.n || "—"}</td>
    </tr>`).join("")}
  </table>
</div>` : ""}

${opts.criteria !== false ? `
<!-- Critères ISO 5725-2 -->
<div class="section">
  <div class="section-h">Critères de justesse et fidélité (ISO 5725-2)</div>
  <table>
    <tr><th>Niveau</th><th>X̄ réf.</th><th>Z̄ ret.</th><th>sr</th><th>sB</th><th>sFI</th><th>CV%</th><th>CVr%</th><th>Biais%</th><th>Récouv.%</th></tr>
    ${criteria.map(c => {
      const biasOk = Math.abs(c.bRel) <= lambda * 100;
      return `<tr>
        <td><strong>${c.niveau}</strong></td>
        <td>${(c.xMean||0).toFixed(4)}</td>
        <td>${(c.zMean||0).toFixed(4)}</td>
        <td>${(c.sr||0).toFixed(4)}</td>
        <td>${(c.sB||0).toFixed(4)}</td>
        <td>${(c.sFI||0).toFixed(4)}</td>
        <td>${(c.cv||0).toFixed(2)}</td>
        <td>${(c.cvR||0).toFixed(2)}</td>
        <td class="${biasOk ? "" : "invalid"}">${(c.bRel||0).toFixed(3)}</td>
        <td>${(c.recouvMoy||0).toFixed(3)}</td>
      </tr>`;
    }).join("")}
  </table>
</div>` : ""}

${opts.tolerance !== false ? `
<!-- Intervalles β-expectation -->
<div class="section">
  <div class="section-h">Intervalles β-expectation (Mee, 1984)</div>
  <table>
    <tr><th>Niveau</th><th>X̄ réf.</th><th>sIT</th><th>k_tol</th><th>ν</th><th>LTB%</th><th>LTH%</th><th>L.A.basse%</th><th>L.A.haute%</th><th>Err. tot.%</th><th>Statut</th></tr>
    ${tolerances.map(t => `<tr>
      <td><strong>${t.niveau}</strong></td>
      <td>${(t.xMean||0).toFixed(4)}</td>
      <td>${(t.sIT||0).toFixed(4)}</td>
      <td>${(t.ktol||0).toFixed(4)}</td>
      <td>${t.nu}</td>
      <td>${(t.ltbRel||0).toFixed(3)}</td>
      <td>${(t.lthRel||0).toFixed(3)}</td>
      <td>${(t.laBasse||90).toFixed(1)}</td>
      <td>${(t.laHaute||110).toFixed(1)}</td>
      <td>${(t.errorTotal||0).toFixed(3)}</td>
      <td class="${t.accept ? "valid" : "invalid"}">${t.accept ? "VALIDE" : "NON VALIDE"}</td>
    </tr>`).join("")}
  </table>
</div>` : ""}

${opts.outliers !== false ? `
<!-- Aberrants -->
<div class="section">
  <div class="section-h">Détection des aberrants — Test de Grubbs (α=5%)</div>
  <table>
    <tr><th>Niveau</th><th>X̄ réf.</th><th>n</th><th>G</th><th>G_crit</th><th>Statut</th><th>Valeur suspecte</th></tr>
    ${outliers.map(o => `<tr>
      <td><strong>${o.niveau}</strong></td>
      <td>${(o.xMean||0).toFixed(4)}</td>
      <td>${o.n}</td>
      <td>${(o.G||0).toFixed(4)}</td>
      <td>${(o.Gcrit||0).toFixed(4)}</td>
      <td class="${o.suspect ? "warn" : "valid"}">${o.suspect ? (o.classification||"SUSPECT").toUpperCase() : "OK"}</td>
      <td>${o.suspect ? (o.suspectVal||0).toFixed(6) : "—"}</td>
    </tr>`).join("")}
  </table>
</div>` : ""}

${normativeChecks.length > 0 ? `
<!-- Vérifications normatives -->
<div class="section">
  <div class="section-h">Vérifications normatives</div>
  ${normativeChecks.map(item => `
    <div class="interp-item interp-${item.severity}">
      <strong>[${item.category}]</strong> ${item.message}
    </div>`).join("")}
</div>` : ""}

${interpretation.length > 0 ? `
<!-- Interprétation par règles -->
<div class="section">
  <div class="section-h">Interprétation automatique</div>
  ${interpretation.map(item => `
    <div class="interp-item interp-${item.severity}">
      <strong>[${item.category}]</strong> ${item.message}
    </div>`).join("")}
</div>` : ""}

${opts.ai && aiContent ? `
<!-- Interprétation IA -->
<div class="section">
  <div class="section-h">Interprétation par IA</div>
  <div class="ai-box">${_escapeHtml(aiContent)}</div>
</div>` : ""}

<!-- Footer -->
<div class="rpt-footer">
  Rapport généré par Accuracy Profile v2 · ${dateStr} ·
  Basé sur : ISO 5725-2, ICH Q2(R1), Feinberg M. (2010), Mee R.W. (1984)
</div>

</div>
</body>
</html>`;
  }

  function downloadHTMLReport(html, filename) {
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }


  // ─── Helpers ─────────────────────────────────────────────────────────────────

  function _normalizeConfig(config) {
    return {
      methode:    config.methode    || "",
      materiau:   config.materiau   || "",
      unite:      config.unite      || "",
      methodType: config.methodType || "indirect",
      beta:       config.beta       ?? 0.80,
      lambdaVal:  config.lambda     ?? config.lambdaVal ?? 0.10,
      laboratoire: config.laboratoire || "",
      analyste:    config.analyste   || "",
    };
  }

  function _escapeHtml(str) {
    return String(str)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>")
      .replace(/\n/g,"<br>");
  }


  // ─── API publique ────────────────────────────────────────────────────────────

  return {
    generateHTMLReport,
    downloadHTMLReport,
    downloadPDFReport,
  };
})();

window.ReportGenerator = ReportGenerator;