/**
 * ================================================================
 * modules/excelBridge.js — Interface avec Office.js / Excel
 * Adapté pour le format de données du backend Python v2
 * ================================================================
 */
"use strict";

const ExcelBridge = (() => {

  // ─── Lecture des données ────────────────────────────────────────────────────

  async function detectUsedRange() {
    return Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const range = sheet.getUsedRange();
      range.load("address");
      await ctx.sync();
      // Retourner l'adresse qualifiée complète, ex: "Plan_Validation!A1:F28"
      return range.address;   // Excel renvoie déjà "SheetName!A1:Zn"
    });
  }

  /**
   * Résout le worksheet et la plage locale depuis une adresse qui peut être :
   *   - qualifiée  : "Plan_Validation!A1:F28"  ou  "Plan_Validation!A:F"
   *   - simple     : "A1:F28"   → feuille active
   * Retourne { sheet, localRange }
   */
  function _resolveSheetAndRange(ctx, address) {
    const bang = address.indexOf("!");
    if (bang !== -1) {
      // Supprimer les guillemets simples éventuels autour du nom de feuille
      const sheetPart = address.slice(0, bang).replace(/^'|'$/g, "");
      const rangePart = address.slice(bang + 1);
      const sheet = ctx.workbook.worksheets.getItem(sheetPart);
      return { sheet, localRange: rangePart };
    }
    return {
      sheet:      ctx.workbook.worksheets.getActiveWorksheet(),
      localRange: address,
    };
  }

  /**
   * Lit le plan de validation depuis Excel.
   * Supporte les adresses qualifiées : "Plan_Validation!A1:F28"
   * Format colonnes : Niveau | Série | Rép. | X réf. | Y réponse
   */
  async function readPlanValidation(rangeAddress) {
    return Excel.run(async ctx => {
      const { sheet, localRange } = _resolveSheetAndRange(ctx, rangeAddress.trim());
      const range = sheet.getRange(localRange);
      range.load("values");
      await ctx.sync();

      const values = range.values;
      if (!values || values.length < 2)
        throw new Error("Plan de validation vide ou incomplet (minimum 2 lignes)");

      // Détecter l'en-tête : la première cellule n'est pas un nombre
      const hasHeader = isNaN(parseFloat(String(values[0][0]).trim()));
      const dataRows  = hasHeader ? values.slice(1) : values;

      const plan = dataRows
        .filter(row => row.some(v => v !== "" && v !== null && v !== undefined))
        .map(row => ({
          niveau:    String(row[0] ?? "").trim(),
          serie:     String(row[1] ?? "").trim(),
          rep:       parseInt(row[2]) || 1,
          xRef:      parseFloat(row[3]),
          yResponse: parseFloat(row[4]),
        }))
        .filter(r => !isNaN(r.xRef) && !isNaN(r.yResponse)
                  && r.xRef > 0 && r.niveau !== "" && r.serie !== "");

      if (!plan.length)
        throw new Error(
          "Aucune ligne valide dans le plan de validation.\n" +
          "Vérifiez que les colonnes D (X réf.) et E (Y réponse) contiennent des nombres."
        );
      return plan;
    });
  }

  /**
   * Lit le plan d'étalonnage depuis Excel.
   * Supporte les adresses qualifiées : "Plan_Etalonnage!A1:E13"
   * Format colonnes : Niveau | Série | Rép. | X étalon | Y réponse
   */
  async function readPlanEtalonnage(rangeAddress) {
    return Excel.run(async ctx => {
      const { sheet, localRange } = _resolveSheetAndRange(ctx, rangeAddress.trim());
      const range = sheet.getRange(localRange);
      range.load("values");
      await ctx.sync();

      const values    = range.values;
      if (!values || values.length < 2)
        throw new Error("Plan d'étalonnage vide");

      const hasHeader = isNaN(parseFloat(String(values[0][0]).trim()));
      const dataRows  = hasHeader ? values.slice(1) : values;

      const plan = dataRows
        .filter(row => row.some(v => v !== "" && v !== null && v !== undefined))
        .map(row => ({
          niveau:    String(row[0] ?? "").trim(),
          serie:     String(row[1] ?? "").trim(),
          rep:       parseInt(row[2]) || 1,
          xEtalon:   parseFloat(row[3]),
          yResponse: parseFloat(row[4]),
        }))
        .filter(r => !isNaN(r.xEtalon) && !isNaN(r.yResponse)
                  && r.xEtalon > 0 && r.niveau !== "" && r.serie !== "");

      if (!plan.length)
        throw new Error(
          "Aucune ligne valide dans le plan d'étalonnage.\n" +
          "Vérifiez que les colonnes D (X étalon) et E (Y réponse) contiennent des nombres."
        );
      return plan;
    });
  }


  // ─── Génération du plan expérimental ────────────────────────────────────────

  async function generatePlanValidation(K, I, J, unite = "", methodType = "indirect") {
    return Excel.run(async ctx => {
      const wb        = ctx.workbook;
      const isDirect  = methodType === "direct";
      const sheetName = "Plan_Validation";

      let sheet = wb.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      if (sheet.isNullObject) {
        sheet = wb.worksheets.add(sheetName);
      } else {
        sheet.getUsedRangeOrNullObject().clear();
      }
      await ctx.sync();

      const headers = isDirect
        ? ["Niveau (k)", "Série (i)", "Répétition (j)",
           `Valeur référence X (${unite})`,
           `Concentration mesurée Z (${unite})`, "Remarque"]
        : ["Niveau (k)", "Série (i)", "Répétition (j)",
           `Valeur référence X (${unite})`,
           "Réponse instrumentale Y", "Unité"];

      const rows = [headers];
      for (let k = 1; k <= K; k++)
        for (let i = 1; i <= I; i++)
          for (let j = 1; j <= J; j++)
            rows.push([`N${k}`, `Jour ${i}`, j, "", "", ""]);

      const range = sheet.getRange(`A1:F${rows.length}`);
      range.values = rows;

      const hdr = sheet.getRange("A1:F1");
      hdr.format.fill.color = "#0B1929";
      hdr.format.font.color = "#F5A623";
      hdr.format.font.bold  = true;
      hdr.format.font.size  = 9;

      for (let r = 3; r <= rows.length; r += 2)
        sheet.getRange(`A${r}:F${r}`).format.fill.color = "#F0F3F8";

      sheet.getRange("D2").format.fill.color = "#FFFDE7";
      sheet.getRange("E2").format.fill.color = "#E8F5E9";

      ["A","B","C","D","E","F"].forEach((col, i) => {
        sheet.getRange(`${col}1`).format.columnWidth = [70, 70, 80, 160, 160, 80][i];
      });

      sheet.activate();
      await ctx.sync();
      // dataEndRow = rows.length (header + data), dernière ligne de données
      const dataEndRow = rows.length;
      return {
        sheetName,
        rows:          rows.length - 1,
        qualifiedRange: `${sheetName}!A1:F${dataEndRow}`,  // plage qualifiée pour auto-fill
      };
    });
  }

  async function generatePlanEtalonnage(I, niveaux = 2, J = 2, unite = "") {
    return Excel.run(async ctx => {
      const wb        = ctx.workbook;
      const sheetName = "Plan_Etalonnage";

      let sheet = wb.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      if (sheet.isNullObject) { sheet = wb.worksheets.add(sheetName); }
      else { sheet.getUsedRangeOrNullObject().clear(); }
      await ctx.sync();

      const headers = [
        "Niveau étalon (k')", "Série (i)", "Répétition (j')",
        `Concentration étalon X (${unite})`, "Réponse instrumentale Y",
      ];
      const rows = [headers];
      for (let k = 1; k <= niveaux; k++)
        for (let i = 1; i <= I; i++)
          for (let j = 1; j <= J; j++)
            rows.push([`E${k}`, `Jour ${i}`, j, "", ""]);

      const range = sheet.getRange(`A1:E${rows.length}`);
      range.values = rows;

      const hdr = sheet.getRange("A1:E1");
      hdr.format.fill.color = "#0B1929";
      hdr.format.font.color = "#F5A623";
      hdr.format.font.bold  = true;
      hdr.format.font.size  = 9;

      sheet.activate();
      await ctx.sync();
      return {
        sheetName,
        rows:          rows.length - 1,
        qualifiedRange: `${sheetName}!A1:E${rows.length}`,
      };
    });
  }


  // ─── Écriture des résultats backend ─────────────────────────────────────────

  /**
   * Écrit les résultats complets de l'API backend dans Excel.
   * Utilise le format de réponse v2 : criteria, tolerances, outliers, qualityScore…
   */
  async function writeAnalysisResults(results, config) {
    return Excel.run(async ctx => {
      const wb        = ctx.workbook;
      const sheetName = "Résultats_Calculs";
      let   sheet     = wb.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      if (sheet.isNullObject) { sheet = wb.worksheets.add(sheetName); }
      else { sheet.getUsedRangeOrNullObject().clear(); }
      await ctx.sync();
      sheet.tabColor = "#F5A623";

      const { criteria = [], tolerances = [], outliers = [], validity = {}, qualityScore = {} } = results;
      const lambda = config.lambda ?? config.lambdaVal ?? 0.10;

      let row = 1;

      // ── Titre ──────────────────────────────────────────────────────────────
      _writeRow(sheet, row++, ["PROFIL D'EXACTITUDE — RÉSULTATS"], "#0B1929", "#FFFFFF", true, 12);
      _writeRow(sheet, row++, [`Méthode : ${config.methode || "—"}  |  λ=±${(lambda*100).toFixed(0)}%  |  β=${(config.beta*100).toFixed(0)}%`], "#F0F3F8", "#4A6080", false, 9);
      row++;

      // ── Score de qualité ────────────────────────────────────────────────────
      if (qualityScore.overall !== undefined) {
        _writeRow(sheet, row++, ["SCORE DE QUALITÉ", `${qualityScore.overall}/100 (${qualityScore.label})`], "#122339", "#F5A623", true, 10);
        row++;
      }

      // ── Validité ────────────────────────────────────────────────────────────
      const validLabel = validity.valid ? "MÉTHODE VALIDE ✓" :
                         validity.partial ? "VALIDATION PARTIELLE ⚠" : "NON VALIDE ✗";
      _writeRow(sheet, row++,
        ["STATUT", validLabel, `${validity.nValid ?? 0}/${validity.nTotal ?? 0} niveaux (${validity.pct ?? 0}%)`],
        validity.valid ? "#166534" : validity.partial ? "#92400E" : "#991B1B", "#FFFFFF", true, 10
      );
      row++;

      // ── Critères justesse / fidélité ────────────────────────────────────────
      _writeRow(sheet, row++,
        ["CRITÈRES ISO 5725-2 — JUSTESSE ET FIDÉLITÉ"],
        "#0B1929", "#F5A623", true, 10);
      _writeRow(sheet, row++,
        ["Niveau", "X̄ réf.", "Z̄ retrouvée", "sr", "sB", "sFI", "CV%", "CVr%", "Biais%", "Récouv.%", "Shapiro-p"],
        "#122339", "#9BBDD6", true, 9);

      for (const c of criteria) {
        const biasFlag = Math.abs(c.bRel) > lambda * 100;
        _writeRow(sheet, row++, [
          c.niveau, _f(c.xMean), _f(c.zMean),
          _f(c.sr), _f(c.sB), _f(c.sFI),
          _f(c.cv, 2), _f(c.cvR, 2),
          _f(c.bRel, 3), _f(c.recouvMoy, 3),
          c.shapiro_p != null ? _f(c.shapiro_p, 4) : "—",
        ], biasFlag ? "#991B1B" : null, biasFlag ? "#FFFFFF" : null, false, 9);
      }
      row++;

      // ── Intervalles β-expectation ────────────────────────────────────────────
      _writeRow(sheet, row++,
        ["INTERVALLES β-EXPECTATION (Mee, 1984)"],
        "#0B1929", "#F5A623", true, 10);
      _writeRow(sheet, row++,
        ["Niveau", "X̄ réf.", "sIT", "k_tol", "ν", "LTB%", "LTH%", "L.A. basse%", "L.A. haute%", "Erreur totale%", "Statut"],
        "#122339", "#9BBDD6", true, 9);

      for (const t of tolerances) {
        const bgColor = t.accept ? "#F0FDF4" : "#FEF2F2";
        const fgColor = t.accept ? "#166534" : "#991B1B";
        _writeRow(sheet, row++, [
          t.niveau, _f(t.xMean), _f(t.sIT), _f(t.ktol, 4),
          t.nu, _f(t.ltbRel, 3), _f(t.lthRel, 3),
          _f(t.laBasse, 1), _f(t.laHaute, 1),
          _f(t.errorTotal, 3),
          t.accept ? "VALIDE" : "NON VALIDE",
        ], bgColor, fgColor, false, 9);
      }
      row++;

      // ── Aberrants ────────────────────────────────────────────────────────────
      _writeRow(sheet, row++,
        ["TEST DE GRUBBS (α=5%)"],
        "#0B1929", "#F5A623", true, 10);
      _writeRow(sheet, row++,
        ["Niveau", "X̄ réf.", "n", "G", "G_crit", "Statut", "Valeur suspecte"],
        "#122339", "#9BBDD6", true, 9);

      for (const o of outliers) {
        const isOk = !o.suspect;
        _writeRow(sheet, row++, [
          o.niveau, _f(o.xMean), o.n ?? o.n,
          _f(o.G, 4), _f(o.Gcrit, 4),
          isOk ? "OK" : o.classification?.toUpperCase() ?? "SUSPECT",
          isOk ? "—" : _f(o.suspectVal, 6),
        ], isOk ? null : "#FEF9C3", isOk ? null : "#92400E", false, 9);
      }

      // Largeurs colonnes
      const widths = [80, 90, 90, 80, 80, 90, 90, 90, 90, 100, 100];
      for (let i = 0; i < widths.length; i++) {
        try { sheet.getColumn(i + 1).width = widths[i]; } catch { /**/ }
      }

      sheet.activate();
      await ctx.sync();
      return sheetName;
    });
  }


  // ─── Insertion du graphique de profil ────────────────────────────────────────

  /**
   * Insère le profil d'exactitude dans Excel via Chart.js natif Office.
   * Si le backend a renvoyé une image base64, elle est insérée en image.
   */
  async function insertProfileChart(tolerances, config, chartBase64 = null) {
    return Excel.run(async ctx => {
      const wb        = ctx.workbook;
      const sheetName = "Profil_Exactitude";
      let   sheet     = wb.worksheets.getItemOrNullObject(sheetName);
      await ctx.sync();
      if (sheet.isNullObject) { sheet = wb.worksheets.add(sheetName); }
      else { sheet.getUsedRangeOrNullObject().clear(); }
      await ctx.sync();

      sheet.tabColor = "#22C55E";

      // Si on a une image base64 du backend, on l'insère directement
      if (chartBase64 && chartBase64.startsWith("data:image/png;base64,")) {
        const b64 = chartBase64.replace("data:image/png;base64,", "");
        sheet.getRange("A1").values = [["Profil d'exactitude — généré par le backend"]];
        // Office.js ne supporte pas l'insertion d'image via base64 directement
        // → On écrit les données et on crée un graphique natif
      }

      // Écrire les données source pour le graphique natif
      const headers = ["Concentration", "Recouvrement (%)", "LTB (%)", "LTH (%)",
                       `L.A. basse (${tolerances[0]?.laBasse?.toFixed(0) || 90}%)`,
                       `L.A. haute (${tolerances[0]?.laHaute?.toFixed(0) || 110}%)`];
      const dataRows = tolerances.map(t => [
        t.xMean,
        parseFloat((t.recouvRel || 100).toFixed(4)),
        parseFloat((t.ltbRel || 0).toFixed(4)),
        parseFloat((t.lthRel || 0).toFixed(4)),
        parseFloat((t.laBasse || 90).toFixed(1)),
        parseFloat((t.laHaute || 110).toFixed(1)),
      ]);

      const allRows = [headers, ...dataRows];
      const dataRange = sheet.getRange(`A1:F${allRows.length}`);
      dataRange.values = allRows;

      // Style en-tête
      const hdr = sheet.getRange("A1:F1");
      hdr.format.fill.color = "#0B1929";
      hdr.format.font.color = "#F5A623";
      hdr.format.font.bold  = true;
      hdr.format.font.size  = 9;

      // Graphique natif Excel
      const chartDataRange = sheet.getRange(`A1:F${allRows.length}`);
      const chart = sheet.charts.add(Excel.ChartType.line, chartDataRange, Excel.ChartSeriesBy.columns);
      chart.title.text = `Profil d'exactitude — λ=±${((config.lambda || 0.10)*100).toFixed(0)}% | β=${((config.beta || 0.80)*100).toFixed(0)}%`;
      chart.title.font.size  = 11;
      chart.title.font.color = "#0B1929";
      chart.setPosition(`H1`, `R20`);

      sheet.activate();
      await ctx.sync();
      return sheetName;
    });
  }


  // ─── Feuilles de saisie (templates) ─────────────────────────────────────────

  async function generateBlankTemplates({ K = 3, I = 3, J = 3, unite = "", methodType = "indirect" }) {
    const isDirect = methodType === "direct";
    const resV = await generatePlanValidation(K, I, J, unite, methodType);
    let   resE = null;
    if (!isDirect) {
      resE = await generatePlanEtalonnage(I, 2, 2, unite);
    }
    await _generateParamsSheet(K, I, J, unite, methodType);
    return { sheetV: resV.sheetName, sheetE: resE?.sheetName || null, sheetP: "Paramètres" };
  }

  async function _generateParamsSheet(K, I, J, unite, methodType) {
    return Excel.run(async ctx => {
      const wb  = ctx.workbook;
      let sheet = wb.worksheets.getItemOrNullObject("Paramètres");
      await ctx.sync();
      if (sheet.isNullObject) { sheet = wb.worksheets.add("Paramètres"); }
      else { sheet.getUsedRangeOrNullObject().clear(); }
      await ctx.sync();
      sheet.tabColor = "#22C55E";

      _writeRow(sheet, 1, ["PARAMÈTRES DE VALIDATION"], "#0B1929", "#F5A623", true, 12);
      const params = [
        ["Paramètre",           "Valeur",      "Explication"],
        ["Méthode analytique",  "",            "Nom complet de la méthode à valider"],
        ["Matériau",            "",            "Nature du matériau (MRC, ajout dosé…)"],
        ["Unité",               unite || "",   "ex: mg/L, µg/kg, %"],
        ["Type de méthode",     methodType,   "direct | indirect"],
        ["Limite λ (%)",        "10",          "Limite d'acceptabilité. Ex: 10 pour ±10%"],
        ["Proportion β (%)",    "80",          "Min recommandé: 80%"],
        ["Niveaux K",           String(K),     "Min: 3"],
        ["Séries I",            String(I),     "Min: 3"],
        ["Répétitions J",       String(J),     "Min: 2"],
        ["Modèle étalonnage",   "linear",      "linear | origin | quad | auto"],
        ["Framework normatif",  "iso5725",     "iso5725 | ichq2 | sfstp"],
      ];

      const pRange = sheet.getRange(`A3:C${3 + params.length - 1}`);
      pRange.values = params;
      sheet.getRange("A3:C3").format.fill.color = "#0B1929";
      sheet.getRange("A3:C3").format.font.color = "#F5A623";
      sheet.getRange("A3:C3").format.font.bold  = true;
      sheet.getRange(`B4:B${3 + params.length - 1}`).format.fill.color = "#FFFDE7";

      [240, 120, 360].forEach((w, i) =>
        sheet.getRange(String.fromCharCode(65 + i) + "1").format.columnWidth = w
      );

      sheet.activate();
      await ctx.sync();
    });
  }


  // ─── Helpers privés ─────────────────────────────────────────────────────────

  function _writeRow(sheet, rowNum, values, bgColor = null, fontColor = null, bold = false, fontSize = 9) {
    const endCol = String.fromCharCode(64 + values.length);
    const range  = sheet.getRange(`A${rowNum}:${endCol}${rowNum}`);
    range.values = [values];
    if (bgColor)    range.format.fill.color  = bgColor;
    if (fontColor)  range.format.font.color  = fontColor;
    if (bold)       range.format.font.bold   = bold;
    if (fontSize)   range.format.font.size   = fontSize;
  }

  function _f(val, dec = 4) {
    if (val == null || val === undefined) return "—";
    return parseFloat(val).toFixed(dec);
  }


  // ─── API publique ─────────────────────────────────────────────────────────────

  return {
    detectUsedRange,
    readPlanValidation,
    readPlanEtalonnage,
    generatePlanValidation,
    generatePlanEtalonnage,
    writeAnalysisResults,
    insertProfileChart,
    generateBlankTemplates,
  };
})();

window.ExcelBridge = ExcelBridge;