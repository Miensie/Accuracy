"""
================================================================
services/statistics.py — Calculs statistiques complets
ISO 5725-2 · Mee (1984) · Tests de Grubbs, Shapiro, Levene
================================================================
"""
from __future__ import annotations
import logging
from collections import defaultdict
from typing import List, Dict, Any, Optional, Tuple

import numpy as np
from scipy import stats
from scipy.stats import shapiro, levene, bartlett, f as f_dist
import statsmodels.api as sm

logger = logging.getLogger(__name__)


# ─── Statistiques descriptives de base ───────────────────────────────────────

def basic_stats(data: List[float], alpha: float = 0.05) -> Dict[str, float]:
    """
    Calcule les statistiques descriptives complètes pour un vecteur.

    Returns
    -------
    dict avec : n, mean, median, std, variance, cv, min, max, range,
                ci_low, ci_high, sem, skewness, kurtosis
    """
    arr = np.array(data, dtype=float)
    n   = len(arr)

    if n == 0:
        return {}

    mean     = float(np.mean(arr))
    std      = float(np.std(arr, ddof=1)) if n > 1 else 0.0
    variance = std ** 2
    sem      = std / np.sqrt(n) if n > 1 else 0.0
    cv       = (std / abs(mean)) * 100 if mean != 0 else 0.0

    # Intervalle de confiance t de Student
    if n > 1:
        t_crit = stats.t.ppf(1 - alpha / 2, df=n - 1)
        ci_low  = mean - t_crit * sem
        ci_high = mean + t_crit * sem
    else:
        ci_low = ci_high = mean

    return {
        "n":        n,
        "mean":     round(mean,     8),
        "median":   round(float(np.median(arr)), 8),
        "std":      round(std,      8),
        "variance": round(variance, 8),
        "cv":       round(cv,       4),
        "min":      round(float(arr.min()), 8),
        "max":      round(float(arr.max()), 8),
        "range":    round(float(arr.max() - arr.min()), 8),
        "ci_low":   round(ci_low,   8),
        "ci_high":  round(ci_high,  8),
        "sem":      round(sem,      8),
        "skewness": round(float(stats.skew(arr)), 4),
        "kurtosis": round(float(stats.kurtosis(arr)), 4),
    }


# ─── Régression linéaire ──────────────────────────────────────────────────────

def linear_regression(x: np.ndarray, y: np.ndarray) -> Dict[str, Any]:
    """
    Régression linéaire y = a0 + a1·x par OLS.
    Retourne coefficients, R², résidus, RMSE, SEy, p-values.
    """
    n = len(x)
    if n < 2:
        return {"a0": 0.0, "a1": 0.0, "r2": 0.0, "r": 0.0, "n": n,
                "modelType": "linear"}

    slope, intercept, r_val, p_val, stderr = stats.linregress(x, y)
    y_pred    = intercept + slope * x
    residuals = (y - y_pred).tolist()
    rmse      = float(np.sqrt(np.mean((y - y_pred) ** 2)))
    sse       = float(np.sum((y - y_pred) ** 2))
    sey       = float(np.sqrt(sse / max(1, n - 2)))

    return {
        "a0":       round(float(intercept), 8),
        "a1":       round(float(slope),     8),
        "r2":       round(float(r_val ** 2), 8),
        "r":        round(float(r_val),      8),
        "p_value":  round(float(p_val),      6),
        "stderr":   round(float(stderr),     8),
        "n":        n,
        "residuals": [round(v, 6) for v in residuals],
        "rmse":     round(rmse, 6),
        "sey":      round(sey, 6),
        "modelType": "linear",
    }


def origin_regression(x: np.ndarray, y: np.ndarray) -> Dict[str, Any]:
    """Régression par l'origine y = a1·x (sans ordonnée)."""
    n  = len(x)
    a1 = float(np.dot(x, y) / np.dot(x, x)) if np.dot(x, x) != 0 else 0.0
    y_pred    = a1 * x
    residuals = (y - y_pred).tolist()
    ss_res    = float(np.sum((y - y_pred) ** 2))
    ss_tot    = float(np.sum((y - np.mean(y)) ** 2))
    r2        = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0
    rmse      = float(np.sqrt(np.mean((y - y_pred) ** 2)))

    return {
        "a0":       0.0,
        "a1":       round(a1,   8),
        "r2":       round(r2,   8),
        "r":        round(float(np.sqrt(max(0, r2))), 8),
        "n":        n,
        "residuals": [round(v, 6) for v in residuals],
        "rmse":     round(rmse, 6),
        "modelType": "origin",
    }


def quadratic_regression(x: np.ndarray, y: np.ndarray) -> Dict[str, Any]:
    """Régression quadratique y = a0 + a1·x + a2·x² via numpy polyfit."""
    n    = len(x)
    coef = np.polyfit(x, y, 2)
    a2, a1, a0 = coef
    y_pred    = np.polyval(coef, x)
    residuals = (y - y_pred).tolist()
    ss_res    = float(np.sum((y - y_pred) ** 2))
    ss_tot    = float(np.sum((y - np.mean(y)) ** 2))
    r2        = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0
    rmse      = float(np.sqrt(np.mean((y - y_pred) ** 2)))

    return {
        "a0":       round(float(a0), 8),
        "a1":       round(float(a1), 8),
        "a2":       round(float(a2), 8),
        "r2":       round(r2, 8),
        "r":        round(float(np.sqrt(max(0, r2))), 8),
        "n":        n,
        "residuals": [round(v, 6) for v in residuals],
        "rmse":     round(rmse, 6),
        "modelType": "quad",
    }


def select_best_model(x: np.ndarray, y: np.ndarray) -> Dict[str, Any]:
    """Choisit le meilleur modèle (linéaire vs quadratique) par critère AIC."""
    lin  = linear_regression(x, y)
    quad = quadratic_regression(x, y)
    # AIC simplifié : n·ln(SSR/n) + 2k
    n = len(x)
    ssr_lin  = float(np.sum(np.array(lin["residuals"]) ** 2))
    ssr_quad = float(np.sum(np.array(quad["residuals"]) ** 2))
    aic_lin  = n * np.log(ssr_lin / n + 1e-14) + 2 * 2
    aic_quad = n * np.log(ssr_quad / n + 1e-14) + 2 * 3
    return quad if aic_quad < aic_lin - 2 else lin


# ─── Modèles d'étalonnage ─────────────────────────────────────────────────────

def compute_calibration_models(
    etalonnage_rows: List[Dict[str, Any]],
    model_type: str = "linear"
) -> Dict[str, Dict[str, Any]]:
    """
    Calcule un modèle d'étalonnage par série.
    Supporte: 'linear', 'origin', 'quad', 'auto'.
    """
    by_serie: Dict[str, Dict] = defaultdict(lambda: {"x": [], "y": []})
    for row in etalonnage_rows:
        serie = row.get("serie") or row.get("serie_id", "S1")
        by_serie[serie]["x"].append(row.get("xEtalon") or row.get("concentration", 0))
        by_serie[serie]["y"].append(row.get("yResponse") or row.get("measured", 0))

    models = {}
    for serie, data in by_serie.items():
        x = np.array(data["x"], dtype=float)
        y = np.array(data["y"], dtype=float)

        if model_type == "origin":
            m = origin_regression(x, y)
        elif model_type == "quad":
            m = quadratic_regression(x, y)
        elif model_type == "auto":
            m = select_best_model(x, y)
        else:  # linear (par défaut)
            m = linear_regression(x, y)

        m["serie"]    = serie
        m["x_data"]   = x.tolist()
        m["y_data"]   = y.tolist()
        models[serie] = m

    return models


# ─── Concentrations retrouvées ────────────────────────────────────────────────

def compute_found_concentrations(
    validation_rows: List[Dict[str, Any]],
    models: Dict[str, Dict[str, Any]],
    method_type: str = "indirect"
) -> List[Dict[str, Any]]:
    """
    Calcule les concentrations retrouvées Z par rétro-prédiction (méthode
    indirecte) ou lecture directe (méthode directe).
    """
    result = []
    warnings = []

    for row in validation_rows:
        x_ref     = row.get("xRef") or row.get("reference", 0)
        y_raw     = row.get("yResponse") or row.get("measured", 0)
        niveau    = row.get("niveau", "?")
        serie     = row.get("serie", "S1")
        rep       = row.get("rep", 1)

        if method_type == "direct":
            z = y_raw
        else:
            m = models.get(serie)
            if m is None:
                warnings.append(f"Pas de modèle pour la série '{serie}' — ligne ignorée")
                continue
            a1 = m.get("a1", 0)
            a0 = m.get("a0", 0)
            a2 = m.get("a2", None)
            if abs(a1) < 1e-14:
                warnings.append(f"Pente nulle pour '{serie}' — ligne ignorée")
                continue
            if a2 is not None:
                # Racine positive de a2·z² + a1·z + (a0 - y) = 0
                disc = a1 ** 2 - 4 * a2 * (a0 - y_raw)
                if disc < 0:
                    z = (y_raw - a0) / a1  # Fallback linéaire
                else:
                    z1 = (-a1 + np.sqrt(disc)) / (2 * a2)
                    z2 = (-a1 - np.sqrt(disc)) / (2 * a2)
                    # Choisir la racine la plus proche de x_ref
                    z = z1 if abs(z1 - x_ref) <= abs(z2 - x_ref) else z2
            else:
                z = (y_raw - a0) / a1

        b_abs = z - x_ref
        b_rel = (b_abs / x_ref) * 100 if x_ref != 0 else None

        result.append({
            "niveau":    niveau,
            "serie":     serie,
            "rep":       rep,
            "xRef":      round(x_ref, 8),
            "yResponse": round(y_raw, 8),
            "zRetrouvee": round(z, 8),
            "bAbs":      round(b_abs, 8),
            "bRel":      round(b_rel, 4) if b_rel is not None else None,
        })

    if warnings:
        logger.warning("compute_found_concentrations: %s", "; ".join(warnings))

    return result


# ─── Critères ISO 5725-2 ─────────────────────────────────────────────────────

def compute_criteria(found: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Calcule les critères de justesse et de fidélité par niveau selon
    ISO 5725-2 (modèle à deux composantes de variance : répétabilité
    et variance inter-séries).

    Retourne une liste triée par xMean croissant.
    """
    by_niveau: Dict[str, list] = defaultdict(list)
    for r in found:
        by_niveau[r["niveau"]].append(r)

    criteria = []
    for niveau, rows in by_niveau.items():
        by_serie: Dict[str, list] = defaultdict(list)
        for r in rows:
            by_serie[r["serie"]].append(r)

        series = list(by_serie.values())
        I = len(series)
        # On utilise le min de J par série pour robustesse
        J_list = [len(s) for s in series]
        J_min  = min(J_list)
        J      = J_min  # Plan équilibré supposé
        N      = sum(J_list)

        x_mean  = np.mean([r["xRef"] for r in rows])
        z_vals  = np.array([r["zRetrouvee"] for r in rows])
        z_mean  = float(np.mean(z_vals))

        serie_means = np.array([np.mean([r["zRetrouvee"] for r in s]) for s in series])

        # Sommes des carrés
        sce_r = sum(
            (r["zRetrouvee"] - serie_means[i]) ** 2
            for i, s in enumerate(series) for r in s
        )
        sce_b = sum(len(s) * (zm - z_mean) ** 2
                    for s, zm in zip(series, serie_means))

        ddr   = sum(len(s) - 1 for s in series)
        ddb   = I * (J - 1)
        sr2   = sce_r / ddr if ddr > 0 else 0.0
        sb2   = max(0.0, (sce_b /(I-1) - sr2) / J) if ddb > 0 else 0.0
        sfi2  = sr2 + sb2

        sr   = float(np.sqrt(sr2))
        sB   = float(np.sqrt(sb2))
        sFI  = float(np.sqrt(sfi2))
        cv   = (sFI / abs(z_mean)) * 100 if z_mean != 0 else 0.0
        cvR  = (sr / abs(z_mean)) * 100 if z_mean != 0 else 0.0

        bias_abs = z_mean - float(x_mean)
        bias_rel = (bias_abs / float(x_mean)) * 100 if x_mean != 0 else 0.0
        recouv   = (z_mean / float(x_mean)) * 100 if x_mean != 0 else 0.0

        # Test de normalité des résidus
        residuals = z_vals - np.repeat(serie_means, J_list)
        shapiro_stat = shapiro_p = shapiro_normal = None
        if len(residuals) >= 3:
            try:
                shap_stat, shap_p = shapiro(residuals)
                shapiro_stat   = round(float(shap_stat), 5)
                shapiro_p      = round(float(shap_p), 5)
                shapiro_normal = bool(shap_p > 0.05)
            except Exception:
                pass

        criteria.append({
            "niveau":    niveau,
            "xMean":     round(float(x_mean), 8),
            "zMean":     round(z_mean, 8),
            "Sr2":       round(sr2, 10),
            "SB2":       round(sb2, 10),
            "SFI2":      round(sfi2, 10),
            "sr":        round(sr, 8),
            "sB":        round(sB, 8),
            "sFI":       round(sFI, 8),
            "cv":        round(cv, 4),
            "cvR":       round(cvR, 4),
            "biasMoy":   round(bias_abs, 8),
            "bRel":      round(bias_rel, 4),
            "recouvMoy": round(recouv, 4),
            "I":  I, "J": J, "N": N,
            "shapiro_stat":   shapiro_stat,
            "shapiro_p":      shapiro_p,
            "shapiro_normal": shapiro_normal,
        })

    criteria.sort(key=lambda c: c["xMean"])
    return criteria


# ─── Intervalles β-expectation (Mee, 1984) ────────────────────────────────────

def compute_tolerance_intervals(
    criteria: List[Dict[str, Any]],
    beta: float = 0.80,
    lambda_val: float = 0.10
) -> List[Dict[str, Any]]:
    """
    Calcule les intervalles β-expectation de tolérance selon la méthode
    de Mee (1984) comme formalisée par Feinberg (2010).

    LTB = Z̄ - k_tol · s_IT    (Limite de Tolérance Basse)
    LTH = Z̄ + k_tol · s_IT    (Limite de Tolérance Haute)

    k_tol est le quantile d'ordre (1+β)/2 de la loi t(ν).
    """
    result = []
    la_basse = (1 - lambda_val) * 100
    la_haute = (1 + lambda_val) * 100

    for c in criteria:
        I, J      = c["I"], c["J"]
        sr2, sb2  = c["Sr2"], c["SB2"]
        sfi2      = c["SFI2"]
        sfi       = float(np.sqrt(sfi2))
        z_mean    = c["zMean"]
        x_mean    = c["xMean"]

        # Ratio R = s²B / s²r
        R = sb2 / sr2 if sr2 > 1e-20 else 0.0

        # Facteur d'élargissement B (Mee)
        denom_B = J * R + 1
        B       = float(np.sqrt((R + 1) / denom_B)) if denom_B > 0 else 1.0

        # Écart-type d'incertitude totale s_IT
        denom_sit = I * J * B ** 2
        s_it = sfi * float(np.sqrt(1 + (1 / denom_sit if denom_sit > 0 else 0)))

        # Degrés de liberté effectifs (Welch-Satterthwaite)
        num_nu = (R + 1) ** 2
        den_nu = ((R + 1 / J) ** 2) / max(1, I - 1) + (1 - 1 / J) / (I * J)
        nu     = int(round(num_nu / den_nu)) if den_nu > 0 else I * (J - 1)
        nu     = max(1, nu)

        # Quantile de tolérance
        p_student = (1 + beta) / 2
        k_tol     = abs(float(stats.t.ppf(p_student, df=nu)))

        # Bornes absolues
        ltb_abs = z_mean - k_tol * s_it
        lth_abs = z_mean + k_tol * s_it

        # Bornes relatives (%)
        ltb_rel = (ltb_abs / x_mean) * 100 if x_mean != 0 else None
        lth_rel = (lth_abs / x_mean) * 100 if x_mean != 0 else None
        recouv  = (z_mean / x_mean) * 100 if x_mean != 0 else 100.0

        # Erreur totale = |biais%| + k·sFI%
        sfi_rel    = (sfi / abs(x_mean)) * 100 if x_mean != 0 else 0.0
        biais_rel  = abs(c["bRel"])
        error_total = biais_rel + k_tol * sfi_rel

        # Critère d'acceptabilité
        accept = (
            ltb_rel is not None and lth_rel is not None
            and ltb_rel >= la_basse
            and lth_rel <= la_haute
        )

        result.append({
            "niveau":     c["niveau"],
            "xMean":      round(x_mean, 8),
            "zMean":      round(z_mean, 8),
            "recouvRel":  round(recouv, 4),
            "sIT":        round(s_it, 8),
            "ktol":       round(k_tol, 6),
            "nu":         nu,
            "R":          round(R, 6),
            "ltbAbs":     round(ltb_abs, 8),
            "lthAbs":     round(lth_abs, 8),
            "ltbRel":     round(ltb_rel, 4) if ltb_rel is not None else None,
            "lthRel":     round(lth_rel, 4) if lth_rel is not None else None,
            "laBasse":    round(la_basse, 2),
            "laHaute":    round(la_haute, 2),
            "accept":     accept,
            "errorTotal": round(error_total, 4),
        })

    return result


# ─── Test de Grubbs ───────────────────────────────────────────────────────────

def grubbs_test(
    data: List[float],
    alpha: float = 0.05,
    iterative: bool = True
) -> Dict[str, Any]:
    """
    Test de Grubbs pour la détection d'une valeur aberrante.
    Si iterative=True, réapplique après exclusion (max 2 passes).

    Classification :
    - 'ok'       : G ≤ G_crit
    - 'suspect'  : G_crit < G (p ~ 0.05)
    - 'aberrant' : G > G_crit à α/2
    """
    arr  = np.array(data, dtype=float)
    n    = len(arr)

    if n < 3:
        return {"error": "Minimum 3 valeurs requises", "suspect": False}

    mean = float(np.mean(arr))
    std  = float(np.std(arr, ddof=1))

    if std < 1e-14:
        return {"G": 0.0, "Gcrit": float("inf"), "suspect": False,
                "mean": mean, "std": std, "classification": "ok"}

    G_vals = np.abs(arr - mean) / std
    G      = float(np.max(G_vals))
    idx    = int(np.argmax(G_vals))

    # G_crit à α selon Grubbs (1969)
    def gcrit(n: int, a: float) -> float:
        t_c = stats.t.ppf(1 - a / (2 * n), df=n - 2)
        return ((n - 1) / np.sqrt(n)) * np.sqrt(t_c ** 2 / (n - 2 + t_c ** 2))

    G_crit    = float(gcrit(n, alpha))
    G_crit_05 = float(gcrit(n, 0.05))

    suspect   = bool(G > G_crit)
    classif   = "ok"
    if G > G_crit:
        classif = "aberrant" if G > G_crit_05 * 1.15 else "suspect"

    result = {
        "G":          round(G, 5),
        "Gcrit":      round(G_crit, 5),
        "suspect":    suspect,
        "suspectIdx": idx,
        "suspectVal": round(float(arr[idx]), 6),
        "mean":       round(mean, 6),
        "std":        round(std, 6),
        "classification": classif,
        "n":          n,
    }

    # Passe itérative (Grubbs bilatéral)
    if iterative and suspect and n > 4:
        arr2    = np.delete(arr, idx)
        result2 = grubbs_test(arr2.tolist(), alpha=alpha, iterative=False)
        result["second_pass"] = result2

    return result


def compute_outliers_by_level(
    found: List[Dict[str, Any]],
    alpha: float = 0.05
) -> List[Dict[str, Any]]:
    """Applique le test de Grubbs par niveau sur les Z retrouvées."""
    by_niveau: Dict[str, list] = defaultdict(list)
    for r in found:
        by_niveau[r["niveau"]].append(r["zRetrouvee"])

    results = []
    for niveau, z_vals in by_niveau.items():
        all_z = [r["zRetrouvee"] for r in found if r["niveau"] == niveau]
        x_mean = np.mean([r["xRef"] for r in found if r["niveau"] == niveau])
        g = grubbs_test(z_vals, alpha=alpha)
        g["niveau"] = niveau
        g["xMean"]  = round(float(x_mean), 8)
        results.append(g)

    return results


# ─── Tests statistiques ───────────────────────────────────────────────────────

def shapiro_wilk_by_level(found: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Test de Shapiro-Wilk de normalité par niveau (résidus)."""
    by_niveau: Dict[str, list] = defaultdict(list)
    for r in found:
        by_niveau[r["niveau"]].append(r)

    results = []
    for niveau, rows in by_niveau.items():
        by_serie: Dict[str, list] = defaultdict(list)
        for r in rows:
            by_serie[r["serie"]].append(r["zRetrouvee"])

        serie_means = {s: np.mean(v) for s, v in by_serie.items()}
        residuals   = [r["zRetrouvee"] - serie_means[r["serie"]] for r in rows]

        if len(residuals) < 3:
            continue
        try:
            stat, p = shapiro(residuals)
            results.append({
                "niveau":  niveau,
                "stat":    round(float(stat), 5),
                "p_value": round(float(p), 5),
                "normal":  bool(p > 0.05),
            })
        except Exception as e:
            logger.warning("Shapiro-Wilk niveau %s: %s", niveau, e)

    return results


def homogeneity_of_variance(
    found: List[Dict[str, Any]],
    alpha: float = 0.05
) -> Dict[str, Any]:
    """
    Test d'homogénéité des variances inter-niveaux (Levene et Bartlett).
    Utilise les séries moyennées par niveau.
    """
    by_niveau: Dict[str, list] = defaultdict(list)
    for r in found:
        by_niveau[r["niveau"]].append(r["zRetrouvee"])

    groups = list(by_niveau.values())
    if len(groups) < 2 or any(len(g) < 2 for g in groups):
        return {"error": "Données insuffisantes pour le test d'homogénéité"}

    try:
        lev_stat, lev_p   = levene(*groups, center="median")
        bart_stat, bart_p = bartlett(*groups)

        return {
            "levene": {
                "test":        "levene",
                "stat":        round(float(lev_stat), 5),
                "p_value":     round(float(lev_p), 5),
                "homogeneous": bool(lev_p > alpha),
            },
            "bartlett": {
                "test":        "bartlett",
                "stat":        round(float(bart_stat), 5),
                "p_value":     round(float(bart_p), 5),
                "homogeneous": bool(bart_p > alpha),
            },
            "conclusion": "Variances homogènes" if lev_p > alpha else "Hétérogénéité des variances détectée",
        }
    except Exception as e:
        return {"error": str(e)}


def anova_one_way(
    data_by_group: Dict[str, List[float]]
) -> Dict[str, Any]:
    """
    ANOVA à un facteur (séries) pour analyse de la variance inter-séries.
    Retourne F, p-value, eta², omega².
    """
    groups  = list(data_by_group.values())
    labels  = list(data_by_group.keys())
    n_total = sum(len(g) for g in groups)
    k       = len(groups)

    if k < 2 or n_total < k + 1:
        return {"error": "Données insuffisantes pour l'ANOVA"}

    try:
        f_stat, p_val = stats.f_oneway(*groups)

        # Eta² (rapport de corrélation)
        grand_mean = np.mean([v for g in groups for v in g])
        ss_between = sum(len(g) * (np.mean(g) - grand_mean) ** 2 for g in groups)
        ss_total   = sum((v - grand_mean) ** 2 for g in groups for v in g)
        eta2       = ss_between / ss_total if ss_total > 0 else 0.0
        # Omega² (estimateur sans biais)
        ss_within  = ss_total - ss_between
        ms_within  = ss_within / max(1, n_total - k)
        omega2     = (ss_between - (k - 1) * ms_within) / (ss_total + ms_within) \
                     if (ss_total + ms_within) > 0 else 0.0

        return {
            "F":       round(float(f_stat), 5),
            "p_value": round(float(p_val), 5),
            "df_between": k - 1,
            "df_within":  n_total - k,
            "eta2":    round(eta2, 5),
            "omega2":  round(omega2, 5),
            "significant": bool(p_val < 0.05),
            "groups":  labels,
        }
    except Exception as e:
        return {"error": str(e)}


# ─── Validité et domaine ──────────────────────────────────────────────────────

def compute_validity(
    tolerances: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Synthèse de la validité et détermination du domaine de validité."""
    n_valid  = sum(1 for t in tolerances if t["accept"])
    n_total  = len(tolerances)
    valid    = n_valid == n_total and n_total > 0
    partial  = 0 < n_valid < n_total
    invalid  = n_valid == 0

    # Domaine de concentrations validées (niveaux consécutifs acceptés)
    valid_levels = [t for t in tolerances if t["accept"]]
    domain: Optional[Dict[str, Any]] = None
    if valid_levels:
        x_min = min(t["xMean"] for t in valid_levels)
        x_max = max(t["xMean"] for t in valid_levels)
        domain = {
            "xMin": x_min,
            "xMax": x_max,
            "nLevels": len(valid_levels),
            "levels": [t["niveau"] for t in valid_levels],
        }

    return {
        "valid":       valid,
        "partial":     partial,
        "invalid":     invalid,
        "nValid":      n_valid,
        "nTotal":      n_total,
        "pct":         round(n_valid / n_total * 100, 1) if n_total > 0 else 0.0,
        "validDomain": domain,
    }


# ─── Score de qualité ─────────────────────────────────────────────────────────

def compute_quality_score(
    criteria: List[Dict[str, Any]],
    tolerances: List[Dict[str, Any]],
    normality: List[Dict[str, Any]],
    homogeneity: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Score de qualité global de la méthode sur 100 points.
    Pondération :
      - Profil (niveaux valides)   40%
      - Justesse (biais moyen)     25%
      - Fidélité (CV moyen)        20%
      - Normalité résidus          10%
      - Homogénéité variances       5%
    """
    details = []

    # 1. Profil
    n_valid = sum(1 for t in tolerances if t["accept"])
    n_total = len(tolerances)
    score_profil = (n_valid / n_total * 100) if n_total > 0 else 0.0

    # 2. Justesse (biais relatif moyen)
    if criteria:
        mean_bias = np.mean([abs(c["bRel"]) for c in criteria])
        # 0% biais → 100 ; ≥10% biais → 0
        score_just = max(0.0, 100 - mean_bias * 10)
        if mean_bias > 5:
            details.append(f"Biais moyen élevé : {mean_bias:.1f}%")
    else:
        score_just = 0.0

    # 3. Fidélité (CV moyen)
    if criteria:
        mean_cv = np.mean([c["cv"] for c in criteria])
        score_fid = max(0.0, 100 - mean_cv * 5)
        if mean_cv > 10:
            details.append(f"CV moyen élevé : {mean_cv:.1f}%")
    else:
        score_fid = 0.0

    # 4. Normalité
    if normality:
        n_normal    = sum(1 for n_ in normality if n_.get("normal", True))
        score_norm  = (n_normal / len(normality)) * 100
    else:
        score_norm  = 80.0  # Neutre si non calculé

    # 5. Homogénéité
    if homogeneity and "levene" in homogeneity:
        score_homo = 100.0 if homogeneity["levene"].get("homogeneous", True) else 40.0
    else:
        score_homo = 80.0

    # Score global pondéré
    overall = (
        0.40 * score_profil +
        0.25 * score_just   +
        0.20 * score_fid    +
        0.10 * score_norm   +
        0.05 * score_homo
    )

    if overall >= 90:   label = "Excellent"
    elif overall >= 75: label = "Bon"
    elif overall >= 55: label = "Acceptable"
    elif overall >= 35: label = "Insuffisant"
    else:               label = "Critique"

    return {
        "overall":     round(overall, 1),
        "justesse":    round(score_just, 1),
        "fidelite":    round(score_fid, 1),
        "profil":      round(score_profil, 1),
        "normalite":   round(score_norm, 1),
        "homogeneite": round(score_homo, 1),
        "label":       label,
        "details":     details,
    }
