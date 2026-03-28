/**
 * ================================================================
 * modules/demoData.js — Données de démonstration (Feinberg, 2010)
 * Dosage du nicotinamide (vitamine B3) dans le lait par HPLC
 * Méthode indirecte, I=3 séries, J=3 répétitions, K=3 niveaux
 * ================================================================
 */
"use strict";

const FEINBERG_VALIDATION = [
  { niveau: "A", serie: "Jour 1", rep: 1, xRef: 0.40, yResponse: 22.6 },
  { niveau: "A", serie: "Jour 1", rep: 2, xRef: 0.40, yResponse: 22.1 },
  { niveau: "A", serie: "Jour 1", rep: 3, xRef: 0.40, yResponse: 22.4 },
  { niveau: "A", serie: "Jour 2", rep: 1, xRef: 0.40, yResponse: 23.3 },
  { niveau: "A", serie: "Jour 2", rep: 2, xRef: 0.40, yResponse: 24.1 },
  { niveau: "A", serie: "Jour 2", rep: 3, xRef: 0.40, yResponse: 23.9 },
  { niveau: "A", serie: "Jour 3", rep: 1, xRef: 0.40, yResponse: 23.8 },
  { niveau: "A", serie: "Jour 3", rep: 2, xRef: 0.40, yResponse: 23.6 },
  { niveau: "A", serie: "Jour 3", rep: 3, xRef: 0.40, yResponse: 23.5 },
  { niveau: "B", serie: "Jour 1", rep: 1, xRef: 2.0, yResponse: 135.0 },
  { niveau: "B", serie: "Jour 1", rep: 2, xRef: 2.0, yResponse: 135.1 },
  { niveau: "B", serie: "Jour 1", rep: 3, xRef: 2.0, yResponse: 129.9 },
  { niveau: "B", serie: "Jour 2", rep: 1, xRef: 2.0, yResponse: 137.6 },
  { niveau: "B", serie: "Jour 2", rep: 2, xRef: 2.0, yResponse: 135.2 },
  { niveau: "B", serie: "Jour 2", rep: 3, xRef: 2.0, yResponse: 138.8 },
  { niveau: "B", serie: "Jour 3", rep: 1, xRef: 2.0, yResponse: 136.3 },
  { niveau: "B", serie: "Jour 3", rep: 2, xRef: 2.0, yResponse: 135.1 },
  { niveau: "B", serie: "Jour 3", rep: 3, xRef: 2.0, yResponse: 134.4 },
  { niveau: "C", serie: "Jour 1", rep: 1, xRef: 4.0, yResponse: 275.2 },
  { niveau: "C", serie: "Jour 1", rep: 2, xRef: 4.0, yResponse: 276.9 },
  { niveau: "C", serie: "Jour 1", rep: 3, xRef: 4.0, yResponse: 261.3 },
  { niveau: "C", serie: "Jour 2", rep: 1, xRef: 4.0, yResponse: 268.1 },
  { niveau: "C", serie: "Jour 2", rep: 2, xRef: 4.0, yResponse: 269.7 },
  { niveau: "C", serie: "Jour 2", rep: 3, xRef: 4.0, yResponse: 276.9 },
  { niveau: "C", serie: "Jour 3", rep: 1, xRef: 4.0, yResponse: 271.6 },
  { niveau: "C", serie: "Jour 3", rep: 2, xRef: 4.0, yResponse: 273.3 },
  { niveau: "C", serie: "Jour 3", rep: 3, xRef: 4.0, yResponse: 275.0 },
];

const FEINBERG_ETALONNAGE = [
  { serie: "Jour 1", niveau: "Bas",  rep: 1, xEtalon: 0.4, yResponse: 22.7 },
  { serie: "Jour 1", niveau: "Bas",  rep: 2, xEtalon: 0.4, yResponse: 23.1 },
  { serie: "Jour 2", niveau: "Bas",  rep: 1, xEtalon: 0.4, yResponse: 22.9 },
  { serie: "Jour 2", niveau: "Bas",  rep: 2, xEtalon: 0.4, yResponse: 23.2 },
  { serie: "Jour 3", niveau: "Bas",  rep: 1, xEtalon: 0.4, yResponse: 21.9 },
  { serie: "Jour 3", niveau: "Bas",  rep: 2, xEtalon: 0.4, yResponse: 22.1 },
  { serie: "Jour 1", niveau: "Haut", rep: 1, xEtalon: 4.0, yResponse: 281.6 },
  { serie: "Jour 1", niveau: "Haut", rep: 2, xEtalon: 4.0, yResponse: 275.3 },
  { serie: "Jour 2", niveau: "Haut", rep: 1, xEtalon: 4.0, yResponse: 275.3 },
  { serie: "Jour 2", niveau: "Haut", rep: 2, xEtalon: 4.0, yResponse: 274.6 },
  { serie: "Jour 3", niveau: "Haut", rep: 1, xEtalon: 4.0, yResponse: 272.0 },
  { serie: "Jour 3", niveau: "Haut", rep: 2, xEtalon: 4.0, yResponse: 273.0 },
];

const FEINBERG_CONFIG = {
  methode:    "Dosage du nicotinamide (vitamine B3) dans le lait par HPLC fluorimétrique",
  materiau:   "Nicotinamide",
  unite:      "mg/L",
  methodType: "indirect",
  modelType:  "linear",
  beta:       0.80,
  lambda:     0.10,
  K: 3, I: 3, J: 3,
  reference:  "Feinberg M. (2010). Interprétation du profil d'exactitude. SFSTP.",
};

const DIRECT_VALIDATION = [
  { niveau: "N1", serie: "Jour 1", rep: 1, xRef: 50.0,  yResponse: 49.82 },
  { niveau: "N1", serie: "Jour 1", rep: 2, xRef: 50.0,  yResponse: 50.14 },
  { niveau: "N1", serie: "Jour 1", rep: 3, xRef: 50.0,  yResponse: 49.95 },
  { niveau: "N1", serie: "Jour 2", rep: 1, xRef: 50.0,  yResponse: 50.31 },
  { niveau: "N1", serie: "Jour 2", rep: 2, xRef: 50.0,  yResponse: 49.76 },
  { niveau: "N1", serie: "Jour 2", rep: 3, xRef: 50.0,  yResponse: 50.08 },
  { niveau: "N1", serie: "Jour 3", rep: 1, xRef: 50.0,  yResponse: 50.22 },
  { niveau: "N1", serie: "Jour 3", rep: 2, xRef: 50.0,  yResponse: 49.91 },
  { niveau: "N1", serie: "Jour 3", rep: 3, xRef: 50.0,  yResponse: 50.05 },
  { niveau: "N2", serie: "Jour 1", rep: 1, xRef: 150.0, yResponse: 149.45 },
  { niveau: "N2", serie: "Jour 1", rep: 2, xRef: 150.0, yResponse: 150.82 },
  { niveau: "N2", serie: "Jour 1", rep: 3, xRef: 150.0, yResponse: 150.11 },
  { niveau: "N2", serie: "Jour 2", rep: 1, xRef: 150.0, yResponse: 151.03 },
  { niveau: "N2", serie: "Jour 2", rep: 2, xRef: 150.0, yResponse: 149.78 },
  { niveau: "N2", serie: "Jour 2", rep: 3, xRef: 150.0, yResponse: 150.44 },
  { niveau: "N2", serie: "Jour 3", rep: 1, xRef: 150.0, yResponse: 150.29 },
  { niveau: "N2", serie: "Jour 3", rep: 2, xRef: 150.0, yResponse: 149.62 },
  { niveau: "N2", serie: "Jour 3", rep: 3, xRef: 150.0, yResponse: 150.87 },
  { niveau: "N3", serie: "Jour 1", rep: 1, xRef: 300.0, yResponse: 299.12 },
  { niveau: "N3", serie: "Jour 1", rep: 2, xRef: 300.0, yResponse: 300.84 },
  { niveau: "N3", serie: "Jour 1", rep: 3, xRef: 300.0, yResponse: 300.23 },
  { niveau: "N3", serie: "Jour 2", rep: 1, xRef: 300.0, yResponse: 301.45 },
  { niveau: "N3", serie: "Jour 2", rep: 2, xRef: 300.0, yResponse: 299.67 },
  { niveau: "N3", serie: "Jour 2", rep: 3, xRef: 300.0, yResponse: 300.91 },
  { niveau: "N3", serie: "Jour 3", rep: 1, xRef: 300.0, yResponse: 300.34 },
  { niveau: "N3", serie: "Jour 3", rep: 2, xRef: 300.0, yResponse: 299.88 },
  { niveau: "N3", serie: "Jour 3", rep: 3, xRef: 300.0, yResponse: 301.12 },
];

const DIRECT_CONFIG = {
  methode:    "Dosage du NaCl par titrimétrie (méthode de Mohr) — méthode directe",
  materiau:   "Chlorure de sodium (NaCl)",
  unite:      "mg/L",
  methodType: "direct",
  beta:       0.80,
  lambda:     0.05,
  K: 3, I: 3, J: 3,
};

window.DemoData = {
  FEINBERG_VALIDATION,
  FEINBERG_ETALONNAGE,
  FEINBERG_CONFIG,
  DIRECT_VALIDATION,
  DIRECT_CONFIG,
};