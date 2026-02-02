/**
 * ==========================================================================================
 * WRITERS.gs - Scrittura Risultati v1.0.0
 * ==========================================================================================
 * Tutte le funzioni per scrivere risultati AI nel foglio LOG_IN
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LAYER 0 - SPAM FILTER
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scrivi risultati Layer 0 (Spam Filter) nel foglio
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {Object} l0Result - Risultato analisi L0
 */
function scriviRisultatoL0(sheet, rowNum, l0Result) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colL0Timestamp = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_TIMESTAMP) + 1;
  var colL0IsSpam = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_IS_SPAM) + 1;
  var colL0Confidence = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_CONFIDENCE) + 1;
  var colL0Reason = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_REASON) + 1;
  var colL0Modello = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_MODELLO) + 1;
  
  if (colL0Timestamp > 0) {
    sheet.getRange(rowNum, colL0Timestamp).setValue(l0Result.timestamp || new Date());
  }
  if (colL0IsSpam > 0) {
    sheet.getRange(rowNum, colL0IsSpam).setValue(l0Result.isSpam ? "SI" : "NO");
  }
  if (colL0Confidence > 0) {
    sheet.getRange(rowNum, colL0Confidence).setValue(l0Result.confidence || 0);
  }
  if (colL0Reason > 0) {
    sheet.getRange(rowNum, colL0Reason).setValue(l0Result.reason || "");
  }
  if (colL0Modello > 0) {
    sheet.getRange(rowNum, colL0Modello).setValue(l0Result.modello || CONFIG.AI_CONFIG.MODEL_SPAM_FILTER);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 1 - GPT CATEGORIZATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scrivi risultati Layer 1 (GPT) nel foglio
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {Object} risultato - Risultato analisi L1
 */
function scriviRisultatiL1(sheet, rowNum, risultato) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var updates = {
    L1_TIMESTAMP: new Date(),
    L1_TAGS: Array.isArray(risultato.tags) ? risultato.tags.join(", ") : "",
    L1_SINTESI: risultato.sintesi || "",
    L1_SCORES_JSON: JSON.stringify(risultato.scores || {}),
    L1_MODELLO: risultato.modello || "gpt-4o"
  };
  
  aggiornaColonneEmail(sheet, rowNum, headers, updates);
}

/**
 * Scrivi errore Layer 1
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {String} errore - Messaggio errore
 */
function scriviErroreL1(sheet, rowNum, errore) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var updates = {
    L1_TIMESTAMP: new Date(),
    L1_SINTESI: "ERRORE: " + (errore || "").substring(0, 200)
  };
  
  aggiornaColonneEmail(sheet, rowNum, headers, updates);
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 2 - CLAUDE VERIFICATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scrivi risultati Layer 2 (Claude) nel foglio
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {Object} risultato - Risultato analisi L2
 */
function scriviRisultatiL2(sheet, rowNum, risultato) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var updates = {
    L2_TIMESTAMP: new Date(),
    L2_TAGS: Array.isArray(risultato.tags) ? risultato.tags.join(", ") : "",
    L2_SINTESI: risultato.sintesi || "",
    L2_SCORES_JSON: JSON.stringify(risultato.scores || {}),
    L2_CONFIDENCE: risultato.confidence || 0,
    L2_RICHIESTA_RETRY: risultato.richiestaRetry ? "SI" : "NO",
    L2_MODELLO: risultato.modello || "claude-3-5-sonnet-20241022"
  };
  
  aggiornaColonneEmail(sheet, rowNum, headers, updates);
}

/**
 * Scrivi errore Layer 2
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {String} errore - Messaggio errore
 */
function scriviErroreL2(sheet, rowNum, errore) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var updates = {
    L2_TIMESTAMP: new Date(),
    L2_SINTESI: "ERRORE: " + (errore || "").substring(0, 200)
  };
  
  aggiornaColonneEmail(sheet, rowNum, headers, updates);
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 3 - MERGE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scrivi risultati Layer 3 (Merge) nel foglio
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga
 * @param {Object} risultato - Risultato merge L3
 */
function scriviRisultatiL3(sheet, rowNum, risultato) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var azioniStr = "";
  if (risultato.azioniSuggerite && Array.isArray(risultato.azioniSuggerite)) {
    azioniStr = risultato.azioniSuggerite.join(", ");
  }
  
  var updates = {
    L3_TIMESTAMP: new Date(),
    L3_TAGS_FINALI: Array.isArray(risultato.tags) ? risultato.tags.join(", ") : "",
    L3_SINTESI_FINALE: risultato.sintesi || "",
    L3_SCORES_JSON: JSON.stringify(risultato.scores || {}),
    L3_CONFIDENCE: risultato.confidence || 0,
    L3_DIVERGENZA: risultato.divergenza || 0,
    L3_NEEDS_REVIEW: risultato.needsReview ? "SI" : "NO",
    L3_AZIONI_SUGGERITE: azioniStr
  };
  
  aggiornaColonneEmail(sheet, rowNum, headers, updates);
}