/**
 * ==========================================================================================
 * HELPERS.gs - Utility Condivise v1.0.0
 * ==========================================================================================
 * Funzioni di utilità usate da tutti gli altri file
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LOGGING
// ═══════════════════════════════════════════════════════════════════════

/**
 * Log messaggio nel foglio LOG_SISTEMA
 * @param {String} messaggio - Messaggio da loggare
 */
function logSistema(messaggio) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_SISTEMA);
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// AGGIORNAMENTO STATUS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna STATUS di una email
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} rowNum - Numero riga (1-based)
 * @param {String} status - Nuovo status
 */
function aggiornaStatus(sheet, rowNum, status) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  if (colStatus >= 0) {
    sheet.getRange(rowNum, colStatus + 1).setValue(status);
  } else {
    Logger.log("⚠️ Colonna STATUS non trovata");
  }
}

/**
 * Aggiorna STATUS email (alias per compatibilità)
 */
function aggiornaStatusEmail(sheet, rowNum, status) {
  aggiornaStatus(sheet, rowNum, status);
}

// ═══════════════════════════════════════════════════════════════════════
// CARICAMENTO EMAIL
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica email da analizzare per uno specifico layer
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {String} layer - "L0", "L1", "L2", "L3"
 * @param {Number} maxEmail - Limite email da caricare
 * @returns {Array} Array di email object
 */
function caricaEmailDaAnalizzare(sheet, layer, maxEmail) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Mappa colonne
  var colMap = {};
  Object.keys(CONFIG.COLONNE_LOG_IN).forEach(function(key) {
    var nome = CONFIG.COLONNE_LOG_IN[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  var emails = [];
  
  for (var i = 0; i < rows.length && emails.length < maxEmail; i++) {
    var row = rows[i];
    var rowNum = i + 2;
    
    // Skip righe vuote
    if (!row[colMap.ID_EMAIL]) continue;
    
    var email = {
      rowNum: rowNum,
      timestamp: row[colMap.TIMESTAMP],
      id: row[colMap.ID_EMAIL],
      mittente: row[colMap.MITTENTE],
      oggetto: row[colMap.OGGETTO],
      corpo: row[colMap.CORPO],
      status: row[colMap.STATUS],
      
      // Layer 0
      l0Timestamp: row[colMap.L0_TIMESTAMP],
      l0IsSpam: row[colMap.L0_IS_SPAM],
      l0Confidence: row[colMap.L0_CONFIDENCE],
      
      // Layer 1
      l1Timestamp: row[colMap.L1_TIMESTAMP],
      l1Tags: row[colMap.L1_TAGS],
      l1Sintesi: row[colMap.L1_SINTESI],
      l1ScoresJson: row[colMap.L1_SCORES_JSON],
      l1Modello: row[colMap.L1_MODELLO],
      
      // Layer 2
      l2Timestamp: row[colMap.L2_TIMESTAMP],
      l2Tags: row[colMap.L2_TAGS],
      l2Sintesi: row[colMap.L2_SINTESI],
      l2ScoresJson: row[colMap.L2_SCORES_JSON],
      l2Confidence: row[colMap.L2_CONFIDENCE],
      l2RichiestaRetry: row[colMap.L2_RICHIESTA_RETRY],
      l2Modello: row[colMap.L2_MODELLO],
      
      // Layer 3
      l3Timestamp: row[colMap.L3_TIMESTAMP],
      l3TagsFinali: row[colMap.L3_TAGS_FINALI],
      l3SintesiFinale: row[colMap.L3_SINTESI_FINALE],
      l3ScoresJson: row[colMap.L3_SCORES_JSON],
      l3Confidence: row[colMap.L3_CONFIDENCE],
      l3Divergenza: row[colMap.L3_DIVERGENZA],
      l3NeedsReview: row[colMap.L3_NEEDS_REVIEW],
      l3AzioniSuggerite: row[colMap.L3_AZIONI_SUGGERITE]
    };
    
    // Filtra in base al layer richiesto
    var shouldInclude = false;
    
    switch (layer) {
      case "L0":
        // Email DA_FILTRARE senza L0 completato
        shouldInclude = !email.l0Timestamp && 
                       (email.status === CONFIG.STATUS_EMAIL.DA_FILTRARE);
        break;
        
      case "L1":
        // Email DA_ANALIZZARE senza L1 completato
        shouldInclude = !email.l1Timestamp && 
                       (email.status === CONFIG.STATUS_EMAIL.DA_ANALIZZARE || 
                        email.status === "" || 
                        !email.status);
        break;
      
      case "L2":
        // Email con L1 completato ma senza L2
        shouldInclude = email.l1Timestamp && !email.l2Timestamp;
        break;
      
      case "L3":
        // Email con L1 e L2 completati ma senza L3
        shouldInclude = email.l1Timestamp && email.l2Timestamp && !email.l3Timestamp;
        break;
    }
    
    if (shouldInclude) {
      emails.push(email);
    }
  }
  
  return emails;
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY COLONNE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna colonne specifiche in un foglio
 * @param {Sheet} sheet - Foglio
 * @param {Number} rowNum - Numero riga
 * @param {Array} headers - Array headers
 * @param {Object} updates - {colName: valore}
 */
function aggiornaColonneEmail(sheet, rowNum, headers, updates) {
  Object.keys(updates).forEach(function(key) {
    var value = updates[key];
    var colName = CONFIG.COLONNE_LOG_IN[key] || key;
    var colIndex = headers.indexOf(colName);
    
    if (colIndex >= 0) {
      sheet.getRange(rowNum, colIndex + 1).setValue(value);
    } else {
      Logger.log("⚠️ Colonna non trovata: " + colName + " (key: " + key + ")");
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// PROMPT HELPERS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica un prompt dal foglio PROMPTS
 * @param {String} key - Chiave prompt
 * @returns {Object|null}
 */
function caricaPrompt(key) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PROMPTS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return null;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colKey = headers.indexOf(CONFIG.COLONNE_PROMPTS.KEY);
    var colLivello = headers.indexOf(CONFIG.COLONNE_PROMPTS.LIVELLO_AI);
    var colTesto = headers.indexOf(CONFIG.COLONNE_PROMPTS.TESTO);
    var colCategoria = headers.indexOf(CONFIG.COLONNE_PROMPTS.CATEGORIA);
    var colTemp = headers.indexOf(CONFIG.COLONNE_PROMPTS.TEMPERATURA);
    var colTokens = headers.indexOf(CONFIG.COLONNE_PROMPTS.MAX_TOKENS);
    var colAttivo = headers.indexOf(CONFIG.COLONNE_PROMPTS.ATTIVO);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[colKey] === key && row[colAttivo] === "SI") {
        return {
          key: row[colKey],
          livello: row[colLivello],
          testo: row[colTesto],
          categoria: row[colCategoria],
          temperatura: row[colTemp] || 0.7,
          maxTokens: row[colTokens] || 500
        };
      }
    }
    
    return null;
    
  } catch (error) {
    logSistema("❌ Errore caricamento prompt " + key + ": " + error.toString());
    return null;
  }
}

/**
 * Sostituisci placeholder in un testo
 * @param {String} testo - Testo con {{placeholder}}
 * @param {Object} variabili - {placeholder: valore}
 * @returns {String}
 */
function sostituisciPlaceholder(testo, variabili) {
  var risultato = testo;
  
  Object.keys(variabili).forEach(function(key) {
    var value = variabili[key];
    var regex = new RegExp("{{" + key + "}}", "g");
    risultato = risultato.replace(regex, value || "");
  });
  
  return risultato;
}