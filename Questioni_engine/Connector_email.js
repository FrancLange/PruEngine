/**
 * ==========================================================================================
 * CONNECTOR_EMAIL.js - Connettore Email Engine v1.0.0
 * ==========================================================================================
 * Sincronizza LOG_IN da Email Engine → EMAIL_IN locale
 * Pattern identico a Connector_Fornitori
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CONFIGURAZIONE CONNETTORE
// ═══════════════════════════════════════════════════════════════════════

var CONNECTOR_EMAIL = {
  NOME: "Connector_Email",
  VERSIONE: "1.0.0",
  
  // Tab sorgente nel foglio Email Engine
  SOURCE_SHEET: "LOG_IN",
  
  // Tab destinazione locale
  DEST_SHEET: "EMAIL_IN",
  
  // Colonne sorgente (Email Engine CONFIG.COLONNE_LOG_IN)
  COLONNE_SORGENTE: {
    TIMESTAMP: "TIMESTAMP",
    ID_EMAIL: "ID_EMAIL",
    MITTENTE: "MITTENTE",
    OGGETTO: "OGGETTO",
    CORPO: "CORPO",
    L3_TAGS_FINALI: "L3_TAGS_FINALI",
    L3_SINTESI_FINALE: "L3_SINTESI_FINALE",
    L3_SCORES_JSON: "L3_SCORES_JSON",
    L3_CONFIDENCE: "L3_CONFIDENCE",
    L3_AZIONI_SUGGERITE: "L3_AZIONI_SUGGERITE",
    STATUS: "STATUS"
  },
  
  // Cache runtime
  _cache: {
    emails: null,
    timestamp: null
  }
};

// ═══════════════════════════════════════════════════════════════════════
// VERIFICA CONNETTORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se il connettore è attivo e configurato
 */
function isConnectorEmailAttivo() {
  try {
    var enabled = getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_SYNC_ENABLED, 
      "SI"
    );
    
    if (enabled === "NO" || enabled === false) {
      return false;
    }
    
    var engineId = getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, 
      ""
    );
    
    if (!engineId || engineId === "" || engineId.indexOf("Incolla") >= 0) {
      return false;
    }
    
    return true;
    
  } catch(e) {
    return false;
  }
}

/**
 * Verifica se il connettore può essere attivato
 */
function checkConnectorEmailReady() {
  var result = { canActivate: false, reason: "" };
  
  try {
    var engineId = getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, 
      ""
    );
    
    if (!engineId || engineId === "" || engineId.indexOf("Incolla") >= 0) {
      result.reason = "ID foglio Email Engine non configurato in SETUP";
      return result;
    }
    
    // Prova ad aprire
    try {
      var sourceSs = SpreadsheetApp.openById(engineId);
      var sourceSheet = sourceSs.getSheetByName(CONNECTOR_EMAIL.SOURCE_SHEET);
      
      if (!sourceSheet) {
        result.reason = "Tab LOG_IN non trovata nel foglio Email Engine";
        return result;
      }
      
      result.canActivate = true;
      result.reason = "OK - Email Engine raggiungibile";
      
    } catch(e) {
      result.reason = "Impossibile accedere al foglio Email Engine: " + e.message;
    }
    
  } catch(e) {
    result.reason = "Errore verifica: " + e.message;
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// SYNC PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Sincronizza email da Email Engine
 * @param {Boolean} silent - Se true, non mostra alert
 * @returns {Object} {success, righe, filtrate, durata, errore}
 */
function syncEmailDaEngine(silent) {
  var startTime = new Date();
  var risultato = {
    success: false,
    righe: 0,
    filtrate: 0,
    durata: 0,
    errore: null
  };
  
  logQuestioni("Sync Email START");
  
  try {
    // 1. Ottieni ID sorgente
    var engineId = getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, 
      ""
    );
    
    if (!engineId || engineId === "") {
      throw new Error("ID Email Engine non configurato in SETUP");
    }
    
    // 2. Apri foglio sorgente
    var sourceSs;
    try {
      sourceSs = SpreadsheetApp.openById(engineId);
    } catch(e) {
      throw new Error("Impossibile aprire Email Engine. Verifica ID e permessi: " + e.message);
    }
    
    var sourceSheet = sourceSs.getSheetByName(CONNECTOR_EMAIL.SOURCE_SHEET);
    if (!sourceSheet) {
      throw new Error("Tab LOG_IN non trovata in Email Engine");
    }
    
    // 3. Leggi dati sorgente
    var sourceData = sourceSheet.getDataRange().getValues();
    
    if (sourceData.length <= 1) {
      logQuestioni("⚠️ LOG_IN vuoto nel Email Engine");
      risultato.success = true;
      risultato.righe = 0;
      return risultato;
    }
    
    var sourceHeaders = sourceData[0];
    var sourceRows = sourceData.slice(1);
    
    logQuestioni("Lette " + sourceRows.length + " righe da Email Engine");
    
    // 4. Mappa colonne sorgente
    var colMapSource = {};
    Object.keys(CONNECTOR_EMAIL.COLONNE_SORGENTE).forEach(function(key) {
      var nome = CONNECTOR_EMAIL.COLONNE_SORGENTE[key];
      var idx = sourceHeaders.indexOf(nome);
      if (idx >= 0) colMapSource[key] = idx;
    });
    
    // 5. Carica parametri filtro
    var maxRighe = parseInt(getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_MAX_RIGHE,
      QUESTIONI_CONFIG.DEFAULTS.EMAIL_MAX_RIGHE
    ));
    
    var giorniIndietro = parseInt(getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_GIORNI_INDIETRO,
      QUESTIONI_CONFIG.DEFAULTS.EMAIL_GIORNI_INDIETRO
    ));
    
    var filtroScore = parseInt(getQuestioniSetupValue(
      QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_FILTRO_SCORE,
      QUESTIONI_CONFIG.DEFAULTS.EMAIL_FILTRO_SCORE
    ));
    
    var dataLimite = new Date();
    dataLimite.setDate(dataLimite.getDate() - giorniIndietro);
    
    // 6. Filtra email rilevanti
    var emailFiltrate = [];
    
    for (var i = 0; i < sourceRows.length; i++) {
      var row = sourceRows[i];
      
      // Skip righe vuote
      if (!row[colMapSource.ID_EMAIL]) continue;
      
      // Skip email non analizzate
      var status = row[colMapSource.STATUS] || "";
      if (status !== "ANALIZZATO" && status !== "NEEDS_REVIEW") continue;
      
      // Filtro data
      var timestamp = new Date(row[colMapSource.TIMESTAMP]);
      if (timestamp < dataLimite) continue;
      
      // Filtro score (problema o urgente)
      var scoresJson = row[colMapSource.L3_SCORES_JSON] || "{}";
      var scores = {};
      try {
        scores = JSON.parse(scoresJson);
      } catch(e) {}
      
      var scoreProblema = parseInt(scores.problema) || 0;
      var scoreUrgente = parseInt(scores.urgente) || 0;
      
      // Includi se problema >= filtroScore OR urgente >= 80 OR ha azione CREA_QUESTIONE
      var azioniSuggerite = row[colMapSource.L3_AZIONI_SUGGERITE] || "";
      var haAzioneQuestione = azioniSuggerite.indexOf("CREA_QUESTIONE") >= 0;
      
      if (scoreProblema >= filtroScore || scoreUrgente >= 80 || haAzioneQuestione) {
        emailFiltrate.push({
          timestamp: row[colMapSource.TIMESTAMP],
          id: row[colMapSource.ID_EMAIL],
          mittente: row[colMapSource.MITTENTE],
          oggetto: row[colMapSource.OGGETTO],
          corpo: row[colMapSource.CORPO],
          tags: row[colMapSource.L3_TAGS_FINALI],
          sintesi: row[colMapSource.L3_SINTESI_FINALE],
          scoresJson: scoresJson,
          confidence: row[colMapSource.L3_CONFIDENCE],
          azioni: azioniSuggerite,
          status: status,
          scoreProblema: scoreProblema,
          scoreUrgente: scoreUrgente
        });
      }
    }
    
    risultato.filtrate = emailFiltrate.length;
    logQuestioni("Filtrate " + emailFiltrate.length + " email rilevanti");
    
    // 7. Ordina per data (più recenti prima) e limita
    emailFiltrate.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    if (emailFiltrate.length > maxRighe) {
      emailFiltrate = emailFiltrate.slice(0, maxRighe);
      logQuestioni("Limitate a " + maxRighe + " righe");
    }
    
    // 8. Prepara tab destinazione
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.EMAIL_IN);
    
    if (!destSheet) {
      destSheet = ss.insertSheet(QUESTIONI_CONFIG.SHEETS.EMAIL_IN);
      destSheet.setTabColor("#4CAF50");
      logQuestioni("Creata tab EMAIL_IN");
    }
    
    // 9. Leggi ID esistenti per preservare dati locali (ID_QUESTIONE, PROCESSATO)
    var existingData = {};
    if (destSheet.getLastRow() > 1) {
      var oldData = destSheet.getDataRange().getValues();
      var oldHeaders = oldData[0];
      var colId = oldHeaders.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_EMAIL);
      var colQuestione = oldHeaders.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_QUESTIONE);
      var colProcessato = oldHeaders.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.PROCESSATO);
      var colStatusQ = oldHeaders.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_QUESTIONE);
      
      for (var j = 1; j < oldData.length; j++) {
        var emailId = oldData[j][colId];
        if (emailId) {
          existingData[emailId] = {
            idQuestione: colQuestione >= 0 ? oldData[j][colQuestione] : "",
            processato: colProcessato >= 0 ? oldData[j][colProcessato] : "",
            statusQuestione: colStatusQ >= 0 ? oldData[j][colStatusQ] : ""
          };
        }
      }
    }
    
    // 10. Costruisci nuovi dati - ORDINE ESPLICITO (importante!)
    var headers = [
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.TIMESTAMP_SYNC,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_EMAIL,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.TIMESTAMP_ORIGINALE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.MITTENTE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.OGGETTO,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.CORPO,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_TAGS_FINALI,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_SINTESI_FINALE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_SCORES_JSON,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_CONFIDENCE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_AZIONI_SUGGERITE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_EMAIL,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_QUESTIONE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_QUESTIONE,
      QUESTIONI_CONFIG.COLONNE_EMAIL_IN.PROCESSATO
    ];
    var newData = [headers];
    
    var now = new Date();
    
    emailFiltrate.forEach(function(email) {
      var existing = existingData[email.id] || {};
      
      var row = [
        now,                          // TIMESTAMP_SYNC
        email.id,                     // ID_EMAIL
        email.timestamp,              // TIMESTAMP_ORIGINALE
        email.mittente,               // MITTENTE
        email.oggetto,                // OGGETTO
        email.corpo,                  // CORPO
        email.tags,                   // L3_TAGS_FINALI
        email.sintesi,                // L3_SINTESI_FINALE
        email.scoresJson,             // L3_SCORES_JSON
        email.confidence,             // L3_CONFIDENCE
        email.azioni,                 // L3_AZIONI_SUGGERITE
        email.status,                 // STATUS_EMAIL
        existing.idQuestione || "",   // ID_QUESTIONE (preservato)
        existing.statusQuestione || "",// STATUS_QUESTIONE (preservato)
        existing.processato || ""     // PROCESSATO (preservato)
      ];
      
      newData.push(row);
    });
    
    // 11. Scrivi
    destSheet.clear();
    
    if (newData.length > 0 && newData[0].length > 0) {
      destSheet.getRange(1, 1, newData.length, newData[0].length)
        .setValues(newData);
    }
    
    // 12. Formattazione
    destSheet.setFrozenRows(1);
    destSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#E8F5E9");
    
    // 13. Aggiorna tracking
    setQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_LAST_SYNC, new Date().toISOString());
    
    // 14. Invalida cache
    CONNECTOR_EMAIL._cache.emails = null;
    CONNECTOR_EMAIL._cache.timestamp = null;
    
    // Risultato
    risultato.success = true;
    risultato.righe = newData.length - 1;
    risultato.durata = Math.round((new Date() - startTime) / 1000);
    
    logQuestioni("✅ Sync OK: " + risultato.righe + " email in " + risultato.durata + "s");
    
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "✅ Sync Email Completato",
        "Sincronizzate " + risultato.righe + " email\n" +
        "Filtrate da " + sourceRows.length + " totali\n" +
        "Durata: " + risultato.durata + " secondi\n\n" +
        "Fonte: " + sourceSs.getName(),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch(e) {
    risultato.errore = e.toString();
    logQuestioni("❌ Sync ERRORE: " + risultato.errore);
    
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "❌ Errore Sync Email",
        risultato.errore,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  return risultato;
}

/**
 * Sync silenzioso
 */
function syncEmailSilent() {
  return syncEmailDaEngine(true);
}

/**
 * Sync da menu
 */
function syncEmailMenu() {
  return syncEmailDaEngine(false);
}

// ═══════════════════════════════════════════════════════════════════════
// LETTURA EMAIL LOCALI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica email da tab locale EMAIL_IN (con cache)
 */
function caricaEmailLocali() {
  var cache = CONNECTOR_EMAIL._cache;
  var now = new Date().getTime();
  
  if (cache.emails && cache.timestamp) {
    var age = (now - cache.timestamp) / 1000;
    if (age < QUESTIONI_CONFIG.DEFAULTS.CACHE_TTL_SEC) {
      return cache.emails;
    }
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.EMAIL_IN);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { headers: [], rows: [], colMap: {} };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    var colMap = {};
    Object.keys(QUESTIONI_CONFIG.COLONNE_EMAIL_IN).forEach(function(key) {
      var nome = QUESTIONI_CONFIG.COLONNE_EMAIL_IN[key];
      var idx = headers.indexOf(nome);
      if (idx >= 0) colMap[key] = idx;
    });
    
    var result = {
      headers: headers,
      rows: rows,
      colMap: colMap
    };
    
    CONNECTOR_EMAIL._cache.emails = result;
    CONNECTOR_EMAIL._cache.timestamp = now;
    
    return result;
    
  } catch(e) {
    logQuestioni("Errore caricamento email locali: " + e.message);
    return { headers: [], rows: [], colMap: {} };
  }
}

/**
 * Carica email da processare (non ancora collegate a questioni)
 */
function caricaEmailDaProcessare() {
  var dati = caricaEmailLocali();
  
  if (!dati.rows || dati.rows.length === 0) {
    return [];
  }
  
  var colMap = dati.colMap;
  var emails = [];
  
  for (var i = 0; i < dati.rows.length; i++) {
    var row = dati.rows[i];
    
    // Skip se già processata
    var processato = row[colMap.PROCESSATO] || "";
    if (processato === "SI") continue;
    
    // Skip se già collegata a questione
    var idQuestione = row[colMap.ID_QUESTIONE] || "";
    if (idQuestione && idQuestione !== "") continue;
    
    var scores = {};
    try {
      scores = JSON.parse(row[colMap.L3_SCORES_JSON] || "{}");
    } catch(e) {}
    
    emails.push({
      rowNum: i + 2,
      id: row[colMap.ID_EMAIL],
      timestamp: row[colMap.TIMESTAMP_ORIGINALE],
      mittente: row[colMap.MITTENTE],
      oggetto: row[colMap.OGGETTO],
      corpo: row[colMap.CORPO],
      tags: row[colMap.L3_TAGS_FINALI],
      sintesi: row[colMap.L3_SINTESI_FINALE],
      scores: scores,
      confidence: row[colMap.L3_CONFIDENCE],
      azioni: row[colMap.L3_AZIONI_SUGGERITE],
      statusEmail: row[colMap.STATUS_EMAIL]
    });
  }
  
  return emails;
}

/**
 * Aggiorna email locale dopo collegamento a questione
 */
function collegaEmailAQuestioneLocale(idEmail, idQuestione) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.EMAIL_IN);
  
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colId = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_EMAIL);
  var colQuestione = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_QUESTIONE);
  var colStatusQ = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_QUESTIONE);
  var colProcessato = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.PROCESSATO);
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][colId] === idEmail) {
      var rowNum = i + 1;
      
      if (colQuestione >= 0) {
        sheet.getRange(rowNum, colQuestione + 1).setValue(idQuestione);
      }
      if (colStatusQ >= 0) {
        sheet.getRange(rowNum, colStatusQ + 1).setValue("COLLEGATA");
      }
      if (colProcessato >= 0) {
        sheet.getRange(rowNum, colProcessato + 1).setValue("SI");
      }
      
      // Invalida cache
      CONNECTOR_EMAIL._cache.emails = null;
      
      return true;
    }
  }
  
  return false;
}

// ═══════════════════════════════════════════════════════════════════════
// STATISTICHE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Statistiche sync email
 */
function getStatsConnectorEmail() {
  var stats = {
    attivo: isConnectorEmailAttivo(),
    tabEsiste: false,
    righe: 0,
    daProcessare: 0,
    collegate: 0,
    ultimoSync: null,
    engineConfigurato: false
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.EMAIL_IN);
    
    if (sheet && sheet.getLastRow() > 1) {
      stats.tabEsiste = true;
      stats.righe = sheet.getLastRow() - 1;
      
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var colProcessato = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.PROCESSATO);
      var colQuestione = headers.indexOf(QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_QUESTIONE);
      
      for (var i = 1; i < data.length; i++) {
        var processato = colProcessato >= 0 ? data[i][colProcessato] : "";
        var idQ = colQuestione >= 0 ? data[i][colQuestione] : "";
        
        if (processato !== "SI" && (!idQ || idQ === "")) {
          stats.daProcessare++;
        }
        if (idQ && idQ !== "") {
          stats.collegate++;
        }
      }
    }
    
    var lastSync = getQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_LAST_SYNC, null);
    if (lastSync) {
      stats.ultimoSync = new Date(lastSync);
    }
    
    var engineId = getQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, "");
    stats.engineConfigurato = !!(engineId && engineId.length > 10);
    
  } catch(e) {
    // Ignora
  }
  
  return stats;
}

// ═══════════════════════════════════════════════════════════════════════
// TRIGGER AUTO-SYNC
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se sync necessario
 */
function isSyncEmailNecessario(maxMinuti) {
  maxMinuti = maxMinuti || QUESTIONI_CONFIG.DEFAULTS.SYNC_INTERVAL_MIN;
  
  var lastSync = getQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_LAST_SYNC, null);
  
  if (!lastSync) return true;
  
  try {
    var lastDate = new Date(lastSync);
    var now = new Date();
    var diffMinuti = (now - lastDate) / (1000 * 60);
    
    return diffMinuti >= maxMinuti;
    
  } catch(e) {
    return true;
  }
}

/**
 * Auto-sync se necessario
 */
function autoSyncEmailSeNecessario() {
  if (!isConnectorEmailAttivo()) {
    return false;
  }
  
  if (isSyncEmailNecessario()) {
    logQuestioni("Auto-sync email necessario");
    var result = syncEmailDaEngine(true);
    return result.success;
  }
  
  return true;
}

/**
 * Configura trigger sync automatico
 */
function configuraTriggerSyncEmail() {
  var ui = SpreadsheetApp.getUi();
  
  var check = checkConnectorEmailReady();
  if (!check.canActivate) {
    ui.alert("❌ Connettore Non Pronto", check.reason, ui.ButtonSet.OK);
    return;
  }
  
  rimuoviTriggerSyncEmail();
  
  var intervallo = QUESTIONI_CONFIG.DEFAULTS.SYNC_INTERVAL_MIN;
  
  ScriptApp.newTrigger('syncEmailSilent')
    .timeBased()
    .everyMinutes(intervallo)
    .create();
  
  logQuestioni("Trigger sync email configurato: ogni " + intervallo + " minuti");
  
  ui.alert(
    "✅ Trigger Configurato",
    "Sync automatico ogni " + intervallo + " minuti.",
    ui.ButtonSet.OK
  );
}

/**
 * Rimuove trigger sync email
 */
function rimuoviTriggerSyncEmail() {
  var triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === 'syncEmailSilent' || handler === 'syncEmailDaEngine') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Mostra stato connettore
 */
function mostraStatoConnectorEmail() {
  var ui = SpreadsheetApp.getUi();
  var stats = getStatsConnectorEmail();
  
  var report = "📧 CONNECTOR EMAIL ENGINE\n";
  report += "══════════════════════════════\n\n";
  
  report += "Stato: " + (stats.attivo ? "✅ ATTIVO" : "❌ NON ATTIVO") + "\n\n";
  
  report += "━━━ CONFIGURAZIONE ━━━\n";
  report += "Email Engine configurato: " + (stats.engineConfigurato ? "✅" : "❌") + "\n";
  report += "Tab EMAIL_IN presente: " + (stats.tabEsiste ? "✅" : "❌") + "\n\n";
  
  report += "━━━ STATISTICHE ━━━\n";
  report += "Email sincronizzate: " + stats.righe + "\n";
  report += "Da processare: " + stats.daProcessare + "\n";
  report += "Collegate a questioni: " + stats.collegate + "\n\n";
  
  report += "━━━ SYNC ━━━\n";
  report += "Ultimo sync: " + (stats.ultimoSync ? stats.ultimoSync.toLocaleString('it-IT') : "Mai") + "\n";
  
  ui.alert("Stato Connector Email", report, ui.ButtonSet.OK);
}

/**
 * Test lookup email
 */
function testConnectorEmail() {
  var ui = SpreadsheetApp.getUi();
  
  var check = checkConnectorEmailReady();
  
  if (!check.canActivate) {
    ui.alert("❌ Test Fallito", check.reason, ui.ButtonSet.OK);
    return;
  }
  
  try {
    var engineId = getQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, "");
    var sourceSs = SpreadsheetApp.openById(engineId);
    var sourceSheet = sourceSs.getSheetByName(CONNECTOR_EMAIL.SOURCE_SHEET);
    
    var righe = sourceSheet.getLastRow() - 1;
    var colonne = sourceSheet.getLastColumn();
    
    ui.alert(
      "✅ Test Connessione OK",
      "Email Engine: " + sourceSs.getName() + "\n" +
      "Tab LOG_IN: " + righe + " email\n" +
      "Colonne: " + colonne + "\n\n" +
      "Permessi lettura verificati.",
      ui.ButtonSet.OK
    );
    
    logQuestioni("Test connessione Email Engine OK");
    
  } catch(e) {
    ui.alert("❌ Test Fallito", e.toString(), ui.ButtonSet.OK);
  }
}