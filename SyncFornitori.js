/**
 * ==========================================================================================
 * SYNC FORNITORI - Email Engine v1.1.0
 * ==========================================================================================
 * Gestisce la sincronizzazione tra Email Engine e Fornitori Engine.
 * Mantiene una copia locale (FORNITORI_SYNC) per lookup rapidi.
 * 
 * v1.1.0 - Fix mapping colonne reali Fornitori Engine
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CONFIGURAZIONE SYNC
// ═══════════════════════════════════════════════════════════════════════

var SYNC_CONFIG = {
  // Nome tab locale per cache fornitori
  TAB_SYNC: "FORNITORI_SYNC",
  
  // Mapping: colonna MASTER → colonna LOCALE
  // Le colonne locali sono quelle che creiamo in FORNITORI_SYNC
  MAPPING_COLONNE: {
    "ID_FORNITORE": "ID_FORNITORE",
    "NOME_AZIENDA": "RAGIONE_SOCIALE",
    "EMAIL_ORDINI": "EMAIL_PRINCIPALE",
    "EMAIL_ALTRI": "EMAIL_SECONDARIE",
    "STATUS_FORNITORE": "STATUS",
    "SCONTO_PERCENTUALE": "SCONTO_BASE",
    "NOTE": "NOTE_INTERNE",
    "DATA_ULTIMO_ORDINE": "ULTIMO_ORDINE",
    "PRIORITA_URGENTE": "PRIORITA",
    "PERFORMANCE_SCORE": "PERFORMANCE"
  },
  
  // Colonne nel tab SYNC locale (quello che creiamo)
  COLONNE_SYNC: [
    "ID_FORNITORE",
    "RAGIONE_SOCIALE",
    "EMAIL_PRINCIPALE",
    "EMAIL_SECONDARIE",
    "DOMINIO_EMAIL",      // Calcolato automaticamente
    "STATUS",
    "SCONTO_BASE",
    "NOTE_INTERNE",
    "ULTIMO_ORDINE",
    "PRIORITA",
    "PERFORMANCE",
    "ULTIMA_SYNC"
  ],
  
  // Foglio master nel Fornitori Engine
  FOGLIO_MASTER: "FORNITORI",
  
  // Chiave SETUP per ID Fornitori Engine
  KEY_ID_FORNITORI: "ID_FORNITORI_SOURCE"
};

// ═══════════════════════════════════════════════════════════════════════
// SYNC COMPLETO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Esegue sync completo dal Fornitori Engine
 * Copia tutti i fornitori attivi nel tab locale FORNITORI_SYNC
 * @returns {Object} {success, count, errors}
 */
function syncFornitoriCompleto() {
  var startTime = new Date();
  logSistema("🔄 SYNC FORNITORI: Inizio sync completo");
  
  try {
    // 1. Ottieni ID Fornitori Engine da SETUP
    var idFornitori = getSetupValue(SYNC_CONFIG.KEY_ID_FORNITORI, "");
    
    if (!idFornitori || idFornitori === "" || idFornitori === "Incolla qui ID Fornitori") {
      logSistema("⚠️ SYNC: ID Fornitori Engine non configurato in SETUP");
      return { 
        success: false, 
        count: 0, 
        error: "ID_FORNITORI_SOURCE non configurato. Vai in SETUP e inserisci l'ID del Fornitori Engine." 
      };
    }
    
    // 2. Apri Fornitori Engine
    var ssFornitori;
    try {
      ssFornitori = SpreadsheetApp.openById(idFornitori);
    } catch (e) {
      logSistema("❌ SYNC: Impossibile aprire Fornitori Engine: " + e.toString());
      return { 
        success: false, 
        count: 0, 
        error: "Impossibile aprire Fornitori Engine. Verifica ID e permessi." 
      };
    }
    
    // 3. Leggi dati dal foglio master
    var sheetMaster = ssFornitori.getSheetByName(SYNC_CONFIG.FOGLIO_MASTER);
    if (!sheetMaster) {
      logSistema("❌ SYNC: Foglio FORNITORI non trovato nel Fornitori Engine");
      return { 
        success: false, 
        count: 0, 
        error: "Foglio FORNITORI non trovato nel Fornitori Engine" 
      };
    }
    
    var dataMaster = sheetMaster.getDataRange().getValues();
    if (dataMaster.length <= 1) {
      logSistema("⚠️ SYNC: Fornitori Engine vuoto (solo header)");
      return { success: true, count: 0, error: null };
    }
    
    var headersMaster = dataMaster[0];
    var rowsMaster = dataMaster.slice(1);
    
    // 4. Mappa indici colonne master
    var colMapMaster = {};
    headersMaster.forEach(function(header, idx) {
      colMapMaster[header] = idx;
    });
    
    // Debug: log colonne trovate
    logSistema("📋 Colonne master trovate: " + headersMaster.join(", "));
    
    // 5. Estrai e trasforma dati
    var fornitoriDaSincronizzare = [];
    
    rowsMaster.forEach(function(row, rowIndex) {
      // Leggi ID_FORNITORE
      var idFornitore = colMapMaster.ID_FORNITORE !== undefined ? row[colMapMaster.ID_FORNITORE] : null;
      
      // Skip righe vuote
      if (!idFornitore || idFornitore === "") return;
      
      // Leggi status (se esiste)
      var status = colMapMaster.STATUS_FORNITORE !== undefined ? row[colMapMaster.STATUS_FORNITORE] : "ATTIVO";
      if (status === "DISATTIVO" || status === "ELIMINATO") return;
      
      // Leggi email principale per calcolare dominio
      var emailPrincipale = colMapMaster.EMAIL_ORDINI !== undefined ? row[colMapMaster.EMAIL_ORDINI] : "";
      var dominio = estraiDominio(emailPrincipale);
      
      // Costruisci riga sync con mapping
      var rowSync = [
        idFornitore,
        colMapMaster.NOME_AZIENDA !== undefined ? row[colMapMaster.NOME_AZIENDA] : "",
        emailPrincipale,
        colMapMaster.EMAIL_ALTRI !== undefined ? row[colMapMaster.EMAIL_ALTRI] : "",
        dominio,  // Calcolato
        status,
        colMapMaster.SCONTO_PERCENTUALE !== undefined ? row[colMapMaster.SCONTO_PERCENTUALE] : "",
        colMapMaster.NOTE !== undefined ? row[colMapMaster.NOTE] : "",
        colMapMaster.DATA_ULTIMO_ORDINE !== undefined ? row[colMapMaster.DATA_ULTIMO_ORDINE] : "",
        colMapMaster.PRIORITA_URGENTE !== undefined ? row[colMapMaster.PRIORITA_URGENTE] : "",
        colMapMaster.PERFORMANCE_SCORE !== undefined ? row[colMapMaster.PERFORMANCE_SCORE] : "",
        new Date()  // ULTIMA_SYNC
      ];
      
      fornitoriDaSincronizzare.push(rowSync);
    });
    
    // 6. Scrivi nel tab locale
    var ssLocale = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSync = ssLocale.getSheetByName(SYNC_CONFIG.TAB_SYNC);
    
    // Crea tab se non esiste
    if (!sheetSync) {
      sheetSync = ssLocale.insertSheet(SYNC_CONFIG.TAB_SYNC);
      sheetSync.setTabColor("#8E44AD"); // Viola
      logSistema("✅ SYNC: Creato tab " + SYNC_CONFIG.TAB_SYNC);
    }
    
    // Pulisci e riscrivi
    sheetSync.clear();
    
    // Header
    sheetSync.getRange(1, 1, 1, SYNC_CONFIG.COLONNE_SYNC.length)
      .setValues([SYNC_CONFIG.COLONNE_SYNC])
      .setFontWeight("bold")
      .setBackground("#EFEFEF");
    sheetSync.setFrozenRows(1);
    
    // Dati
    if (fornitoriDaSincronizzare.length > 0) {
      sheetSync.getRange(2, 1, fornitoriDaSincronizzare.length, SYNC_CONFIG.COLONNE_SYNC.length)
        .setValues(fornitoriDaSincronizzare);
    }
    
    var durata = Math.round((new Date() - startTime) / 1000);
    logSistema("✅ SYNC COMPLETATO: " + fornitoriDaSincronizzare.length + " fornitori in " + durata + "s");
    
    // Aggiorna SETUP con timestamp ultimo sync
    aggiornaSetup("FORNITORI_ULTIMO_SYNC", new Date());
    aggiornaSetup("FORNITORI_COUNT_SYNC", fornitoriDaSincronizzare.length);
    
    return { 
      success: true, 
      count: fornitoriDaSincronizzare.length, 
      durata: durata,
      error: null 
    };
    
  } catch (e) {
    logSistema("❌ SYNC ERRORE: " + e.toString());
    return { success: false, count: 0, error: e.toString() };
  }
}

/**
 * Estrae dominio da email
 * @param {String} email
 * @returns {String} dominio o stringa vuota
 */
function estraiDominio(email) {
  if (!email || email === "") return "";
  email = email.toString().toLowerCase().trim();
  var parts = email.split("@");
  return parts.length > 1 ? parts[1] : "";
}

// ═══════════════════════════════════════════════════════════════════════
// LOOKUP FORNITORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cerca fornitore per indirizzo email
 * Usa cache locale FORNITORI_SYNC per performance
 * @param {String} email - Email da cercare
 * @returns {Object|null} Dati fornitore o null se non trovato
 */
function lookupFornitoreByEmail(email) {
  if (!email || email === "") return null;
  
  email = email.toLowerCase().trim();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      // Tab sync vuoto, prova sync
      logSistema("⚠️ LOOKUP: Tab FORNITORI_SYNC vuoto, eseguo sync...");
      var syncResult = syncFornitoriCompleto();
      if (!syncResult.success || syncResult.count === 0) {
        return null;
      }
      sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    // Mappa colonne
    var colEmail = headers.indexOf("EMAIL_PRINCIPALE");
    var colEmailSec = headers.indexOf("EMAIL_SECONDARIE");
    var colDominio = headers.indexOf("DOMINIO_EMAIL");
    
    // Estrai dominio dall'email cercata
    var dominioEmail = estraiDominio(email);
    
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      
      // Match email principale
      if (colEmail >= 0) {
        var emailPrincipale = (row[colEmail] || "").toString().toLowerCase().trim();
        if (emailPrincipale === email) {
          return rigaToFornitore(row, headers);
        }
      }
      
      // Match email secondarie (separate da virgola o punto e virgola)
      if (colEmailSec >= 0) {
        var emailSecondarie = (row[colEmailSec] || "").toString().toLowerCase();
        var listaEmail = emailSecondarie.split(/[,;]/).map(function(e) { return e.trim(); });
        if (listaEmail.indexOf(email) >= 0) {
          return rigaToFornitore(row, headers);
        }
      }
      
      // Match dominio
      if (colDominio >= 0 && dominioEmail) {
        var dominioFornitore = (row[colDominio] || "").toString().toLowerCase().trim();
        if (dominioFornitore && dominioFornitore === dominioEmail) {
          return rigaToFornitore(row, headers);
        }
      }
    }
    
    return null;
    
  } catch (e) {
    logSistema("❌ LOOKUP ERRORE: " + e.toString());
    return null;
  }
}

/**
 * Converte riga array in oggetto fornitore
 */
function rigaToFornitore(row, headers) {
  var fornitore = {};
  headers.forEach(function(header, idx) {
    fornitore[header] = row[idx];
  });
  return fornitore;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI PER EMAIL ENGINE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se email proviene da fornitore noto
 * Usato per bypass spam filter
 * @param {String} email - Email mittente
 * @returns {Boolean}
 */
function isFornitoreNoto(email) {
  var fornitore = lookupFornitoreByEmail(email);
  return fornitore !== null;
}

/**
 * Ottiene contesto fornitore per arricchire analisi AI
 * @param {String} email - Email mittente
 * @returns {String} Contesto testuale o stringa vuota
 */
function getContestoFornitore(email) {
  var fornitore = lookupFornitoreByEmail(email);
  
  if (!fornitore) return "";
  
  var contesto = [];
  
  contesto.push("FORNITORE NOTO: " + (fornitore.RAGIONE_SOCIALE || "N/D"));
  
  if (fornitore.SCONTO_BASE) {
    contesto.push("Sconto base accordato: " + fornitore.SCONTO_BASE + "%");
  }
  
  if (fornitore.ULTIMO_ORDINE) {
    contesto.push("Ultimo ordine: " + fornitore.ULTIMO_ORDINE);
  }
  
  if (fornitore.PRIORITA) {
    contesto.push("Priorità: " + fornitore.PRIORITA);
  }
  
  if (fornitore.NOTE_INTERNE) {
    contesto.push("Note: " + fornitore.NOTE_INTERNE);
  }
  
  return contesto.join(" | ");
}

/**
 * Verifica se email deve bypassare spam filter
 * @param {String} email - Email mittente
 * @returns {Object} {bypass: boolean, reason: string}
 */
function checkBypassSpamFilter(email) {
  var fornitore = lookupFornitoreByEmail(email);
  
  if (fornitore) {
    return {
      bypass: true,
      reason: "Fornitore noto: " + (fornitore.RAGIONE_SOCIALE || fornitore.ID_FORNITORE)
    };
  }
  
  return {
    bypass: false,
    reason: null
  };
}

// ═══════════════════════════════════════════════════════════════════════
// MENU E TEST
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test sync da menu
 */
function menuTestSyncFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🔄 Test Sync Fornitori',
    'Questo test eseguirà:\n\n' +
    '1. Connessione al Fornitori Engine\n' +
    '2. Lettura fornitori attivi\n' +
    '3. Creazione/aggiornamento tab FORNITORI_SYNC\n\n' +
    'Assicurati di aver configurato ID_FORNITORI_SOURCE in SETUP.\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  var result = syncFornitoriCompleto();
  
  if (result.success) {
    ui.alert(
      '✅ Sync Completato!',
      'Fornitori sincronizzati: ' + result.count + '\n' +
      'Durata: ' + result.durata + ' secondi\n\n' +
      'Controlla il tab FORNITORI_SYNC per verificare i dati.',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '❌ Sync Fallito',
      'Errore: ' + result.error + '\n\n' +
      'Verifica:\n' +
      '1. ID_FORNITORI_SOURCE in SETUP\n' +
      '2. Permessi di accesso al Fornitori Engine\n' +
      '3. Esistenza foglio FORNITORI nel master',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Test lookup fornitore da menu
 */
function menuTestLookupFornitore() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.prompt(
    '🔍 Test Lookup Fornitore',
    'Inserisci email da cercare:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var email = risposta.getResponseText().trim();
  
  if (!email) {
    ui.alert('⚠️ Email non inserita');
    return;
  }
  
  var fornitore = lookupFornitoreByEmail(email);
  
  if (fornitore) {
    var info = '✅ FORNITORE TROVATO!\n\n';
    info += 'ID: ' + (fornitore.ID_FORNITORE || 'N/D') + '\n';
    info += 'Ragione Sociale: ' + (fornitore.RAGIONE_SOCIALE || 'N/D') + '\n';
    info += 'Email: ' + (fornitore.EMAIL_PRINCIPALE || 'N/D') + '\n';
    info += 'Dominio: ' + (fornitore.DOMINIO_EMAIL || 'N/D') + '\n';
    info += 'Sconto Base: ' + (fornitore.SCONTO_BASE || '0') + '%\n';
    info += 'Status: ' + (fornitore.STATUS || 'N/D') + '\n';
    info += 'Priorità: ' + (fornitore.PRIORITA || 'N/D') + '\n';
    
    ui.alert('Risultato Lookup', info, ui.ButtonSet.OK);
    logSistema("LOOKUP OK: " + email + " → " + fornitore.RAGIONE_SOCIALE);
  } else {
    ui.alert(
      '❌ Fornitore Non Trovato',
      'Nessun fornitore trovato per: ' + email + '\n\n' +
      'Verifiche:\n' +
      '1. Email corretta?\n' +
      '2. Fornitore presente nel master?\n' +
      '3. Tab FORNITORI_SYNC aggiornato?',
      ui.ButtonSet.OK
    );
    logSistema("LOOKUP MISS: " + email);
  }
}

/**
 * Test bypass spam filter
 */
function menuTestBypassSpam() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.prompt(
    '🛡️ Test Bypass Spam Filter',
    'Inserisci email mittente da verificare:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var email = risposta.getResponseText().trim();
  
  if (!email) {
    ui.alert('⚠️ Email non inserita');
    return;
  }
  
  var result = checkBypassSpamFilter(email);
  
  if (result.bypass) {
    ui.alert(
      '✅ BYPASS ATTIVO',
      'Email: ' + email + '\n\n' +
      'Questa email BYPASSA il filtro spam.\n' +
      'Motivo: ' + result.reason + '\n\n' +
      'L\'email andrà direttamente a Layer 1 (analisi).',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '🛡️ NESSUN BYPASS',
      'Email: ' + email + '\n\n' +
      'Questa email passerà dal filtro spam (Layer 0).',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Mostra statistiche sync
 */
function menuStatisticheSync() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var ultimoSync = getSetupValue("FORNITORI_ULTIMO_SYNC", "Mai");
  var countSync = getSetupValue("FORNITORI_COUNT_SYNC", 0);
  var idFornitori = getSetupValue(SYNC_CONFIG.KEY_ID_FORNITORI, "Non configurato");
  
  var sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
  var righeLocali = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  
  var stats = '📊 STATISTICHE SYNC FORNITORI\n\n' +
    '━━━ Configurazione ━━━\n' +
    'ID Fornitori Engine: ' + (idFornitori.length > 20 ? idFornitori.substring(0, 20) + '...' : idFornitori) + '\n' +
    'Tab locale: ' + SYNC_CONFIG.TAB_SYNC + '\n\n' +
    '━━━ Ultimo Sync ━━━\n' +
    'Data: ' + ultimoSync + '\n' +
    'Fornitori sincronizzati: ' + countSync + '\n' +
    'Righe in cache locale: ' + righeLocali + '\n\n' +
    '━━━ Colonne Sync ━━━\n' +
    SYNC_CONFIG.COLONNE_SYNC.join(', ');
  
  ui.alert('Statistiche Sync', stats, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES (se non già presenti)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Leggi valore da SETUP
 */
function getSetupValue(key, defaultValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SETUP");
    if (!sheet) return defaultValue;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1] !== "" ? data[i][1] : defaultValue;
      }
    }
    return defaultValue;
  } catch (e) {
    return defaultValue;
  }
}

/**
 * Aggiorna valore in SETUP
 */
function aggiornaSetup(key, value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SETUP");
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
    // Se non trovato, aggiungi
    sheet.appendRow([key, value]);
  } catch (e) {
    Logger.log("Errore aggiornaSetup: " + e);
  }
}

/**
 * Log sistema (se non già presente)
 */
function logSistema(messaggio) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LOG_SISTEMA");
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}

/**
 * ==========================================================================================
 * AGGIUNGI QUESTO A SyncFornitori.gs
 * ==========================================================================================
 * Funzioni richieste da Logic.gs v1.3.0 per integrazione bypass fornitori
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI RICHIESTE DA LOGIC.gs
// ═══════════════════════════════════════════════════════════════════════

/**
 * Pre-check email per bypass Layer 0 (Spam Filter)
 * Chiamata da Logic.gs prima di eseguire L0
 * @param {String} mittente - Email mittente
 * @returns {Object} {isFornitoreNoto, skipL0, fornitore, contesto}
 */
function preCheckEmailFornitore(mittente) {
  if (!mittente || mittente === "") {
    return { isFornitoreNoto: false, skipL0: false, fornitore: null, contesto: "" };
  }
  
  var fornitore = lookupFornitoreByEmail(mittente);
  
  if (fornitore) {
    return {
      isFornitoreNoto: true,
      skipL0: true,  // Skip spam filter per fornitori noti
      fornitore: {
        idFornitore: fornitore.ID_FORNITORE,
        nomeAzienda: fornitore.RAGIONE_SOCIALE,
        email: fornitore.EMAIL_PRINCIPALE,
        sconto: fornitore.SCONTO_BASE,
        status: fornitore.STATUS
      },
      contesto: getContestoFornitore(mittente)
    };
  }
  
  return { isFornitoreNoto: false, skipL0: false, fornitore: null, contesto: "" };
}

/**
 * Post-analisi: aggiorna dati fornitore dopo analisi email
 * Chiamata da Logic.gs dopo L3
 * @param {String} mittente - Email mittente
 * @param {Object} risultato - Risultato analisi L3
 * @param {String} idEmail - ID email analizzata
 */
function postAnalisiEmailFornitore(mittente, risultato, idEmail) {
  // Per ora solo logging - in futuro può aggiornare il master Fornitori
  var fornitore = lookupFornitoreByEmail(mittente);
  
  if (fornitore) {
    logSistema("📊 Post-analisi fornitore: " + fornitore.RAGIONE_SOCIALE + 
               " | Email: " + idEmail + 
               " | Confidence: " + risultato.confidence + "%");
    
    // TODO: Aggiornare campi nel master Fornitori Engine
    // - DATA_ULTIMA_EMAIL
    // - EMAIL_ANALIZZATE_COUNT++
    // - etc.
  }
}

/**
 * Processa azioni suggerite relative ai fornitori
 * Chiamata da Logic.gs dopo L3 se ci sono azioni
 * @param {String} idFornitore - ID fornitore
 * @param {Array} azioni - Array di azioni suggerite
 * @param {String} idEmail - ID email
 */
function processaAzioniSuggeriteFornitori(idFornitore, azioni, idEmail) {
  if (!azioni || azioni.length === 0) return;
  
  azioni.forEach(function(azione) {
    switch (azione) {
      case "FLAG_FATTURA":
      case "FLAG_FATTURA_FORNITORE":
        logSistema("📄 AZIONE: Fattura da " + idFornitore + " (Email: " + idEmail + ")");
        // TODO: Creare record in foglio FATTURE o notificare
        break;
        
      case "FLAG_NOVITA":
      case "FLAG_NOVITA_FORNITORE":
        logSistema("🆕 AZIONE: Novità da " + idFornitore + " (Email: " + idEmail + ")");
        // TODO: Flag nel master fornitori
        break;
        
      case "FLAG_CATALOGO":
      case "FLAG_CATALOGO_FORNITORE":
        logSistema("📚 AZIONE: Catalogo da " + idFornitore + " (Email: " + idEmail + ")");
        break;
        
      case "FLAG_PREZZI":
      case "FLAG_PREZZI_FORNITORE":
        logSistema("💰 AZIONE: Aggiornamento prezzi da " + idFornitore + " (Email: " + idEmail + ")");
        break;
        
      case "CREA_QUESTIONE_PROBLEMA":
        logSistema("⚠️ AZIONE: Problema da " + idFornitore + " - RICHIEDE ATTENZIONE (Email: " + idEmail + ")");
        // TODO: Creare ticket/questione
        break;
        
      case "ALERT_URGENTE":
        logSistema("🚨 AZIONE URGENTE da " + idFornitore + " (Email: " + idEmail + ")");
        // TODO: Notifica immediata
        break;
    }
  });
}

/**
 * Verifica se il connettore Fornitori è attivo
 * Chiamata da Logic.gs per check disponibilità
 * @returns {Boolean}
 */
function isConnectorFornitoriAttivo() {
  // Verifica che il tab FORNITORI_SYNC esista e abbia dati
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
    
    if (!sheet) return false;
    if (sheet.getLastRow() <= 1) return false;
    
    // Verifica che ID_FORNITORI_SOURCE sia configurato
    var idFornitori = getSetupValue(SYNC_CONFIG.KEY_ID_FORNITORI, "");
    if (!idFornitori || idFornitori === "" || idFornitori === "Incolla qui ID Fornitori") {
      return false;
    }
    
    return true;
  } catch (e) {
    return false;
  }
}