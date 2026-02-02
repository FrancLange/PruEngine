/**
 * ==========================================================================================
 * QUESTIONI_HELPERS.js - Utility e Funzioni Helper v1.0.0 STANDALONE
 * ==========================================================================================
 * Funzioni di supporto per il motore questioni
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LOOKUP QUESTIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cerca questione per ID
 * @param {String} idQuestione - ID questione (es. QST-001)
 * @returns {Object|null}
 */
function getQuestioneById(idQuestione) {
  if (!idQuestione) return null;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) return null;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colMap = {};
    Object.keys(QUESTIONI_CONFIG.COLONNE_QUESTIONI).forEach(function(key) {
      var nome = QUESTIONI_CONFIG.COLONNE_QUESTIONI[key];
      var idx = headers.indexOf(nome);
      if (idx >= 0) colMap[key] = idx;
    });
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][colMap.ID_QUESTIONE] === idQuestione) {
        return costruisciOggettoQuestione(data[i], colMap, i + 1);
      }
    }
    
    return null;
    
  } catch(e) {
    logQuestioni("Errore getQuestioneById: " + e.message);
    return null;
  }
}

/**
 * Ottieni tutte le questioni aperte
 * @returns {Array}
 */
function getQuestioniAperte() {
  return getQuestioniByStatus([
    QUESTIONI_CONFIG.STATUS_QUESTIONE.APERTA,
    QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_LAVORAZIONE,
    QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_ATTESA_RISPOSTA
  ]);
}

/**
 * Ottieni questioni per status
 * @param {Array} statusList - Lista status da cercare
 * @returns {Array}
 */
function getQuestioniByStatus(statusList) {
  var questioni = [];
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) return questioni;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colMap = {};
    Object.keys(QUESTIONI_CONFIG.COLONNE_QUESTIONI).forEach(function(key) {
      var nome = QUESTIONI_CONFIG.COLONNE_QUESTIONI[key];
      var idx = headers.indexOf(nome);
      if (idx >= 0) colMap[key] = idx;
    });
    
    for (var i = 1; i < data.length; i++) {
      var status = data[i][colMap.STATUS];
      
      if (statusList.indexOf(status) >= 0) {
        questioni.push(costruisciOggettoQuestione(data[i], colMap, i + 1));
      }
    }
    
  } catch(e) {
    logQuestioni("Errore getQuestioniByStatus: " + e.message);
  }
  
  return questioni;
}

/**
 * Ottieni questioni per fornitore
 * @param {String} emailFornitore - Email o dominio fornitore
 * @returns {Array}
 */
function getQuestioniByFornitore(emailFornitore) {
  var questioni = [];
  
  if (!emailFornitore) return questioni;
  
  var dominio = estraiDominioEmail(emailFornitore);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) return questioni;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colMap = {};
    Object.keys(QUESTIONI_CONFIG.COLONNE_QUESTIONI).forEach(function(key) {
      var nome = QUESTIONI_CONFIG.COLONNE_QUESTIONI[key];
      var idx = headers.indexOf(nome);
      if (idx >= 0) colMap[key] = idx;
    });
    
    for (var i = 1; i < data.length; i++) {
      var emailQ = data[i][colMap.EMAIL_FORNITORE] || "";
      var dominioQ = estraiDominioEmail(emailQ);
      
      if (emailQ === emailFornitore || dominioQ === dominio) {
        questioni.push(costruisciOggettoQuestione(data[i], colMap, i + 1));
      }
    }
    
  } catch(e) {
    logQuestioni("Errore getQuestioniByFornitore: " + e.message);
  }
  
  return questioni;
}

/**
 * Costruisce oggetto questione da riga
 */
function costruisciOggettoQuestione(row, colMap, rowNum) {
  return {
    rowNum: rowNum,
    id: row[colMap.ID_QUESTIONE] || "",
    timestampCreazione: row[colMap.TIMESTAMP_CREAZIONE],
    fonte: row[colMap.FONTE] || "",
    
    idFornitore: row[colMap.ID_FORNITORE] || "",
    nomeFornitore: row[colMap.NOME_FORNITORE] || "",
    emailFornitore: row[colMap.EMAIL_FORNITORE] || "",
    
    titolo: row[colMap.TITOLO] || "",
    descrizione: row[colMap.DESCRIZIONE] || "",
    categoria: row[colMap.CATEGORIA] || "",
    tags: row[colMap.TAGS] || "",
    
    priorita: row[colMap.PRIORITA] || "",
    urgente: row[colMap.URGENTE] === "SI" || row[colMap.URGENTE] === true,
    status: row[colMap.STATUS] || "",
    
    dataScadenza: row[colMap.DATA_SCADENZA],
    assegnatoA: row[colMap.ASSEGNATO_A] || "",
    
    emailCollegate: row[colMap.EMAIL_COLLEGATE] || "",
    numEmail: parseInt(row[colMap.NUM_EMAIL]) || 0,
    
    timestampRisoluzione: row[colMap.TIMESTAMP_RISOLUZIONE],
    risoluzioneNote: row[colMap.RISOLUZIONE_NOTE] || "",
    
    aiConfidence: parseInt(row[colMap.AI_CONFIDENCE]) || 0,
    aiClusterId: row[colMap.AI_CLUSTER_ID] || ""
  };
}

// ═══════════════════════════════════════════════════════════════════════
// AGGIORNAMENTO QUESTIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna campi di una questione
 * @param {String} idQuestione - ID questione
 * @param {Object} updates - Campi da aggiornare
 * @returns {Boolean}
 */
function aggiornaQuestione(idQuestione, updates) {
  if (!idQuestione || !updates) return false;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) return false;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colMap = {};
    Object.keys(QUESTIONI_CONFIG.COLONNE_QUESTIONI).forEach(function(key) {
      var nome = QUESTIONI_CONFIG.COLONNE_QUESTIONI[key];
      var idx = headers.indexOf(nome);
      if (idx >= 0) colMap[key] = idx;
    });
    
    // Trova riga
    var rowNum = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][colMap.ID_QUESTIONE] === idQuestione) {
        rowNum = i + 1;
        break;
      }
    }
    
    if (rowNum < 0) return false;
    
    // Aggiorna campi
    Object.keys(updates).forEach(function(campo) {
      var colIndex = colMap[campo];
      if (colIndex !== undefined && colIndex >= 0) {
        sheet.getRange(rowNum, colIndex + 1).setValue(updates[campo]);
      }
    });
    
    logQuestioni("Aggiornata questione " + idQuestione);
    return true;
    
  } catch(e) {
    logQuestioni("Errore aggiornaQuestione: " + e.message);
    return false;
  }
}

/**
 * Collega email a questione esistente
 * @param {String} idQuestione - ID questione
 * @param {String} idEmail - ID email da collegare
 * @returns {Boolean}
 */
function collegaEmailAQuestione(idQuestione, idEmail) {
  if (!idQuestione || !idEmail) return false;
  
  var questione = getQuestioneById(idQuestione);
  if (!questione) return false;
  
  // Aggiungi email alla lista
  var emailCollegate = questione.emailCollegate || "";
  var listaEmail = emailCollegate ? emailCollegate.split(", ") : [];
  
  if (listaEmail.indexOf(idEmail) < 0) {
    listaEmail.push(idEmail);
  }
  
  // Aggiorna questione
  aggiornaQuestione(idQuestione, {
    EMAIL_COLLEGATE: listaEmail.join(", "),
    NUM_EMAIL: listaEmail.length
  });
  
  // Aggiorna email locale
  collegaEmailAQuestioneLocale(idEmail, idQuestione);
  
  logQuestioni("Collegata email " + idEmail + " a " + idQuestione);
  return true;
}

/**
 * Cambia status questione
 * @param {String} idQuestione - ID questione
 * @param {String} nuovoStatus - Nuovo status
 * @param {String} note - Note opzionali
 * @returns {Boolean}
 */
function cambiaStatusQuestione(idQuestione, nuovoStatus, note) {
  var updates = {
    STATUS: nuovoStatus
  };
  
  if (nuovoStatus === QUESTIONI_CONFIG.STATUS_QUESTIONE.RISOLTA ||
      nuovoStatus === QUESTIONI_CONFIG.STATUS_QUESTIONE.CHIUSA) {
    updates.TIMESTAMP_RISOLUZIONE = new Date();
    if (note) {
      updates.RISOLUZIONE_NOTE = note;
    }
  }
  
  return aggiornaQuestione(idQuestione, updates);
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════════════════

/**
 * Estrai dominio da email
 */
function estraiDominioEmail(email) {
  if (!email) return "";
  email = email.toLowerCase().trim();
  var at = email.indexOf("@");
  return at > 0 ? email.substring(at + 1) : email;
}

/**
 * Raggruppa email per fornitore/dominio
 * @param {Array} emails - Array di email objects
 * @returns {Object} {dominio: [emails]}
 */
function raggruppaEmailPerFornitore(emails) {
  var gruppi = {};
  
  emails.forEach(function(email) {
    var dominio = estraiDominioEmail(email.mittente);
    
    if (!gruppi[dominio]) {
      gruppi[dominio] = [];
    }
    
    gruppi[dominio].push(email);
  });
  
  return gruppi;
}

/**
 * Determina categoria da tags e scores
 */
function determinaCategoria(tags, scores) {
  tags = tags || "";
  scores = scores || {};
  
  var tagsArray = typeof tags === 'string' ? tags.split(", ") : tags;
  
  if (tagsArray.indexOf("PROBLEMA") >= 0 || scores.problema > 70) {
    if (tagsArray.indexOf("ORDINE") >= 0 || scores.ordine > 50) {
      return QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.PROBLEMA_ORDINE;
    }
    if (tagsArray.indexOf("FATTURA") >= 0 || scores.fattura > 50) {
      return QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.PROBLEMA_FATTURAZIONE;
    }
    return QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.RECLAMO;
  }
  
  if (scores.urgente > 80) {
    return QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.URGENZA;
  }
  
  return QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.ALTRO;
}

/**
 * Determina priorità da scores
 */
function determinaPriorita(scores) {
  scores = scores || {};
  
  var urgente = parseInt(scores.urgente) || 0;
  var problema = parseInt(scores.problema) || 0;
  
  if (urgente >= 90 || problema >= 90) {
    return QUESTIONI_CONFIG.PRIORITA.CRITICA;
  }
  
  if (urgente >= 70 || problema >= 80) {
    return QUESTIONI_CONFIG.PRIORITA.ALTA;
  }
  
  if (urgente >= 50 || problema >= 60) {
    return QUESTIONI_CONFIG.PRIORITA.MEDIA;
  }
  
  return QUESTIONI_CONFIG.PRIORITA.BASSA;
}

/**
 * Verifica se è urgente
 */
function isUrgente(scores) {
  scores = scores || {};
  var urgente = parseInt(scores.urgente) || 0;
  return urgente >= 80;
}

// ═══════════════════════════════════════════════════════════════════════
// PROMPT HELPERS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica prompt da foglio PROMPTS
 * @param {String} key - Chiave prompt
 * @returns {Object|null}
 */
function caricaPromptQuestioni(key) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.PROMPTS);
    
    if (!sheet || sheet.getLastRow() <= 1) return null;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colKey = headers.indexOf(QUESTIONI_CONFIG.COLONNE_PROMPTS.KEY);
    var colTesto = headers.indexOf(QUESTIONI_CONFIG.COLONNE_PROMPTS.TESTO);
    var colTemp = headers.indexOf(QUESTIONI_CONFIG.COLONNE_PROMPTS.TEMPERATURA);
    var colTokens = headers.indexOf(QUESTIONI_CONFIG.COLONNE_PROMPTS.MAX_TOKENS);
    var colAttivo = headers.indexOf(QUESTIONI_CONFIG.COLONNE_PROMPTS.ATTIVO);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[colKey] === key && row[colAttivo] === "SI") {
        return {
          key: row[colKey],
          testo: row[colTesto],
          temperatura: parseFloat(row[colTemp]) || 0.7,
          maxTokens: parseInt(row[colTokens]) || 500
        };
      }
    }
    
    return null;
    
  } catch(e) {
    logQuestioni("Errore caricaPromptQuestioni: " + e.message);
    return null;
  }
}