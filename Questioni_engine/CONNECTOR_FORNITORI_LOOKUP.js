/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - LOOKUP v1.0.0 (per Questioni Engine)
 * ==========================================================================================
 * Funzioni di ricerca fornitori per associare alle questioni
 * 
 * RICICLATO: Basato su connector fornitori di Email Engine, adattato
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LOOKUP PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cerca fornitore per email mittente
 * @param {String} emailMittente - Email da cercare
 * @returns {Object|null} Oggetto fornitore o null
 */
function lookupFornitoreQEByEmail(emailMittente) {
  if (!emailMittente) return null;
  
  emailMittente = emailMittente.toLowerCase().trim();
  
  // Estrai dominio
  var dominio = "";
  var atIndex = emailMittente.indexOf("@");
  if (atIndex > 0) {
    dominio = emailMittente.substring(atIndex + 1);
  }
  
  // Carica dati (con cache)
  var dati = caricaDatiFornitoriQELocali();
  if (!dati || !dati.rows || dati.rows.length === 0) {
    return null;
  }
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  // Prima passata: match esatto email
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var status = (row[colMap.STATUS_FORNITORE] || "").toUpperCase();
    if (status === "BLACKLIST") continue;
    
    // Check EMAIL_ORDINI
    var emailOrdini = (row[colMap.EMAIL_ORDINI] || "").toLowerCase().trim();
    if (emailOrdini === emailMittente) {
      return costruisciOggettoFornitoreQE(row, colMap, i + 2);
    }
    
    // Check EMAIL_ALTRI (lista separata da virgola)
    var emailAltri = (row[colMap.EMAIL_ALTRI] || "").toLowerCase();
    if (emailAltri) {
      var lista = emailAltri.split(",");
      for (var j = 0; j < lista.length; j++) {
        if (lista[j].trim() === emailMittente) {
          return costruisciOggettoFornitoreQE(row, colMap, i + 2);
        }
      }
    }
  }
  
  // Seconda passata: match per dominio
  if (dominio) {
    for (var k = 0; k < rows.length; k++) {
      var row2 = rows[k];
      
      if (!row2[colMap.ID_FORNITORE]) continue;
      
      var status2 = (row2[colMap.STATUS_FORNITORE] || "").toUpperCase();
      if (status2 === "BLACKLIST") continue;
      
      var emailOrd = (row2[colMap.EMAIL_ORDINI] || "").toLowerCase().trim();
      if (emailOrd && estraiDominioQE(emailOrd) === dominio) {
        return costruisciOggettoFornitoreQE(row2, colMap, k + 2);
      }
      
      var emailAlt = (row2[colMap.EMAIL_ALTRI] || "").toLowerCase();
      if (emailAlt) {
        var listaAlt = emailAlt.split(",");
        for (var m = 0; m < listaAlt.length; m++) {
          if (estraiDominioQE(listaAlt[m].trim()) === dominio) {
            return costruisciOggettoFornitoreQE(row2, colMap, k + 2);
          }
        }
      }
    }
  }
  
  return null;
}

/**
 * Cerca fornitore per ID
 * @param {String} idFornitore - ID (es. FOR-001)
 * @returns {Object|null}
 */
function lookupFornitoreQEById(idFornitore) {
  if (!idFornitore) return null;
  
  var dati = caricaDatiFornitoriQELocali();
  if (!dati) return null;
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][colMap.ID_FORNITORE] === idFornitore) {
      return costruisciOggettoFornitoreQE(rows[i], colMap, i + 2);
    }
  }
  
  return null;
}

/**
 * Cerca fornitori per nome (match parziale)
 * @param {String} query - Testo da cercare
 * @returns {Array} Lista fornitori trovati
 */
function lookupFornitoriQEByNome(query) {
  if (!query) return [];
  
  query = query.toLowerCase().trim();
  var risultati = [];
  
  var dati = caricaDatiFornitoriQELocali();
  if (!dati) return [];
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var nome = (row[colMap.NOME_AZIENDA] || "").toLowerCase();
    
    if (nome.indexOf(query) >= 0) {
      risultati.push(costruisciOggettoFornitoreQE(row, colMap, i + 2));
    }
  }
  
  return risultati;
}

/**
 * Ottiene lista di tutti i fornitori (per dropdown/autocomplete)
 * @returns {Array} [{id, nome, email}, ...]
 */
function getListaFornitoriQE() {
  var dati = caricaDatiFornitoriQELocali();
  if (!dati) return [];
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  var lista = [];
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var status = (row[colMap.STATUS_FORNITORE] || "").toUpperCase();
    if (status === "BLACKLIST") continue;
    
    lista.push({
      id: row[colMap.ID_FORNITORE],
      nome: row[colMap.NOME_AZIENDA] || "",
      email: row[colMap.EMAIL_ORDINI] || "",
      status: row[colMap.STATUS_FORNITORE] || ""
    });
  }
  
  // Ordina per nome
  lista.sort(function(a, b) {
    return (a.nome || "").localeCompare(b.nome || "");
  });
  
  return lista;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONE PRINCIPALE PER QUESTIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Associa fornitore a una questione
 * Cerca fornitore tramite email (dal campo EMAIL_COLLEGATE della questione)
 * 
 * @param {String} dominioEmail - Dominio email (es. "fornitore.it")
 * @param {String} emailCompleta - Email completa (opzionale, prioritaria)
 * @returns {Object} {
 *   trovato: Boolean,
 *   idFornitore: String,
 *   nomeFornitore: String,
 *   fornitore: Object|null
 * }
 */
function associaFornitoreAQuestione(dominioEmail, emailCompleta) {
  var risultato = {
    trovato: false,
    idFornitore: "",
    nomeFornitore: "",
    fornitore: null
  };
  
  // Verifica connettore attivo
  if (!isConnectorFornitoriQEAttivo()) {
    logConnectorFornitoriQE("Connettore non attivo, skip associazione");
    return risultato;
  }
  
  // Auto-sync se necessario
  autoSyncFornitoriQESeNecessario();
  
  // Prova prima con email completa
  var fornitore = null;
  
  if (emailCompleta) {
    fornitore = lookupFornitoreQEByEmail(emailCompleta);
  }
  
  // Se non trovato, prova con dominio
  if (!fornitore && dominioEmail) {
    // Crea email fittiza per match dominio
    fornitore = lookupFornitoreQEByEmail("info@" + dominioEmail);
  }
  
  if (fornitore) {
    risultato.trovato = true;
    risultato.idFornitore = fornitore.id;
    risultato.nomeFornitore = fornitore.nome;
    risultato.fornitore = fornitore;
    
    logConnectorFornitoriQE("Trovato: " + fornitore.nome + " (" + fornitore.id + ")");
  }
  
  return risultato;
}

/**
 * Genera contesto fornitore per arricchire analisi questione
 * @param {Object} fornitore - Oggetto fornitore
 * @returns {String} Contesto formattato
 */
function generaContestoFornitoreQE(fornitore) {
  if (!fornitore) return "";
  
  var linee = [];
  
  linee.push("═══ FORNITORE ═══");
  linee.push("Nome: " + fornitore.nome);
  linee.push("ID: " + fornitore.id);
  
  if (fornitore.prioritaUrgente) {
    linee.push("⚠️ PRIORITA' URGENTE");
  }
  
  if (fornitore.performanceScore) {
    var stelle = "";
    for (var i = 0; i < fornitore.performanceScore; i++) stelle += "★";
    for (var j = fornitore.performanceScore; j < 5; j++) stelle += "☆";
    linee.push("Performance: " + stelle);
  }
  
  if (fornitore.statusUltimaAzione) {
    linee.push("Ultimo contatto: " + fornitore.statusUltimaAzione);
  }
  
  linee.push("═════════════════");
  
  return linee.join("\n");
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER INTERNI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica dati fornitori da tab locale (con cache)
 * @returns {Object|null} {headers, rows, colMap}
 */
function caricaDatiFornitoriQELocali() {
  var cache = CONNECTOR_FORNITORI_QE._cache;
  var now = new Date().getTime();
  
  if (cache.fornitori && cache.timestamp) {
    var age = (now - cache.timestamp) / 1000;
    if (age < CONNECTOR_FORNITORI_QE.DEFAULTS.CACHE_TTL_SEC) {
      return cache.fornitori;
    }
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_SYNC);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return null;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    var colMap = {};
    Object.keys(CONNECTOR_FORNITORI_QE.COLONNE).forEach(function(key) {
      var nomeColonna = CONNECTOR_FORNITORI_QE.COLONNE[key];
      var idx = headers.indexOf(nomeColonna);
      if (idx >= 0) {
        colMap[key] = idx;
      }
    });
    
    var result = {
      headers: headers,
      rows: rows,
      colMap: colMap
    };
    
    CONNECTOR_FORNITORI_QE._cache.fornitori = result;
    CONNECTOR_FORNITORI_QE._cache.timestamp = now;
    
    return result;
    
  } catch(e) {
    logConnectorFornitoriQE("Errore caricamento dati locali: " + e.message);
    return null;
  }
}

/**
 * Costruisce oggetto fornitore da riga
 */
function costruisciOggettoFornitoreQE(row, colMap, rowNum) {
  return {
    id: row[colMap.ID_FORNITORE] || "",
    nome: row[colMap.NOME_AZIENDA] || "",
    email: row[colMap.EMAIL_ORDINI] || "",
    emailAltri: row[colMap.EMAIL_ALTRI] || "",
    contatto: row[colMap.CONTATTO] || "",
    telefono: row[colMap.TELEFONO] || "",
    status: row[colMap.STATUS_FORNITORE] || "",
    prioritaUrgente: row[colMap.PRIORITA_URGENTE] === "X" || row[colMap.PRIORITA_URGENTE] === true,
    scontoPercentuale: parseFloat(row[colMap.SCONTO_PERCENTUALE]) || 0,
    statusUltimaAzione: row[colMap.STATUS_ULTIMA_AZIONE] || "",
    performanceScore: parseInt(row[colMap.PERFORMANCE_SCORE]) || 0,
    rowNum: rowNum,
    _source: "FORNITORI_SYNC"
  };
}

/**
 * Estrae dominio da email
 */
function estraiDominioQE(email) {
  if (!email) return "";
  email = email.toLowerCase().trim();
  var at = email.indexOf("@");
  return at > 0 ? email.substring(at + 1) : "";
}

/**
 * Conta fornitori nel sync locale
 */
function contaFornitoriQELocali() {
  var dati = caricaDatiFornitoriQELocali();
  return dati ? dati.rows.length : 0;
}