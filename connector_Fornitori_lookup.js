/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - LOOKUP v1.0.0
 * ==========================================================================================
 * Funzioni di ricerca nella tab locale FORNITORI_SYNC
 * 
 * AUTONOMO: Cerca solo nella copia locale, veloce e senza chiamate cross-sheet
 * ==========================================================================================
 */

/**
 * MAIN: Cerca fornitore per email mittente
 * 
 * Cerca in ordine:
 * 1. EMAIL_ORDINI (match esatto)
 * 2. EMAIL_ALTRI (lista separata da virgola)
 * 3. Match per dominio (se abilitato)
 * 
 * @param {String} emailMittente - Email da cercare
 * @returns {Object|null} Oggetto fornitore o null se non trovato
 */
function lookupFornitoreByEmail(emailMittente) {
  if (!emailMittente) return null;
  
  emailMittente = emailMittente.toLowerCase().trim();
  
  // Estrai dominio per match secondario
  var dominio = "";
  var atIndex = emailMittente.indexOf("@");
  if (atIndex > 0) {
    dominio = emailMittente.substring(atIndex + 1);
  }
  
  // Carica dati (con cache)
  var dati = caricaDatiFornitoriLocali();
  if (!dati || !dati.rows || dati.rows.length === 0) {
    return null;
  }
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  // Prima passata: match esatto email
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    
    // Skip righe vuote
    if (!row[colMap.ID_FORNITORE]) continue;
    
    // Skip BLACKLIST
    var status = (row[colMap.STATUS_FORNITORE] || "").toUpperCase();
    if (status === "BLACKLIST") continue;
    
    // Check EMAIL_ORDINI
    var emailOrdini = (row[colMap.EMAIL_ORDINI] || "").toLowerCase().trim();
    if (emailOrdini === emailMittente) {
      return costruisciOggettoFornitore(row, colMap, i + 2);
    }
    
    // Check EMAIL_ALTRI (lista separata da virgola)
    var emailAltri = (row[colMap.EMAIL_ALTRI] || "").toLowerCase();
    if (emailAltri) {
      var lista = emailAltri.split(",");
      for (var j = 0; j < lista.length; j++) {
        if (lista[j].trim() === emailMittente) {
          return costruisciOggettoFornitore(row, colMap, i + 2);
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
      
      // Dominio da EMAIL_ORDINI
      var emailOrd = (row2[colMap.EMAIL_ORDINI] || "").toLowerCase().trim();
      if (emailOrd && estraiDominio(emailOrd) === dominio) {
        return costruisciOggettoFornitore(row2, colMap, k + 2);
      }
      
      // Dominio da EMAIL_ALTRI
      var emailAlt = (row2[colMap.EMAIL_ALTRI] || "").toLowerCase();
      if (emailAlt) {
        var listaAlt = emailAlt.split(",");
        for (var m = 0; m < listaAlt.length; m++) {
          if (estraiDominio(listaAlt[m].trim()) === dominio) {
            return costruisciOggettoFornitore(row2, colMap, k + 2);
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
function lookupFornitoreById(idFornitore) {
  if (!idFornitore) return null;
  
  var dati = caricaDatiFornitoriLocali();
  if (!dati) return null;
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][colMap.ID_FORNITORE] === idFornitore) {
      return costruisciOggettoFornitore(rows[i], colMap, i + 2);
    }
  }
  
  return null;
}

/**
 * Cerca fornitori per nome (match parziale)
 * @param {String} query - Testo da cercare
 * @returns {Array} Lista fornitori trovati
 */
function lookupFornitoriByNome(query) {
  if (!query) return [];
  
  query = query.toLowerCase().trim();
  var risultati = [];
  
  var dati = caricaDatiFornitoriLocali();
  if (!dati) return [];
  
  var colMap = dati.colMap;
  var rows = dati.rows;
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var nome = (row[colMap.NOME_AZIENDA] || "").toLowerCase();
    
    if (nome.indexOf(query) >= 0) {
      risultati.push(costruisciOggettoFornitore(row, colMap, i + 2));
    }
  }
  
  return risultati;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONE PRINCIPALE PER EMAIL ENGINE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Pre-Check email per Email Engine
 * 
 * Questa è la FUNZIONE PRINCIPALE chiamata da Filters.gs / Logic.gs
 * 
 * @param {String} emailMittente - Email da verificare
 * @returns {Object} {
 *   isFornitoreNoto: Boolean,
 *   skipL0: Boolean,
 *   contesto: String,
 *   fornitore: Object|null
 * }
 */
function preCheckEmailFornitore(emailMittente) {
  var risultato = {
    isFornitoreNoto: false,
    skipL0: false,
    contesto: "",
    fornitore: null
  };
  
  // Verifica connettore attivo
  if (!isConnectorFornitoriAttivo()) {
    return risultato;
  }
  
  // Auto-sync se necessario (silenzioso, non blocca)
  autoSyncFornitoriSeNecessario();
  
  // Lookup
  var fornitore = lookupFornitoreByEmail(emailMittente);
  
  if (!fornitore) {
    return risultato;
  }
  
  // Fornitore trovato!
  risultato.isFornitoreNoto = true;
  risultato.fornitore = fornitore;
  
  // Verifica se skip L0 (basato su status)
  if (CONNECTOR_FORNITORI.STATUS_SKIP_L0.indexOf(fornitore.status) >= 0) {
    risultato.skipL0 = true;
  }
  
  // Genera contesto per L1
  risultato.contesto = generaContestoFornitore(fornitore);
  
  logConnectorFornitori("Trovato: " + fornitore.nome + " | Skip L0: " + risultato.skipL0);
  
  return risultato;
}

/**
 * Genera stringa contesto per arricchire prompt L1
 * @param {Object} fornitore - Oggetto fornitore
 * @returns {String} Contesto formattato
 */
function generaContestoFornitore(fornitore) {
  if (!fornitore) return "";
  
  var linee = [];
  
  linee.push("═══ FORNITORE CONOSCIUTO ═══");
  linee.push("Nome: " + fornitore.nome);
  
  // Priorità urgente
  if (fornitore.prioritaUrgente) {
    linee.push("⚠️ PRIORITA' URGENTE - Richiede attenzione speciale");
  }
  
  // Sconto attivo
  if (fornitore.scontoPercentuale && fornitore.scontoPercentuale > 0) {
    linee.push("💰 Sconto attivo: " + fornitore.scontoPercentuale + "%");
  }
  
  // Promo imminente
  if (fornitore.dataProssimaPromo) {
    var oggi = new Date();
    var dataPromo = new Date(fornitore.dataProssimaPromo);
    var giorniMancanti = Math.ceil((dataPromo - oggi) / (1000 * 60 * 60 * 24));
    
    if (giorniMancanti >= 0 && giorniMancanti <= 7) {
      linee.push("🏷️ PROMO IMMINENTE: tra " + giorniMancanti + " giorni");
    } else if (giorniMancanti >= -14 && giorniMancanti < 0) {
      linee.push("🏷️ PROMO IN CORSO (iniziata " + Math.abs(giorniMancanti) + " gg fa)");
    }
  }
  
  // Status ultima azione
  if (fornitore.statusUltimaAzione) {
    linee.push("📝 Ultimo contatto: " + fornitore.statusUltimaAzione);
  }
  
  // Performance
  if (fornitore.performanceScore) {
    var stelle = "";
    for (var i = 0; i < fornitore.performanceScore; i++) stelle += "★";
    for (var j = fornitore.performanceScore; j < 5; j++) stelle += "☆";
    linee.push("📊 Performance: " + stelle + " (" + fornitore.performanceScore + "/5)");
  }
  
  linee.push("═══════════════════════════════");
  
  return linee.join("\n");
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER INTERNI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica dati fornitori da tab locale (con cache)
 * @returns {Object|null} {headers, rows, colMap}
 */
function caricaDatiFornitoriLocali() {
  // Check cache
  var cache = CONNECTOR_FORNITORI._cache;
  var now = new Date().getTime();
  
  if (cache.fornitori && cache.timestamp) {
    var age = (now - cache.timestamp) / 1000;
    if (age < CONNECTOR_FORNITORI.DEFAULTS.CACHE_TTL_SEC) {
      return cache.fornitori;
    }
  }
  
  // Carica da foglio
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return null;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    // Costruisci mappa colonne
    var colMap = {};
    Object.keys(CONNECTOR_FORNITORI.COLONNE).forEach(function(key) {
      var nomeColonna = CONNECTOR_FORNITORI.COLONNE[key];
      var idx = headers.indexOf(nomeColonna);
      if (idx >= 0) {
        colMap[key] = idx;
      }
    });
    
    // Salva in cache
    var result = {
      headers: headers,
      rows: rows,
      colMap: colMap
    };
    
    CONNECTOR_FORNITORI._cache.fornitori = result;
    CONNECTOR_FORNITORI._cache.timestamp = now;
    
    return result;
    
  } catch(e) {
    logConnectorFornitori("Errore caricamento dati locali: " + e.message);
    return null;
  }
}

/**
 * Costruisce oggetto fornitore da riga
 */
function costruisciOggettoFornitore(row, colMap, rowNum) {
  return {
    // ID e anagrafica
    id: row[colMap.ID_FORNITORE] || "",
    nome: row[colMap.NOME_AZIENDA] || "",
    email: row[colMap.EMAIL_ORDINI] || "",
    emailAltri: row[colMap.EMAIL_ALTRI] || "",
    contatto: row[colMap.CONTATTO] || "",
    telefono: row[colMap.TELEFONO] || "",
    
    // Status
    status: row[colMap.STATUS_FORNITORE] || "",
    prioritaUrgente: row[colMap.PRIORITA_URGENTE] === "X" || row[colMap.PRIORITA_URGENTE] === true,
    
    // Commerciale
    scontoPercentuale: parseFloat(row[colMap.SCONTO_PERCENTUALE]) || 0,
    dataProssimaPromo: row[colMap.DATA_PROSSIMA_PROMO] || null,
    
    // Tracking
    statusUltimaAzione: row[colMap.STATUS_ULTIMA_AZIONE] || "",
    performanceScore: parseInt(row[colMap.PERFORMANCE_SCORE]) || 0,
    dataUltimaEmail: row[colMap.DATA_ULTIMA_EMAIL] || null,
    
    // Meta
    rowNum: rowNum,
    _source: "FORNITORI_SYNC"
  };
}

/**
 * Estrae dominio da email
 */
function estraiDominio(email) {
  if (!email) return "";
  email = email.toLowerCase().trim();
  var at = email.indexOf("@");
  return at > 0 ? email.substring(at + 1) : "";
}

/**
 * Verifica esistenza tab FORNITORI_SYNC
 */
function esisteTabFornitoriSync() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    return !!(sheet && sheet.getLastRow() > 1);
  } catch(e) {
    return false;
  }
}

/**
 * Conta fornitori nel sync locale
 */
function contaFornitoriLocali() {
  var dati = caricaDatiFornitoriLocali();
  return dati ? dati.rows.length : 0;
}