/**
 * ==========================================================================================
 * HELPERS - MOTORE FORNITORI v1.0.0
 * ==========================================================================================
 * Funzioni utility condivise: lookup, ricerca, generazione ID
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LOOKUP EMAIL (per integrazione Email Engine)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cerca fornitore per email mittente
 * Usato da Email Engine per skip L0 e contesto L1
 * 
 * @param {String} emailMittente - Email da cercare
 * @returns {Object|null} Dati fornitore o null
 */
function lookupFornitoreByEmail(emailMittente) {
  if (!emailMittente) return null;
  
  emailMittente = emailMittente.toLowerCase().trim();
  
  // Estrai dominio
  var dominio = "";
  var atIndex = emailMittente.indexOf("@");
  if (atIndex > 0) {
    dominio = emailMittente.substring(atIndex + 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) return null;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Mappa colonne
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  // Cerca match
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Skip righe vuote o fornitori non attivi
    if (!row[colMap.ID_FORNITORE]) continue;
    var status = row[colMap.STATUS_FORNITORE];
    if (status === "BLACKLIST") continue;
    
    // Check EMAIL_ORDINI
    var emailOrdini = (row[colMap.EMAIL_ORDINI] || "").toLowerCase();
    if (emailOrdini === emailMittente) {
      return costruisciOggettoFornitore(row, colMap);
    }
    
    // Check EMAIL_ALTRI (lista separata da virgola)
    var emailAltri = (row[colMap.EMAIL_ALTRI] || "").toLowerCase();
    if (emailAltri) {
      var listaEmail = emailAltri.split(",").map(function(e) { return e.trim(); });
      if (listaEmail.indexOf(emailMittente) >= 0) {
        return costruisciOggettoFornitore(row, colMap);
      }
    }
    
    // Match per dominio (se abilitato)
    if (CONFIG_FORNITORI.EMAIL_LOOKUP.MATCH_DOMINIO && dominio) {
      // Estrai domini dalle email del fornitore
      var dominioOrdini = estraiDominio(emailOrdini);
      if (dominioOrdini === dominio) {
        return costruisciOggettoFornitore(row, colMap);
      }
      
      // Check domini in EMAIL_ALTRI
      if (emailAltri) {
        var trovato = emailAltri.split(",").some(function(e) {
          return estraiDominio(e.trim()) === dominio;
        });
        if (trovato) {
          return costruisciOggettoFornitore(row, colMap);
        }
      }
    }
  }
  
  return null;
}

/**
 * Estrae dominio da email
 */
function estraiDominio(email) {
  if (!email) return "";
  var atIndex = email.indexOf("@");
  return atIndex > 0 ? email.substring(atIndex + 1).toLowerCase() : "";
}

/**
 * Costruisce oggetto fornitore da riga
 */
function costruisciOggettoFornitore(row, colMap) {
  return {
    id: row[colMap.ID_FORNITORE],
    prioritaUrgente: row[colMap.PRIORITA_URGENTE] === "X",
    nome: row[colMap.NOME_AZIENDA],
    email: row[colMap.EMAIL_ORDINI],
    emailAltri: row[colMap.EMAIL_ALTRI],
    contatto: row[colMap.CONTATTO],
    telefono: row[colMap.TELEFONO],
    
    metodoInvio: row[colMap.METODO_INVIO],
    tipoListino: row[colMap.TIPO_LISTINO],
    minOrdine: row[colMap.MIN_ORDINE],
    
    promoSuRichiesta: row[colMap.PROMO_SU_RICHIESTA] === "SI",
    dataProssimaPromo: row[colMap.DATA_PROSSIMA_PROMO],
    scontoTarget: row[colMap.SCONTO_TARGET],
    scontoPercentuale: row[colMap.SCONTO_PERCENTUALE],
    
    dataUltimoOrdine: row[colMap.DATA_ULTIMO_ORDINE],
    dataUltimaEmail: row[colMap.DATA_ULTIMA_EMAIL],
    statusUltimaAzione: row[colMap.STATUS_ULTIMA_AZIONE],
    performanceScore: row[colMap.PERFORMANCE_SCORE],
    status: row[colMap.STATUS_FORNITORE],
    
    rowNum: null // Può essere aggiunto se serve
  };
}

// ═══════════════════════════════════════════════════════════════════════
// GENERAZIONE CONTESTO (per Email Engine L1)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera stringa contesto fornitore per prompt L1
 * @param {Object} fornitore - Oggetto fornitore
 * @returns {String} Contesto formattato
 */
function generaContestoFornitore(fornitore) {
  if (!fornitore) return "";
  
  var righe = [];
  
  righe.push("FORNITORE CONOSCIUTO: " + fornitore.nome);
  
  if (fornitore.prioritaUrgente) {
    righe.push("⚠️ PRIORITA' URGENTE");
  }
  
  if (fornitore.scontoPercentuale > 0) {
    righe.push("Sconto attivo: " + fornitore.scontoPercentuale + "%");
  }
  
  if (fornitore.dataProssimaPromo) {
    var oggi = new Date();
    var dataPromo = new Date(fornitore.dataProssimaPromo);
    var giorniMancanti = Math.ceil((dataPromo - oggi) / (1000 * 60 * 60 * 24));
    
    if (giorniMancanti > 0 && giorniMancanti <= 30) {
      righe.push("Promo imminente tra " + giorniMancanti + " giorni");
    } else if (giorniMancanti <= 0 && giorniMancanti >= -7) {
      righe.push("PROMO IN CORSO");
    }
  }
  
  if (fornitore.statusUltimaAzione) {
    righe.push("Ultima azione: " + fornitore.statusUltimaAzione);
  }
  
  if (fornitore.performanceScore) {
    righe.push("Performance: " + fornitore.performanceScore + "/5");
  }
  
  if (fornitore.dataUltimoOrdine) {
    var dataOrdine = new Date(fornitore.dataUltimoOrdine);
    var giorniDaOrdine = Math.floor((new Date() - dataOrdine) / (1000 * 60 * 60 * 24));
    righe.push("Ultimo ordine: " + giorniDaOrdine + " giorni fa");
  }
  
  return righe.join("\n");
}

// ═══════════════════════════════════════════════════════════════════════
// RICERCA FORNITORI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cerca fornitore per nome o email (parziale)
 * @param {String} query - Stringa di ricerca
 * @returns {Object|null} Primo fornitore trovato
 */
function cercaFornitore(query) {
  if (!query) return null;
  
  query = query.toLowerCase().trim();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) return null;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var nome = (row[colMap.NOME_AZIENDA] || "").toLowerCase();
    var email = (row[colMap.EMAIL_ORDINI] || "").toLowerCase();
    var id = (row[colMap.ID_FORNITORE] || "").toLowerCase();
    
    if (nome.indexOf(query) >= 0 || email.indexOf(query) >= 0 || id === query) {
      return costruisciOggettoFornitore(row, colMap);
    }
  }
  
  return null;
}

/**
 * Cerca tutti i fornitori che matchano la query
 * @param {String} query - Stringa di ricerca
 * @returns {Array} Lista fornitori
 */
function cercaFornitoriMultipli(query) {
  if (!query) return [];
  
  query = query.toLowerCase().trim();
  var risultati = [];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (!row[colMap.ID_FORNITORE]) continue;
    
    var nome = (row[colMap.NOME_AZIENDA] || "").toLowerCase();
    var email = (row[colMap.EMAIL_ORDINI] || "").toLowerCase();
    
    if (nome.indexOf(query) >= 0 || email.indexOf(query) >= 0) {
      risultati.push(costruisciOggettoFornitore(row, colMap));
    }
  }
  
  return risultati;
}

// ═══════════════════════════════════════════════════════════════════════
// GENERAZIONE ID
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera prossimo ID fornitore (FOR-XXX)
 * @returns {String} Nuovo ID
 */
function generaNuovoIdFornitore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return "FOR-001";
  }
  
  var data = sheet.getRange("A:A").getValues();
  var maxNum = 0;
  
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    if (id && id.toString().startsWith("FOR-")) {
      var num = parseInt(id.toString().replace("FOR-", ""));
      if (num > maxNum) maxNum = num;
    }
  }
  
  var nuovoNum = maxNum + 1;
  return "FOR-" + ("000" + nuovoNum).slice(-3);
}

// ═══════════════════════════════════════════════════════════════════════
// REPORT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera report testuale fornitori
 * @returns {String}
 */
function generaReportFornitori() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return "Nessun fornitore presente";
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  var totale = 0;
  var attivi = 0;
  var urgenti = 0;
  var conPromo = 0;
  var scoreSum = 0;
  var scoreCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[colMap.ID_FORNITORE]) continue;
    
    totale++;
    
    if (row[colMap.STATUS_FORNITORE] === "ATTIVO") attivi++;
    if (row[colMap.PRIORITA_URGENTE] === "X") urgenti++;
    if (row[colMap.DATA_PROSSIMA_PROMO]) conPromo++;
    
    var score = row[colMap.PERFORMANCE_SCORE];
    if (score > 0) {
      scoreSum += score;
      scoreCount++;
    }
  }
  
  var scoreMedio = scoreCount > 0 ? (scoreSum / scoreCount).toFixed(1) : "-";
  
  return "📊 REPORT FORNITORI\n\n" +
    "Totale: " + totale + "\n" +
    "Attivi: " + attivi + "\n" +
    "Urgenti: " + urgenti + "\n" +
    "Con promo pianificata: " + conPromo + "\n" +
    "Performance media: " + scoreMedio + "/5";
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY VARIE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se un fornitore esiste
 * @param {String} idFornitore - ID da verificare
 * @returns {Boolean}
 */
function esisteFornitore(idFornitore) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) return false;
  
  var ids = sheet.getRange("A:A").getValues();
  
  for (var i = 1; i < ids.length; i++) {
    if (ids[i][0] === idFornitore) return true;
  }
  
  return false;
}

/**
 * Ottiene il numero di riga di un fornitore
 * @param {String} idFornitore - ID fornitore
 * @returns {Number|null} Numero riga (1-based) o null
 */
function getRowNumFornitore(idFornitore) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) return null;
  
  var ids = sheet.getRange("A:A").getValues();
  
  for (var i = 1; i < ids.length; i++) {
    if (ids[i][0] === idFornitore) return i + 1;
  }
  
  return null;
}