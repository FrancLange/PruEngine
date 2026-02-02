/**
 * ==========================================================================================
 * LOGIC - MOTORE FORNITORI v1.0.0
 * ==========================================================================================
 * Logica principale e integrazione con Email Engine
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// INTEGRAZIONE EMAIL ENGINE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Pre-check email prima di L0 Spam Filter
 * Chiamato da Email Engine prima di analizzare
 * 
 * @param {String} emailMittente - Email mittente
 * @returns {Object} {isFornitoreNoto, skipL0, contesto, fornitore}
 */
function preCheckEmailFornitore(emailMittente) {
  var risultato = {
    isFornitoreNoto: false,
    skipL0: false,
    contesto: "",
    fornitore: null
  };
  
  if (!emailMittente) return risultato;
  
  // Lookup fornitore
  var fornitore = lookupFornitoreByEmail(emailMittente);
  
  if (fornitore) {
    risultato.isFornitoreNoto = true;
    risultato.fornitore = fornitore;
    
    // Skip L0 solo se fornitore ATTIVO o IN_VALUTAZIONE
    if (fornitore.status === "ATTIVO" || fornitore.status === "IN_VALUTAZIONE" || fornitore.status === "NUOVO") {
      risultato.skipL0 = true;
    }
    
    // Genera contesto per L1
    risultato.contesto = generaContestoFornitore(fornitore);
    
    logSistemaFornitori("📧 Fornitore riconosciuto: " + fornitore.nome + " (" + emailMittente + ")");
  }
  
  return risultato;
}

/**
 * Post-analisi email: aggiorna dati fornitore
 * Chiamato da Email Engine dopo analisi completata
 * 
 * @param {String} emailMittente - Email mittente
 * @param {Object} risultatoAnalisi - Risultato da Email Engine {tags, scores, sintesi}
 * @param {String} idEmail - ID email analizzata
 */
function postAnalisiEmailFornitore(emailMittente, risultatoAnalisi, idEmail) {
  var fornitore = lookupFornitoreByEmail(emailMittente);
  
  if (!fornitore) return;
  
  // Aggiorna ultima email
  var statusAzione = "";
  
  // Determina status in base ai tags/scores
  if (risultatoAnalisi.tags) {
    var tags = risultatoAnalisi.tags;
    
    if (tags.indexOf("PROBLEMA") >= 0 || (risultatoAnalisi.scores && risultatoAnalisi.scores.problema > 70)) {
      statusAzione = "⚠️ Segnalato problema";
    } else if (tags.indexOf("ORDINE") >= 0 || tags.indexOf("CONFERMA") >= 0) {
      statusAzione = "✅ Conferma/Ordine ricevuto";
    } else if (tags.indexOf("PROMO") >= 0) {
      statusAzione = "🏷️ Comunicazione promo";
    } else if (tags.indexOf("LISTINO") >= 0 || tags.indexOf("CATALOGO") >= 0) {
      statusAzione = "📋 Listino/Catalogo ricevuto";
    } else if (tags.indexOf("FATTURA") >= 0) {
      statusAzione = "📄 Fattura ricevuta";
    } else {
      statusAzione = "📧 Email analizzata";
    }
  } else {
    statusAzione = "📧 Email analizzata";
  }
  
  aggiornaUltimaEmail(fornitore.id, statusAzione);
  
  // Registra in storico
  registraAzione(
    fornitore.id, 
    fornitore.nome, 
    CONFIG_FORNITORI.TIPO_AZIONE.EMAIL_RICEVUTA,
    statusAzione + " - " + (risultatoAnalisi.sintesi || "").substring(0, 100),
    idEmail
  );
}

// ═══════════════════════════════════════════════════════════════════════
// ANALISI FORNITORI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Analizza tutti i fornitori e genera alert
 * @returns {Object} {totali, alertPromo, alertOrdini, problemi}
 */
function analizzaFornitori() {
  logSistemaFornitori("🔍 Analisi fornitori START");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { totali: 0, alertPromo: [], alertOrdini: [], problemi: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  var risultato = {
    totali: 0,
    alertPromo: [],
    alertOrdini: [],
    problemi: []
  };
  
  var oggi = new Date();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (!row[colMap.ID_FORNITORE]) continue;
    if (row[colMap.STATUS_FORNITORE] !== "ATTIVO") continue;
    
    risultato.totali++;
    
    var fornitore = {
      id: row[colMap.ID_FORNITORE],
      nome: row[colMap.NOME_AZIENDA],
      priorita: row[colMap.PRIORITA_URGENTE] === "X",
      dataPromo: row[colMap.DATA_PROSSIMA_PROMO],
      dataUltimoOrdine: row[colMap.DATA_ULTIMO_ORDINE],
      performance: row[colMap.PERFORMANCE_SCORE]
    };
    
    // Check promo imminente
    if (fornitore.dataPromo) {
      var dataPromo = new Date(fornitore.dataPromo);
      var giorniMancanti = Math.ceil((dataPromo - oggi) / (1000 * 60 * 60 * 24));
      
      if (giorniMancanti >= 0 && giorniMancanti <= 7) {
        risultato.alertPromo.push({
          fornitore: fornitore.nome,
          id: fornitore.id,
          giorni: giorniMancanti,
          data: dataPromo
        });
      }
    }
    
    // Check ordine scaduto (es. non ordini da 60+ giorni per fornitori attivi)
    if (fornitore.dataUltimoOrdine) {
      var dataOrdine = new Date(fornitore.dataUltimoOrdine);
      var giorniDaOrdine = Math.floor((oggi - dataOrdine) / (1000 * 60 * 60 * 24));
      
      if (giorniDaOrdine > 60) {
        risultato.alertOrdini.push({
          fornitore: fornitore.nome,
          id: fornitore.id,
          giorni: giorniDaOrdine
        });
      }
    }
    
    // Check performance bassa
    if (fornitore.performance && fornitore.performance < 2) {
      risultato.problemi.push({
        fornitore: fornitore.nome,
        id: fornitore.id,
        motivo: "Performance bassa: " + fornitore.performance + "/5"
      });
    }
  }
  
  logSistemaFornitori("🔍 Analisi completata: " + risultato.totali + " fornitori, " + 
    risultato.alertPromo.length + " promo, " + 
    risultato.alertOrdini.length + " ordini scaduti");
  
  return risultato;
}

/**
 * Genera report giornaliero fornitori
 * @returns {String}
 */
function generaReportGiornaliero() {
  var analisi = analizzaFornitori();
  
  var report = "📊 REPORT GIORNALIERO FORNITORI\n";
  report += "Data: " + new Date().toLocaleDateString('it-IT') + "\n\n";
  
  report += "📈 STATISTICHE\n";
  report += "Fornitori attivi: " + analisi.totali + "\n\n";
  
  if (analisi.alertPromo.length > 0) {
    report += "🏷️ PROMO IMMINENTI (" + analisi.alertPromo.length + ")\n";
    analisi.alertPromo.forEach(function(alert) {
      report += "• " + alert.fornitore + " - tra " + alert.giorni + " giorni\n";
    });
    report += "\n";
  }
  
  if (analisi.alertOrdini.length > 0) {
    report += "⏰ DA RIORDINARE (" + analisi.alertOrdini.length + ")\n";
    analisi.alertOrdini.forEach(function(alert) {
      report += "• " + alert.fornitore + " - ultimo ordine " + alert.giorni + " giorni fa\n";
    });
    report += "\n";
  }
  
  if (analisi.problemi.length > 0) {
    report += "⚠️ PROBLEMI (" + analisi.problemi.length + ")\n";
    analisi.problemi.forEach(function(p) {
      report += "• " + p.fornitore + " - " + p.motivo + "\n";
    });
  }
  
  return report;
}

// ═══════════════════════════════════════════════════════════════════════
// AZIONI AUTOMATICHE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa azioni suggerite da Email Engine
 * @param {String} idFornitore - ID fornitore
 * @param {Array} azioni - Array di azioni suggerite
 * @param {String} idEmail - ID email origine
 */
function processaAzioniSuggerite(idFornitore, azioni, idEmail) {
  if (!azioni || azioni.length === 0) return;
  
  var fornitore = getFornitoreById(idFornitore);
  if (!fornitore) return;
  
  azioni.forEach(function(azione) {
    switch(azione) {
      case "FLAG_FATTURA":
        registraAzione(idFornitore, fornitore.nome, "EMAIL_RICEVUTA", "Fattura ricevuta", idEmail);
        break;
        
      case "FLAG_ORDINE":
        // Potenziale aggiornamento ultimo ordine
        registraAzione(idFornitore, fornitore.nome, "EMAIL_RICEVUTA", "Conferma ordine ricevuta", idEmail);
        break;
        
      case "FLAG_CATALOGO":
      case "FLAG_LISTINO":
        registraAzione(idFornitore, fornitore.nome, "EMAIL_RICEVUTA", "Listino/Catalogo ricevuto", idEmail);
        break;
        
      case "FLAG_PROMO":
        registraAzione(idFornitore, fornitore.nome, "PROMO_ATTIVATA", "Comunicazione promo ricevuta", idEmail);
        break;
        
      case "CREA_QUESTIONE_PROBLEMA":
        registraAzione(idFornitore, fornitore.nome, "PROBLEMA", "Problema segnalato da email", idEmail);
        // TODO: Creare questione in Motore Questioni
        break;
        
      case "ALERT_URGENTE":
        if (!fornitore.prioritaUrgente) {
          aggiornaFornitore(idFornitore, { PRIORITA_URGENTE: "X" });
        }
        registraAzione(idFornitore, fornitore.nome, "EMAIL_RICEVUTA", "⚠️ Email urgente", idEmail);
        break;
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// SYNC CON IL SEGRETARIO (se presente)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Importa fornitori da Il Segretario (se configurato)
 * @param {String} sheetId - ID foglio Il Segretario
 * @returns {Object} Risultato import
 */
function importaDaIlSegretario(sheetId) {
  if (!sheetId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setup = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.SETUP);
    // Leggi ID da setup se presente
    // Per ora placeholder
    logSistemaFornitori("⚠️ Import da Il Segretario: ID non configurato");
    return { importati: 0, errori: 0 };
  }
  
  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName("FORNITORI");
    
    if (!sourceSheet) {
      logSistemaFornitori("❌ Foglio FORNITORI non trovato in Il Segretario");
      return { importati: 0, errori: 0 };
    }
    
    var data = sourceSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Mappa colonne Il Segretario (basandoti sulla struttura che hai descritto)
    var colNome = headers.indexOf("Nome Azienda");
    var colEmail = headers.indexOf("Email Ordini");
    var colContatto = headers.indexOf("Contatto");
    var colMetodo = headers.indexOf("Metodo Invio (Sito/Modulo/Lista)");
    var colTipoListino = headers.indexOf("Tipo Listino (Sito/Excel/No)");
    var colPromoRichiesta = headers.indexOf("Promo su richiesta");
    var colDataPromo = headers.indexOf("Data Prox Promo");
    var colScontoTarget = headers.indexOf("Sconto Target");
    var colPriorita = headers.indexOf("Priorità Urgente (X)");
    
    var fornitori = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (!row[colNome] || !row[colEmail]) continue;
      
      fornitori.push({
        nome: row[colNome],
        email: row[colEmail],
        contatto: row[colContatto] || "",
        metodoInvio: row[colMetodo] || "",
        tipoListino: row[colTipoListino] || "",
        promoSuRichiesta: row[colPromoRichiesta] || "",
        dataProssimaPromo: row[colDataPromo] || "",
        scontoTarget: row[colScontoTarget] || "",
        prioritaUrgente: row[colPriorita] === "X"
      });
    }
    
    var risultato = importaFornitoriBulk(fornitori);
    
    logSistemaFornitori("📦 Import da Il Segretario: " + risultato.importati + " fornitori");
    
    return risultato;
    
  } catch(e) {
    logSistemaFornitori("❌ Errore import Il Segretario: " + e.toString());
    return { importati: 0, errori: 1, errore: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════
// MANUTENZIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Pulizia fornitori inattivi
 * @param {Number} giorniInattivita - Giorni senza email/ordine
 * @param {Boolean} dryRun - Se true, solo report senza modifiche
 * @returns {Object}
 */
function pulisciFornitori(giorniInattivita, dryRun) {
  giorniInattivita = giorniInattivita || 180;
  dryRun = dryRun !== false;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { candidati: 0, sospesi: 0 };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colMap = {};
  Object.keys(CONFIG_FORNITORI.COLONNE_FORNITORI).forEach(function(key) {
    var nome = CONFIG_FORNITORI.COLONNE_FORNITORI[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  var oggi = new Date();
  var candidati = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (!row[colMap.ID_FORNITORE]) continue;
    if (row[colMap.STATUS_FORNITORE] !== "ATTIVO") continue;
    
    var dataUltimaAttivita = row[colMap.DATA_ULTIMA_EMAIL] || row[colMap.DATA_ULTIMO_ORDINE] || row[colMap.DATA_CREAZIONE];
    
    if (dataUltimaAttivita) {
      var giorni = Math.floor((oggi - new Date(dataUltimaAttivita)) / (1000 * 60 * 60 * 24));
      
      if (giorni > giorniInattivita) {
        candidati.push({
          id: row[colMap.ID_FORNITORE],
          nome: row[colMap.NOME_AZIENDA],
          giorni: giorni,
          rowNum: i + 1
        });
      }
    }
  }
  
  var sospesi = 0;
  
  if (!dryRun && candidati.length > 0) {
    candidati.forEach(function(c) {
      aggiornaStatusFornitore(c.id, "SOSPESO");
      sospesi++;
    });
  }
  
  var report = "🧹 Pulizia fornitori\n" +
    "Soglia: " + giorniInattivita + " giorni\n" +
    "Candidati: " + candidati.length + "\n" +
    (dryRun ? "(DRY RUN - nessuna modifica)" : "Sospesi: " + sospesi);
  
  logSistemaFornitori(report);
  
  return { candidati: candidati.length, sospesi: sospesi, lista: candidati };
}