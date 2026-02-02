/**
 * ==========================================================================================
 * WRITERS - MOTORE FORNITORI v1.0.0
 * ==========================================================================================
 * Funzioni per scrivere/aggiornare dati fornitori
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE FORNITORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea nuovo fornitore
 * @param {String} nome - Nome azienda
 * @param {String} email - Email ordini
 * @param {Object} datiExtra - Dati opzionali aggiuntivi
 * @returns {String} ID fornitore creato
 */
function creaFornitore(nome, email, datiExtra) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet) {
    throw new Error("Foglio FORNITORI non trovato");
  }
  
  datiExtra = datiExtra || {};
  
  var id = generaNuovoIdFornitore();
  var now = new Date();
  
  // Costruisci riga con defaults
  var nuovaRiga = [
    id,                                          // ID_FORNITORE
    datiExtra.prioritaUrgente ? "X" : "",        // PRIORITA_URGENTE
    nome,                                         // NOME_AZIENDA
    email,                                        // EMAIL_ORDINI
    datiExtra.emailAltri || "",                  // EMAIL_ALTRI
    datiExtra.contatto || "",                    // CONTATTO
    datiExtra.telefono || "",                    // TELEFONO
    datiExtra.partitaIva || "",                  // PARTITA_IVA
    datiExtra.indirizzo || "",                   // INDIRIZZO
    datiExtra.note || "",                        // NOTE
    datiExtra.metodoInvio || "",                 // METODO_INVIO
    datiExtra.tipoListino || "",                 // TIPO_LISTINO
    datiExtra.linkListino || "",                 // LINK_LISTINO
    datiExtra.minOrdine || CONFIG_FORNITORI.DEFAULTS.MIN_ORDINE, // MIN_ORDINE
    datiExtra.leadTime || CONFIG_FORNITORI.DEFAULTS.LEAD_TIME_GIORNI, // LEAD_TIME_GIORNI
    datiExtra.promoSuRichiesta || "",            // PROMO_SU_RICHIESTA
    datiExtra.dataProssimaPromo || "",           // DATA_PROSSIMA_PROMO
    datiExtra.promoAnnualiNum || "",             // PROMO_ANNUALI_NUM
    datiExtra.promoAnnualiRichieste || "",       // PROMO_ANNUALI_RICHIESTE
    datiExtra.scontoTarget || CONFIG_FORNITORI.DEFAULTS.SCONTO_TARGET, // SCONTO_TARGET
    "",                                           // RICHIESTA_SCONTO
    CONFIG_FORNITORI.ESITO_SCONTO.NON_RICHIESTO, // ESITO_RICHIESTA_SCONTO
    "",                                           // TIPO_SCONTO_CONCESSO
    datiExtra.scontoPercentuale || 0,            // SCONTO_PERCENTUALE
    "",                                           // DATA_ULTIMO_ORDINE
    "",                                           // DATA_ULTIMA_EMAIL
    "",                                           // DATA_ULTIMA_ANALISI
    "",                                           // STATUS_ULTIMA_AZIONE
    0,                                            // EMAIL_ANALIZZATE_COUNT
    CONFIG_FORNITORI.DEFAULTS.PERFORMANCE_SCORE, // PERFORMANCE_SCORE
    CONFIG_FORNITORI.DEFAULTS.STATUS_INIZIALE,   // STATUS_FORNITORE
    now                                           // DATA_CREAZIONE
  ];
  
  sheet.appendRow(nuovaRiga);
  
  // Registra azione
  registraAzione(id, nome, CONFIG_FORNITORI.TIPO_AZIONE.CREAZIONE, "Nuovo fornitore creato");
  
  logSistemaFornitori("✅ Creato fornitore: " + id + " - " + nome);
  
  return id;
}

// ═══════════════════════════════════════════════════════════════════════
// AGGIORNAMENTO FORNITORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna campi specifici di un fornitore
 * @param {String} idFornitore - ID fornitore
 * @param {Object} updates - {campo: valore}
 * @returns {Boolean} Successo
 */
function aggiornaFornitore(idFornitore, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet) {
    logSistemaFornitori("❌ Foglio FORNITORI non trovato");
    return false;
  }
  
  var rowNum = getRowNumFornitore(idFornitore);
  if (!rowNum) {
    logSistemaFornitori("❌ Fornitore non trovato: " + idFornitore);
    return false;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var campiAggiornati = [];
  
  Object.keys(updates).forEach(function(campo) {
    var valore = updates[campo];
    
    // Trova nome colonna da CONFIG
    var nomeColonna = CONFIG_FORNITORI.COLONNE_FORNITORI[campo] || campo;
    var colIndex = headers.indexOf(nomeColonna);
    
    if (colIndex >= 0) {
      sheet.getRange(rowNum, colIndex + 1).setValue(valore);
      campiAggiornati.push(campo);
    } else {
      logSistemaFornitori("⚠️ Colonna non trovata: " + nomeColonna);
    }
  });
  
  if (campiAggiornati.length > 0) {
    logSistemaFornitori("✅ Aggiornato " + idFornitore + ": " + campiAggiornati.join(", "));
  }
  
  return campiAggiornati.length > 0;
}

/**
 * Aggiorna data ultima email ricevuta
 * Chiamato da Email Engine dopo analisi
 * @param {String} idFornitore - ID fornitore
 * @param {String} statusAzione - Descrizione azione
 */
function aggiornaUltimaEmail(idFornitore, statusAzione) {
  var now = new Date();
  
  // Incrementa contatore email
  var fornitore = getFornitoreById(idFornitore);
  var count = (fornitore && fornitore.emailAnalizzateCount) ? fornitore.emailAnalizzateCount + 1 : 1;
  
  aggiornaFornitore(idFornitore, {
    DATA_ULTIMA_EMAIL: now,
    DATA_ULTIMA_ANALISI: now,
    STATUS_ULTIMA_AZIONE: statusAzione || "Email analizzata",
    EMAIL_ANALIZZATE_COUNT: count
  });
}

/**
 * Aggiorna data ultimo ordine
 * @param {String} idFornitore - ID fornitore
 * @param {Date} dataOrdine - Data ordine (default now)
 */
function aggiornaUltimoOrdine(idFornitore, dataOrdine) {
  aggiornaFornitore(idFornitore, {
    DATA_ULTIMO_ORDINE: dataOrdine || new Date(),
    STATUS_ULTIMA_AZIONE: "Ordine effettuato"
  });
  
  // Registra storico
  var fornitore = getFornitoreById(idFornitore);
  registraAzione(idFornitore, fornitore ? fornitore.nome : idFornitore, 
    CONFIG_FORNITORI.TIPO_AZIONE.ORDINE, "Ordine registrato");
}

/**
 * Aggiorna performance score
 * @param {String} idFornitore - ID fornitore
 * @param {Number} score - Score 1-5
 */
function aggiornaPerformance(idFornitore, score) {
  if (score < 1) score = 1;
  if (score > 5) score = 5;
  
  aggiornaFornitore(idFornitore, {
    PERFORMANCE_SCORE: score
  });
}

/**
 * Aggiorna status fornitore
 * @param {String} idFornitore - ID fornitore
 * @param {String} nuovoStatus - Nuovo status
 */
function aggiornaStatusFornitore(idFornitore, nuovoStatus) {
  // Valida status
  var statusValidi = Object.values(CONFIG_FORNITORI.STATUS_FORNITORE);
  if (statusValidi.indexOf(nuovoStatus) < 0) {
    logSistemaFornitori("⚠️ Status non valido: " + nuovoStatus);
    return false;
  }
  
  aggiornaFornitore(idFornitore, {
    STATUS_FORNITORE: nuovoStatus
  });
  
  var fornitore = getFornitoreById(idFornitore);
  registraAzione(idFornitore, fornitore ? fornitore.nome : idFornitore,
    CONFIG_FORNITORI.TIPO_AZIONE.MODIFICA, "Status cambiato a: " + nuovoStatus);
  
  return true;
}

// ═══════════════════════════════════════════════════════════════════════
// GESTIONE SCONTI E PROMO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Registra richiesta sconto
 * @param {String} idFornitore - ID fornitore
 * @param {Number} scontoRichiesto - Percentuale richiesta
 */
function registraRichiestaSconto(idFornitore, scontoRichiesto) {
  aggiornaFornitore(idFornitore, {
    RICHIESTA_SCONTO: "X",
    SCONTO_TARGET: scontoRichiesto,
    ESITO_RICHIESTA_SCONTO: CONFIG_FORNITORI.ESITO_SCONTO.IN_ATTESA,
    STATUS_ULTIMA_AZIONE: "Richiesta sconto " + scontoRichiesto + "% inviata"
  });
  
  var fornitore = getFornitoreById(idFornitore);
  registraAzione(idFornitore, fornitore ? fornitore.nome : idFornitore,
    CONFIG_FORNITORI.TIPO_AZIONE.RICHIESTA_SCONTO, "Richiesto sconto " + scontoRichiesto + "%");
}

/**
 * Registra esito richiesta sconto
 * @param {String} idFornitore - ID fornitore
 * @param {String} esito - ACCETTATO/RIFIUTATO/CONTROPROPOSTA
 * @param {Number} scontoConcesso - Percentuale concessa (se applicabile)
 * @param {String} tipoSconto - Tipo sconto concesso
 */
function registraEsitoSconto(idFornitore, esito, scontoConcesso, tipoSconto) {
  var updates = {
    ESITO_RICHIESTA_SCONTO: esito,
    STATUS_ULTIMA_AZIONE: "Esito sconto: " + esito
  };
  
  if (scontoConcesso !== undefined && scontoConcesso > 0) {
    updates.SCONTO_PERCENTUALE = scontoConcesso;
  }
  
  if (tipoSconto) {
    updates.TIPO_SCONTO_CONCESSO = tipoSconto;
  }
  
  aggiornaFornitore(idFornitore, updates);
  
  var fornitore = getFornitoreById(idFornitore);
  var descrizione = "Esito: " + esito;
  if (scontoConcesso) descrizione += " - " + scontoConcesso + "%";
  
  registraAzione(idFornitore, fornitore ? fornitore.nome : idFornitore,
    CONFIG_FORNITORI.TIPO_AZIONE.RICHIESTA_SCONTO, descrizione);
}

/**
 * Registra promo attiva
 * @param {String} idFornitore - ID fornitore
 * @param {Date} dataPromo - Data promo
 * @param {String} descrizione - Descrizione promo
 */
function registraPromo(idFornitore, dataPromo, descrizione) {
  var fornitore = getFornitoreById(idFornitore);
  var richieste = (fornitore && fornitore.promoAnnualiRichieste) ? fornitore.promoAnnualiRichieste + 1 : 1;
  
  aggiornaFornitore(idFornitore, {
    DATA_PROSSIMA_PROMO: dataPromo,
    PROMO_ANNUALI_RICHIESTE: richieste,
    STATUS_ULTIMA_AZIONE: descrizione || "Promo pianificata"
  });
  
  registraAzione(idFornitore, fornitore ? fornitore.nome : idFornitore,
    CONFIG_FORNITORI.TIPO_AZIONE.PROMO_ATTIVATA, descrizione || "Promo pianificata per " + dataPromo);
}

// ═══════════════════════════════════════════════════════════════════════
// LETTURA FORNITORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Ottiene fornitore per ID
 * @param {String} idFornitore - ID fornitore
 * @returns {Object|null}
 */
function getFornitoreById(idFornitore) {
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
    if (data[i][colMap.ID_FORNITORE] === idFornitore) {
      var fornitore = costruisciOggettoFornitore(data[i], colMap);
      fornitore.rowNum = i + 1;
      return fornitore;
    }
  }
  
  return null;
}

// ═══════════════════════════════════════════════════════════════════════
// IMPORT BULK
// ═══════════════════════════════════════════════════════════════════════

/**
 * Importa fornitori da array
 * @param {Array} fornitori - Array di oggetti fornitore
 * @returns {Object} {importati, errori, ids}
 */
function importaFornitoriBulk(fornitori) {
  var risultato = { importati: 0, errori: 0, ids: [] };
  
  fornitori.forEach(function(f) {
    try {
      if (!f.nome || !f.email) {
        risultato.errori++;
        return;
      }
      
      var id = creaFornitore(f.nome, f.email, f);
      risultato.ids.push(id);
      risultato.importati++;
      
    } catch(e) {
      risultato.errori++;
      logSistemaFornitori("❌ Errore import: " + e.toString());
    }
  });
  
  logSistemaFornitori("📦 Import bulk: " + risultato.importati + " ok, " + risultato.errori + " errori");
  
  return risultato;
}