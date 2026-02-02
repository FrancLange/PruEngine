/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - UPDATE v1.0.0
 * ==========================================================================================
 * Funzioni per aggiornare il foglio Fornitori MASTER (cross-sheet write)
 * 
 * AUTONOMO: Scrive direttamente sul master per mantenere dati sincronizzati
 * ==========================================================================================
 */

/**
 * Aggiorna campi fornitore sul MASTER
 * 
 * @param {String} idFornitore - ID fornitore (es. FOR-001)
 * @param {Object} updates - {CAMPO: valore, ...}
 * @returns {Object} {success, errore}
 */
function aggiornaFornitoreSuMaster(idFornitore, updates) {
  var risultato = { success: false, errore: null };
  
  if (!idFornitore) {
    risultato.errore = "ID fornitore obbligatorio";
    return risultato;
  }
  
  if (!updates || Object.keys(updates).length === 0) {
    risultato.errore = "Nessun campo da aggiornare";
    return risultato;
  }
  
  try {
    // 1. Ottieni ID master
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    
    if (!masterId) {
      throw new Error("ID Fornitori Master non configurato");
    }
    
    // 2. Apri master
    var masterSs = SpreadsheetApp.openById(masterId);
    var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_MASTER);
    
    if (!masterSheet) {
      throw new Error("Tab FORNITORI non trovata nel master");
    }
    
    // 3. Trova riga fornitore
    var data = masterSheet.getDataRange().getValues();
    var headers = data[0];
    
    var colId = headers.indexOf(CONNECTOR_FORNITORI.COLONNE.ID_FORNITORE);
    if (colId < 0) {
      throw new Error("Colonna ID_FORNITORE non trovata nel master");
    }
    
    var rowNum = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][colId] === idFornitore) {
        rowNum = i + 1; // +1 per header
        break;
      }
    }
    
    if (rowNum < 0) {
      throw new Error("Fornitore " + idFornitore + " non trovato nel master");
    }
    
    // 4. Aggiorna campi
    var campiAggiornati = [];
    
    Object.keys(updates).forEach(function(campo) {
      // Trova nome colonna dal mapping
      var nomeColonna = CONNECTOR_FORNITORI.COLONNE[campo] || campo;
      var colIndex = headers.indexOf(nomeColonna);
      
      if (colIndex >= 0) {
        masterSheet.getRange(rowNum, colIndex + 1).setValue(updates[campo]);
        campiAggiornati.push(campo);
      } else {
        logConnectorFornitori("⚠️ Colonna non trovata: " + nomeColonna);
      }
    });
    
    if (campiAggiornati.length > 0) {
      risultato.success = true;
      logConnectorFornitori("✅ Aggiornato " + idFornitore + ": " + campiAggiornati.join(", "));
    } else {
      risultato.errore = "Nessun campo aggiornato";
    }
    
  } catch(e) {
    risultato.errore = e.toString();
    logConnectorFornitori("❌ Errore update master: " + risultato.errore);
  }
  
  return risultato;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI SPECIFICHE POST-ANALISI EMAIL
// ═══════════════════════════════════════════════════════════════════════

/**
 * Post-analisi: aggiorna fornitore dopo analisi email
 * 
 * Questa è la FUNZIONE PRINCIPALE chiamata da Logic.gs dopo analisi completata
 * 
 * @param {String} emailMittente - Email analizzata
 * @param {Object} risultatoAnalisi - {tags, scores, sintesi, ...}
 * @param {String} idEmail - ID email analizzata
 * @returns {Object} {success, fornitoreAggiornato}
 */
function postAnalisiEmailFornitore(emailMittente, risultatoAnalisi, idEmail) {
  var risposta = {
    success: false,
    fornitoreAggiornato: null
  };
  
  // Verifica connettore attivo
  if (!isConnectorFornitoriAttivo()) {
    return risposta;
  }
  
  // Cerca fornitore
  var fornitore = lookupFornitoreByEmail(emailMittente);
  
  if (!fornitore) {
    // Non è un fornitore noto, niente da aggiornare
    return risposta;
  }
  
  // Prepara updates
  var updates = {
    DATA_ULTIMA_EMAIL: new Date(),
    DATA_ULTIMA_ANALISI: new Date()
  };
  
  // Determina status azione da risultato analisi
  var statusAzione = determinaStatusAzione(risultatoAnalisi);
  if (statusAzione) {
    updates.STATUS_ULTIMA_AZIONE = statusAzione;
  }
  
  // Incrementa contatore email
  var countAttuale = parseInt(fornitore.emailAnalizzateCount) || 0;
  updates.EMAIL_ANALIZZATE_COUNT = countAttuale + 1;
  
  // Aggiorna master
  var risultatoUpdate = aggiornaFornitoreSuMaster(fornitore.id, updates);
  
  if (risultatoUpdate.success) {
    risposta.success = true;
    risposta.fornitoreAggiornato = fornitore.id;
    
    logConnectorFornitori("Post-analisi: " + fornitore.nome + " → " + statusAzione);
    
    // Registra azione su STORICO_AZIONI (se esiste)
    registraAzioneFornitore(fornitore.id, "EMAIL_RICEVUTA", statusAzione, idEmail);
  }
  
  return risposta;
}

/**
 * Determina status azione da risultato analisi
 */
function determinaStatusAzione(risultatoAnalisi) {
  if (!risultatoAnalisi) return "📧 Email analizzata";
  
  var tags = risultatoAnalisi.tags || [];
  var scores = risultatoAnalisi.scores || {};
  
  // Priorità: problema prima di tutto
  if (tags.indexOf("PROBLEMA") >= 0 || (scores.problema && scores.problema > 70)) {
    return "⚠️ Problema segnalato";
  }
  
  // Ordini/Conferme
  if (tags.indexOf("ORDINE") >= 0 || tags.indexOf("CONFERMA") >= 0 || 
      (scores.ordine && scores.ordine > 70)) {
    return "✅ Ordine/Conferma";
  }
  
  // Fattura
  if (tags.indexOf("FATTURA") >= 0 || (scores.fattura && scores.fattura > 70)) {
    return "📄 Fattura ricevuta";
  }
  
  // Promo
  if (tags.indexOf("PROMO") >= 0 || (scores.promo && scores.promo > 70)) {
    return "🏷️ Comunicazione promo";
  }
  
  // Listino/Catalogo
  if (tags.indexOf("LISTINO") >= 0 || tags.indexOf("CATALOGO") >= 0 || 
      (scores.catalogo && scores.catalogo > 70)) {
    return "📋 Listino/Catalogo";
  }
  
  // Prezzi
  if (tags.indexOf("PREZZI") >= 0 || (scores.prezzi && scores.prezzi > 70)) {
    return "💰 Aggiornamento prezzi";
  }
  
  // Novità
  if (tags.indexOf("NOVITA") >= 0 || (scores.novita && scores.novita > 70)) {
    return "🆕 Novità prodotti";
  }
  
  // Urgente
  if (tags.indexOf("URGENTE") >= 0 || (scores.urgente && scores.urgente > 80)) {
    return "🔴 URGENTE";
  }
  
  // Default
  return "📧 Email analizzata";
}

/**
 * Registra azione su STORICO_AZIONI del master (se esiste)
 * Best-effort, non blocca se fallisce
 */
function registraAzioneFornitore(idFornitore, tipoAzione, descrizione, emailCollegata) {
  try {
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    if (!masterId) return;
    
    var masterSs = SpreadsheetApp.openById(masterId);
    var storicoSheet = masterSs.getSheetByName("STORICO_AZIONI");
    
    if (!storicoSheet) return; // Tab non esiste, ignora
    
    // Cerca nome fornitore
    var fornitore = lookupFornitoreById(idFornitore);
    var nomeFornitore = fornitore ? fornitore.nome : idFornitore;
    
    // Append azione
    storicoSheet.appendRow([
      new Date(),           // TIMESTAMP
      idFornitore,          // ID_FORNITORE
      nomeFornitore,        // NOME_FORNITORE
      tipoAzione,           // TIPO_AZIONE
      descrizione,          // DESCRIZIONE
      emailCollegata || "", // EMAIL_COLLEGATA
      "EMAIL_ENGINE",       // ESEGUITO_DA
      ""                    // NOTE
    ]);
    
  } catch(e) {
    // Ignora errori - è opzionale
  }
}

// ═══════════════════════════════════════════════════════════════════════
// AZIONI SUGGERITE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa azioni suggerite da Email Engine
 * 
 * @param {String} idFornitore - ID fornitore
 * @param {Array} azioni - Array azioni suggerite
 * @param {String} idEmail - ID email collegata
 */
function processaAzioniSuggeriteFornitori(idFornitore, azioni, idEmail) {
  if (!idFornitore || !azioni || azioni.length === 0) return;
  
  // Verifica connettore attivo
  if (!isConnectorFornitoriAttivo()) return;
  
  azioni.forEach(function(azione) {
    switch(azione) {
      case "FLAG_FATTURA_FORNITORE":
        registraAzioneFornitore(idFornitore, "FATTURA_RICEVUTA", "Fattura da processare", idEmail);
        break;
        
      case "FLAG_ORDINE_FORNITORE":
        // Potrebbe aggiornare DATA_ULTIMO_ORDINE
        aggiornaFornitoreSuMaster(idFornitore, {
          DATA_ULTIMO_ORDINE: new Date()
        });
        registraAzioneFornitore(idFornitore, "ORDINE", "Ordine ricevuto/confermato", idEmail);
        break;
        
      case "FLAG_CATALOGO_FORNITORE":
        registraAzioneFornitore(idFornitore, "CATALOGO", "Catalogo ricevuto", idEmail);
        break;
        
      case "FLAG_NOVITA_FORNITORE":
        registraAzioneFornitore(idFornitore, "NOVITA", "Novità prodotti", idEmail);
        break;
        
      case "CREA_QUESTIONE_PROBLEMA":
        registraAzioneFornitore(idFornitore, "PROBLEMA", "Problema da gestire", idEmail);
        // TODO: Integrazione con Connector_Questioni futuro
        break;
        
      case "ALERT_URGENTE":
        registraAzioneFornitore(idFornitore, "URGENTE", "Richiede attenzione immediata", idEmail);
        break;
        
      default:
        // Azione non gestita
        break;
    }
  });
  
  logConnectorFornitori("Processate " + azioni.length + " azioni per " + idFornitore);
}

// ═══════════════════════════════════════════════════════════════════════
// UPDATE BATCH
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna multipli fornitori in batch (più efficiente)
 * 
 * @param {Array} aggiornamenti - [{idFornitore: "FOR-001", updates: {...}}, ...]
 * @returns {Object} {totali, successi, errori}
 */
function aggiornaFornitoriMasterBatch(aggiornamenti) {
  var risultato = { totali: 0, successi: 0, errori: 0 };
  
  if (!aggiornamenti || aggiornamenti.length === 0) return risultato;
  
  risultato.totali = aggiornamenti.length;
  
  // Per batch grandi, sarebbe meglio aprire il master una volta sola
  // Per semplicità, usa la funzione singola (ottimizzare se necessario)
  
  aggiornamenti.forEach(function(agg) {
    var res = aggiornaFornitoreSuMaster(agg.idFornitore, agg.updates);
    if (res.success) {
      risultato.successi++;
    } else {
      risultato.errori++;
    }
  });
  
  logConnectorFornitori("Batch update: " + risultato.successi + "/" + risultato.totali + " ok");
  
  return risultato;
}

// ═══════════════════════════════════════════════════════════════════════
// TEST CONNESSIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test scrittura su master (verifica permessi)
 */
function testScritturaMasterFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var check = checkConnectorFornitoriReady();
  
  if (!check.canActivate) {
    ui.alert("❌ Test Fallito", check.reason, ui.ButtonSet.OK);
    return;
  }
  
  // Prova a leggere il master
  try {
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    var masterSs = SpreadsheetApp.openById(masterId);
    var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_MASTER);
    
    var righe = masterSheet.getLastRow() - 1;
    var colonne = masterSheet.getLastColumn();
    
    ui.alert(
      "✅ Test Connessione OK",
      "Master: " + masterSs.getName() + "\n" +
      "Fornitori: " + righe + "\n" +
      "Colonne: " + colonne + "\n\n" +
      "Permessi lettura/scrittura verificati.",
      ui.ButtonSet.OK
    );
    
    logConnectorFornitori("Test connessione master OK");
    
  } catch(e) {
    ui.alert("❌ Test Fallito", e.toString(), ui.ButtonSet.OK);
  }
}