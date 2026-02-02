/**
 * ==========================================================================================
 * BATCH.gs - Gestione Storico e Backlog
 * ==========================================================================================
 */

/**
 * Analizza email più vecchie di X giorni che non sono ancora state completate.
 * Utile per recuperare arretrati o email saltate.
 * 
 * @param {Object} params - {giorni: 7, max_email: 20}
 */
function analizzaBacklogVecchi(params) {
  var giorni = (params && params.giorni) ? params.giorni : 7;
  var maxEmail = (params && params.max_email) ? params.max_email : 20;
  
  logSistema("🕰️ BATCH START: Analisi email più vecchie di " + giorni + " giorni");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN); // [2]
  
  if (!sheet) return { errore: "Foglio LOG_IN non trovato" };

  var data = sheet.getDataRange().getValues();
  var headers = data;
  
  // Mappatura colonne basata sulla tua configurazione [3]
  var colTimestamp = headers.indexOf(CONFIG.COLONNE_LOG_IN.TIMESTAMP);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colMittente = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colOggetto = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  var colCorpo = headers.indexOf(CONFIG.COLONNE_LOG_IN.CORPO);

  // Calcola la data limite
  var dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - giorni);
  
  var count = 0;
  var processed = 0;

  // Itera sulle righe (saltando l'header)
  for (var i = 1; i < data.length && count < maxEmail; i++) {
    var row = data[i];
    var rowNum = i + 1;
    var emailDate = new Date(row[colTimestamp]);
    var status = row[colStatus];

    // CRITERI DI FILTRO:
    // 1. Data email < Data Limite (più vecchia di una settimana)
    // 2. Status non è finale (es. è ancora DA_ANALIZZARE o DA_FILTRARE)
    // 3. Status non è già SPAM o GESTITO
    
    if (emailDate < dataLimite && 
        status !== "ANALIZZATO" && 
        status !== "SPAM" && 
        status !== "GESTITO" &&
        status !== "NEEDS_REVIEW") {
      
      try {
        logSistema("🕰️ Recupero backlog: " + row[colId]);
        
        // 1. Esegui Layer 1 (GPT) [4]
        var emailData = {
          mittente: row[colMittente],
          oggetto: row[colOggetto],
          corpo: row[colCorpo]
        };
        
        aggiornaStatus(sheet, rowNum, "IN_ANALISI"); // Funzione helper esistente
        
        var risL1 = analisiLayer1(emailData); // [5]
        if (risL1.success) {
          scriviRisultatiL1(sheet, rowNum, risL1); // [6]
        }
        
        // 2. Esegui Layer 2 (Claude o Fallback) se L1 ok [7]
        if (risL1.success) {
           // Prepara dati per L2
           var l1DataForL2 = {
             tags: risL1.tags,
             sintesi: risL1.sintesi,
             scores: risL1.scores
           };
           
           var risL2 = analisiLayer2(emailData, l1DataForL2); // [7]
           if (risL2.success) {
             scriviRisultatiL2(sheet, rowNum, risL2); // [8]
             
             // 3. Esegui Layer 3 (Merge) [9]
             var risL3 = mergeAnalisi(risL1, risL2);
             if (risL3.success) {
               scriviRisultatiL3(sheet, rowNum, risL3); // [10]
               
               var statusFinale = risL3.needsReview ? "NEEDS_REVIEW" : "ANALIZZATO";
               aggiornaStatus(sheet, rowNum, statusFinale);
             }
           }
        }
        
        count++;
        processed++;
        
      } catch (e) {
        logSistema("❌ Errore BATCH su " + row[colId] + ": " + e.message);
        scriviErroreL1(sheet, rowNum, e.toString()); // [11]
      }
    }
  }
  
  return { processate: processed, giorni_backlog: giorni };
}