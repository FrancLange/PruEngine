function testProcessaDiretto() {
  // Invalida cache
  CONNECTOR_EMAIL._cache = { emails: null, timestamp: null };
  
  var emails = caricaEmailDaProcessare();
  Logger.log("Email caricate: " + emails.length);
  
  if (emails.length > 0) {
    Logger.log("Prima email: " + JSON.stringify(emails[0]));
    
    // Prova a creare questione
    var questione = creaQuestioneDaEmail(emails[0]);
    Logger.log("Questione creata: " + JSON.stringify(questione));
  }
}