var COLONNA_CARICA_SOCIALE_ = 13;
var CARICHE_ISTITUZIONALI_ = ["PRESIDENTE", "VICEPRESIDENTE", "SEGRETARIO", "TESORIERE", "CONSIGLIERE"];

function normalizzaCaricaIstituzionale_(carica) {
  return String(carica || "").trim().toUpperCase();
}

function isCaricaIstituzionale_(carica) {
  return CARICHE_ISTITUZIONALI_.indexOf(normalizzaCaricaIstituzionale_(carica)) >= 0;
}

function inviaNotificaCambioCarica_(socio, caricaPrecedente, caricaNuova, righeSoci, indiceRigaModificata) {
  var precedente = normalizzaCaricaIstituzionale_(caricaPrecedente);
  var nuova = normalizzaCaricaIstituzionale_(caricaNuova);
  if (precedente === nuova || (!isCaricaIstituzionale_(precedente) && !isCaricaIstituzionale_(nuova))) return false;

  var nomeSocio = [socio.nome, socio.cognome].filter(function (parte) {
    return String(parte || "").trim() !== "";
  }).join(" ").trim() || "Socio";
  var destinatari = {};

  function aggiungiDestinatario(email, nome, titolareEsistente) {
    var indirizzo = String(email || "").trim();
    if (!indirizzo) {
      console.error("Notifica carica non inviata: indirizzo email mancante per " + (nome || "un socio") + ".");
      return;
    }
    var chiave = indirizzo.toLowerCase();
    if (!destinatari[chiave]) {
      destinatari[chiave] = {
        email: indirizzo,
        nome: nome || "Socio",
        titolareEsistente: titolareEsistente
      };
    }
  }

  aggiungiDestinatario(socio.email, nomeSocio, false);

  if (isCaricaIstituzionale_(nuova)) {
    for (var i = 1; i < righeSoci.length; i++) {
      if (i === indiceRigaModificata) continue;
      var riga = righeSoci[i];
      if (normalizzaCaricaIstituzionale_(riga[COLONNA_CARICA_SOCIALE_ - 1]) !== nuova) continue;
      aggiungiDestinatario(
        riga[2],
        [riga[0], riga[4]].filter(function (parte) {
          return String(parte || "").trim() !== "";
        }).join(" ").trim(),
        true
      );
    }
  }

  var nomeCaricaPrecedente = String(caricaPrecedente || "").trim() || "Socio semplice";
  var nomeCaricaNuova = String(caricaNuova || "").trim() || "Socio semplice";
  var oggetto = "Aggiornamento cariche sociali - Gens Ssshhh";

  Object.keys(destinatari).forEach(function (chiave) {
    var destinatario = destinatari[chiave];
    var corpo;
    if (destinatario.titolareEsistente) {
      corpo = "Ciao " + destinatario.nome + ",\n\n" +
        "ti informiamo che la carica di " + nomeCaricaNuova + " è stata assegnata a " + nomeSocio + ".\n\n" +
        "La variazione riguarda la carica che ricopri attualmente. Per informazioni, contatta il consiglio direttivo.\n\n" +
        "Associazione Gens Ssshhh";
    } else {
      corpo = "Ciao " + destinatario.nome + ",\n\n" +
        "ti informiamo che la tua carica sociale è stata aggiornata da \"" + nomeCaricaPrecedente +
        "\" a \"" + nomeCaricaNuova + "\".\n\n" +
        "Associazione Gens Ssshhh";
    }

    MailApp.sendEmail({
      to: destinatario.email,
      subject: oggetto,
      body: corpo
    });
  });

  return true;
}

function gestisciModificaCaricaIstituzionale(e) {
  if (!e || !e.range) throw new Error("Evento di modifica del foglio mancante.");

  var intervallo = e.range;
  var foglio = intervallo.getSheet();
  var primaColonna = intervallo.getColumn();
  var ultimaColonna = primaColonna + intervallo.getNumColumns() - 1;
  if (foglio.getName().toLowerCase() !== "soci" ||
    COLONNA_CARICA_SOCIALE_ < primaColonna ||
    COLONNA_CARICA_SOCIALE_ > ultimaColonna ||
    intervallo.getRow() < 2) return;

  if (intervallo.getNumRows() !== 1 || intervallo.getNumColumns() !== 1) {
    console.error("Notifica carica non inviata: modificare una sola cella della colonna M per volta.");
    return;
  }

  var riga = intervallo.getRow();
  var righeSoci = foglio.getDataRange().getValues();
  var datiSocio = righeSoci[riga - 1];
  var caricaPrecedente = e.oldValue || "";
  var caricaNuova = intervallo.getValue() || "";

  inviaNotificaCambioCarica_({
    nome: datiSocio[0],
    cognome: datiSocio[4],
    email: datiSocio[2]
  }, caricaPrecedente, caricaNuova, righeSoci, riga - 1);
}

function configuraNotificheCaricheIstituzionali() {
  var nomeHandler = "gestisciModificaCaricaIstituzionale";
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === nomeHandler) ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger(nomeHandler)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}
