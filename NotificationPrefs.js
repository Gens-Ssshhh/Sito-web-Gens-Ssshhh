// Centro preferenze notifiche: ogni notifica del sistema può essere attivata o
// disattivata per canale (email / app). Per i canali "modificabili" il valore
// predefinito è sempre attivo, quindi senza preferenze salvate nulla cambia.
//
// Disponibilità del canale: 0 = non prevista, 1 = modificabile, 2 = obbligatoria.

var NOTIFICHE_PREFIX_PROPRIETA_ = "NOTIFICHE_PREF|";
var NOTIFICHE_CONFIGURAZIONE_PROPRIETA_ = "NOTIFICHE_CONFIGURAZIONE";
var cacheConfigurazioneNotifiche_ = null;

var NOTIFICHE_CATALOGO_ = [
  { id: "password_temporanea", gruppo: "Account e sicurezza", titolo: "Password temporanea", descrizione: "Invio della nuova password dopo un reset.", email: 2, app: 1 },
  { id: "reset_password", gruppo: "Account e sicurezza", titolo: "Link per reimpostare la password", descrizione: "Link per scegliere una nuova password.", email: 2, app: 1 },
  { id: "firma_otp", gruppo: "Account e sicurezza", titolo: "Codice OTP per la firma", descrizione: "Codice di sicurezza richiesto per firmare un documento.", email: 2, app: 1 },
  { id: "tessera_scadenza", gruppo: "Tessera e quota", titolo: "Scadenza della tessera", descrizione: "Avviso a 30 giorni dalla scadenza e quando la tessera risulta scaduta.", email: 1, app: 1 },
  { id: "fattura_ricevuta", gruppo: "Tessera e quota", titolo: "Fatture di cortesia", descrizione: "Fattura o promemoria di pagamento inviati dall'associazione.", email: 1, app: 0 },
  { id: "news_pubblicate", gruppo: "Comunicazioni", titolo: "Nuovi avvisi in bacheca", descrizione: "Avvisi e news pubblicati dall'associazione.", email: 0, app: 1 },
  { id: "digest_settimanale", gruppo: "Comunicazioni", titolo: "Novità della settimana", descrizione: "Riepilogo settimanale delle news della bacheca.", email: 1, app: 1 },
  { id: "carica_cambio", gruppo: "Comunicazioni", titolo: "Cambi di carica sociale", descrizione: "Variazioni delle cariche istituzionali che ti riguardano.", email: 1, app: 0 },
  { id: "firma_richiesta", gruppo: "Firme", titolo: "Documenti da firmare", descrizione: "Documenti in attesa della tua firma.", email: 0, app: 1 },
  { id: "firma_sollecito", gruppo: "Firme", titolo: "Solleciti di firma", descrizione: "Promemoria per i documenti non ancora firmati.", email: 1, app: 1 },
  { id: "firma_completata", gruppo: "Firme", titolo: "Conferma di firma", descrizione: "Certificato inviato dopo aver firmato un documento.", email: 1, app: 1 },
  { id: "firma_controfirma", gruppo: "Firme", titolo: "Documento controfirmato", descrizione: "Il documento è stato controfirmato dall'amministrazione.", email: 1, app: 1 },
  { id: "firma_archiviazione", gruppo: "Firme", titolo: "Pratica di firma archiviata", descrizione: "Registro inviato quando una pratica viene chiusa.", email: 1, app: 1, admin: true },
  { id: "consultazione_aperta", gruppo: "Consultazioni", titolo: "Consultazioni aperte", descrizione: "Consultazioni a cui puoi partecipare.", email: 0, app: 1 },
  { id: "candidature_aperte", gruppo: "Consultazioni", titolo: "Candidature aperte", descrizione: "Elezioni per cui puoi presentare una candidatura.", email: 0, app: 1 },
  { id: "voto_conferma", gruppo: "Consultazioni", titolo: "Conferma del voto", descrizione: "Ricevuta dopo aver registrato un voto.", email: 1, app: 1 },
  { id: "promemoria_elettorale", gruppo: "Consultazioni", titolo: "Promemoria elettorali", descrizione: "Promemoria di partecipazione alle consultazioni.", email: 1, app: 1 },
  { id: "report_elettorale", gruppo: "Consultazioni", titolo: "Report di invio elettorale", descrizione: "Ricevuta dell'invio delle comunicazioni elettorali.", email: 1, app: 1, admin: true },
  { id: "evento_nuovo", gruppo: "Eventi", titolo: "Nuovi questionari evento", descrizione: "Un nuovo evento è stato proposto e richiede la tua risposta.", email: 1, app: 1 },
  { id: "evento_promemoria", gruppo: "Eventi", titolo: "Promemoria questionari evento", descrizione: "Promemoria dopo tre giorni senza risposta.", email: 1, app: 1 },
  { id: "evento_ricompila", gruppo: "Eventi", titolo: "Richieste di ricompilazione", descrizione: "Date, orari o luoghi di un evento sono cambiati.", email: 1, app: 1 },
  { id: "evento_modifica_creatore", gruppo: "Eventi", titolo: "Modifiche ai tuoi eventi", descrizione: "Conferma delle modifiche a un evento che hai proposto.", email: 1, app: 0 },
  { id: "evento_nuova_opzione", gruppo: "Eventi", titolo: "Nuove date proposte", descrizione: "Qualcuno ha aggiunto una data a un evento che hai proposto.", email: 1, app: 0 },
  { id: "ammissione_sostegno", gruppo: "Ammissione soci", titolo: "Richieste di sostegno", descrizione: "Ti viene chiesto di sostenere una candidatura.", email: 1, app: 1 },
  { id: "ammissione_nuova_richiesta", gruppo: "Ammissione soci", titolo: "Nuove richieste di iscrizione", descrizione: "Una persona ha chiesto di diventare socia e attende avvalli.", email: 1, app: 1 },
  { id: "ammissione_voto_aperto", gruppo: "Ammissione soci", titolo: "Voto di ammissione aperto", descrizione: "È aperta una votazione su una richiesta di iscrizione.", email: 1, app: 1 },
  { id: "dimissioni_richiesta", gruppo: "Amministrazione", titolo: "Richieste di dimissioni", descrizione: "Un socio ha richiesto le dimissioni.", email: 1, app: 1, admin: true }
];

var cachePreferenzeNotifiche_ = {};

function trovaNotificaCatalogo_(id) {
  for (var i = 0; i < NOTIFICHE_CATALOGO_.length; i++) {
    if (NOTIFICHE_CATALOGO_[i].id === id) return NOTIFICHE_CATALOGO_[i];
  }
  return null;
}

function leggiPreferenzeNotifiche_(email) {
  var chiave = String(email || "").trim().toLowerCase();
  if (!chiave) return {};
  if (cachePreferenzeNotifiche_[chiave]) return cachePreferenzeNotifiche_[chiave];
  var preferenze = {};
  try {
    var salvato = PropertiesService.getScriptProperties().getProperty(NOTIFICHE_PREFIX_PROPRIETA_ + chiave);
    if (salvato) preferenze = JSON.parse(salvato) || {};
  } catch (e) { }
  cachePreferenzeNotifiche_[chiave] = preferenze;
  return preferenze;
}

// canale: "email" oppure "app" (notifiche push e pannello dell'app).
function notificaAbilitata_(email, id, canale) {
  var voce = trovaNotificaCatalogo_(id);
  if (!voce) return true;
  var disponibilita = voce[canale];
  if (disponibilita === 0) return false;
  if (disponibilita === 2) return true;
  var configurazione = leggiConfigurazioneNotifiche_();
  if (!configurazione.attive || configurazione.notifiche[id] === false) return false;
  var scelta = leggiPreferenzeNotifiche_(email)[id];
  return !(scelta && scelta[canale === "email" ? "e" : "a"] === 0);
}

function leggiConfigurazioneNotifiche_() {
  if (cacheConfigurazioneNotifiche_) return cacheConfigurazioneNotifiche_;
  var configurazione = { attive: true, notifiche: {} };
  var salvato = PropertiesService.getScriptProperties().getProperty(NOTIFICHE_CONFIGURAZIONE_PROPRIETA_);
  if (salvato) {
    try {
      var letta = JSON.parse(salvato) || {};
      if (typeof letta.attive === "boolean") configurazione.attive = letta.attive;
      if (letta.notifiche && typeof letta.notifiche === "object") configurazione.notifiche = letta.notifiche;
    } catch (e) {
      console.error("Configurazione globale notifiche non valida:", e);
      configurazione.attive = false;
    }
  }
  cacheConfigurazioneNotifiche_ = configurazione;
  return configurazione;
}

function utenteEsegretarioNotifiche_(email) {
  var utente = getDatiUtente(email);
  var ruolo = String(utente && utente.ruolo || "").trim().toUpperCase().replace(/[\s-]/g, "");
  return ruolo === "SEGRETARIO";
}

function getConfigurazioneNotificheGlobali(email) {
  if (!utenteEsegretarioNotifiche_(email)) return { ok: false };
  var configurazione = leggiConfigurazioneNotifiche_();
  return {
    ok: true,
    attive: configurazione.attive,
    notifiche: NOTIFICHE_CATALOGO_.map(function (voce) {
      return {
        id: voce.id,
        gruppo: voce.gruppo,
        titolo: voce.titolo,
        descrizione: voce.descrizione,
        attiva: configurazione.notifiche[voce.id] !== false,
        obbligatoria: voce.email === 2 || voce.app === 2
      };
    })
  };
}

function salvaConfigurazioneNotificheGlobali(dati) {
  dati = dati || {};
  if (!utenteEsegretarioNotifiche_(dati.email)) return { ok: false };

  var configurazione = leggiConfigurazioneNotifiche_();
  if (typeof dati.attive === "boolean") configurazione.attive = dati.attive;
  if (dati.notifiche && typeof dati.notifiche === "object" && !Array.isArray(dati.notifiche)) {
    Object.keys(dati.notifiche).forEach(function (id) {
      if (trovaNotificaCatalogo_(id) && typeof dati.notifiche[id] === "boolean") {
        configurazione.notifiche[id] = dati.notifiche[id];
      }
    });
  }
  PropertiesService.getScriptProperties().setProperty(
    NOTIFICHE_CONFIGURAZIONE_PROPRIETA_,
    JSON.stringify(configurazione)
  );
  cacheConfigurazioneNotifiche_ = configurazione;
  return getConfigurazioneNotificheGlobali(dati.email);
}

function filtraDestinatariNotifica_(emails, id, canale) {
  return (emails || []).filter(function (email) {
    return notificaAbilitata_(email, id, canale);
  });
}

function getPreferenzeNotifiche(email) {
  var utente = getDatiUtente(email);
  if (!utente) return { notifiche: [] };
  var admin = typeof isRuoloAmministrativo === "function" && isRuoloAmministrativo(utente.ruolo);
  var notifiche = NOTIFICHE_CATALOGO_.filter(function (voce) {
    return !voce.admin || admin;
  }).map(function (voce) {
    function canale(nome) {
      if (!voce[nome]) return null;
      return { attiva: notificaAbilitata_(email, voce.id, nome), bloccata: voce[nome] === 2 };
    }
    return {
      id: voce.id,
      gruppo: voce.gruppo,
      titolo: voce.titolo,
      descrizione: voce.descrizione,
      soloAdmin: !!voce.admin,
      email: canale("email"),
      app: canale("app")
    };
  });
  return { notifiche: notifiche };
}

// dati: { email, preferenze: { idNotifica: { email: bool, app: bool } } }
function salvaPreferenzeNotifiche(dati) {
  dati = dati || {};
  var email = String(dati.email || "").trim().toLowerCase();
  if (!email || !getDatiUtente(email) || !dati.preferenze || typeof dati.preferenze !== "object") return { ok: false };

  var preferenze = {};
  var correnti = leggiPreferenzeNotifiche_(email);
  Object.keys(correnti).forEach(function (id) { preferenze[id] = { e: correnti[id].e, a: correnti[id].a }; });

  Object.keys(dati.preferenze).forEach(function (id) {
    var voce = trovaNotificaCatalogo_(id);
    var richiesta = dati.preferenze[id];
    if (!voce || !richiesta || typeof richiesta !== "object") return;
    var scelta = preferenze[id] || {};
    if (voce.email === 1 && typeof richiesta.email === "boolean") scelta.e = richiesta.email ? 1 : 0;
    if (voce.app === 1 && typeof richiesta.app === "boolean") scelta.a = richiesta.app ? 1 : 0;
    if (scelta.e === 0 || scelta.a === 0) preferenze[id] = { e: scelta.e === 0 ? 0 : 1, a: scelta.a === 0 ? 0 : 1 };
    else delete preferenze[id];
  });

  var proprieta = PropertiesService.getScriptProperties();
  var chiave = NOTIFICHE_PREFIX_PROPRIETA_ + email;
  if (Object.keys(preferenze).length) proprieta.setProperty(chiave, JSON.stringify(preferenze));
  else if (typeof proprieta.deleteProperty === "function") proprieta.deleteProperty(chiave);
  else proprieta.setProperty(chiave, "{}");
  cachePreferenzeNotifiche_[email] = preferenze;
  return { ok: true, notifiche: getPreferenzeNotifiche(email).notifiche };
}

function ripristinaPreferenzeNotifiche(dati) {
  dati = dati || {};
  var email = String(typeof dati === "string" ? dati : dati.email || "").trim().toLowerCase();
  if (!email || !getDatiUtente(email)) return { ok: false };
  var proprieta = PropertiesService.getScriptProperties();
  if (typeof proprieta.deleteProperty === "function") proprieta.deleteProperty(NOTIFICHE_PREFIX_PROPRIETA_ + email);
  else proprieta.setProperty(NOTIFICHE_PREFIX_PROPRIETA_ + email, "{}");
  cachePreferenzeNotifiche_[email] = {};
  return { ok: true, notifiche: getPreferenzeNotifiche(email).notifiche };
}
