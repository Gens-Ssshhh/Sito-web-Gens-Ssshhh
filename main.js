// ==========================================
// 1. INIZIALIZZAZIONE E UTILITÀ
// ==========================================

function doGet(e) {
  // Il backend ora funge solo da API invisibile. Nessuna interfaccia HTML viene generata.
  return ContentService.createTextOutput("Gens Ssshhh Backend API - Attivo e funzionante");
}

function autorizzaDriveVerbali() {
  var folderId = "1B40jXYlF9nXvR7rKGtIjpWYl4SMgGrXH";
  var folder = DriveApp.getFolderById(folderId);
  return "DRIVE_OK: " + folder.getName();
}


var SECRET_SALT = "INCULATI-BRUTTO-GAYNEU37Y";
var SESSION_CACHE_PREFIX = "gens_session_";
var SESSION_TTL_SECONDS = 15552000;
var AZIONI_API_CONSENTITE = [
  "adminCambiaStatoVoto", "adminCancellaNews", "adminChiudiEGeneraRegistroFirme",
  "adminChiudiEGeneraVerbale", "adminCreaNuovaElezioneAvanzata", "adminEliminaElezione",
  "adminEliminaGruppoRichieste", "adminEseguiControfirma", "adminGestisciAmmissione",
  "adminGetCandidaturePendenti", "adminGetListaElezioni", "adminGetListaPagamenti",
  "adminGetRisultatiLive", "adminGetStatisticheFirme", "adminGetUltimiMovimentiPD",
  "adminInviaDocumentoFirma", "adminInviaSollecito", "adminProcessaCandidatura",
  "adminPubblicaNews", "adminRegistraRinnovo", "adminRigeneraVerbale", "adminUpdateSocio",
  "aggiungiOpzionePropostaEvento", "aggiungiSpesaGruppo", "annullaPropostaEvento",
  "avvallaRichiestaIscrizione", "calcolaSaldiGruppo", "cambiaPassword",
  "checkAccettazioneRegolamento", "checkStatoVoto", "checkStatoVotoCampagna",
  "completaResetPassword", "confermaEventoProposto", "configuraTriggerAmmissioni",
  "creaGruppoSpese", "creaPropostaEvento", "eliminaReportArchivio", "exportBilancioExcel",
  "finalizzaFirmaRichiesta", "generaFatturaXML", "generaTesseraPDF", "getAmmissioniAdmin",
  "getAmmissioniInVoto", "getAnagrafeFornitoriDinamica", "getAvvisiPubblici",
  "getCandidatureAttive", "getConsultazioniAttiveUtente", "getCostiVacanzeDashboard",
  "getDashboardAdmin", "getDatiBilancio", "getDatiBilancioCompleto", "getDatiUtente",
  "getElezioniPerCandidatura", "getGruppiSpese", "getListaDocumenti", "getListaFornitori",
  "getListaSociPerSponsor", "getNomiSociAttivi", "getNotificheUtente",
  "getPianoDeiContiDinamico", "getPreferenzeNotifiche", "getPropostaEventoPubblica", "getProposteEventi",
  "getReportStorici", "ripristinaPreferenzeNotifiche", "salvaPreferenzeNotifiche", "getRichiesteFirmaUtente", "getRichiesteIscrizioneDaAvvallare",
  "getRispostaRicompilazioneOspite", "getRisultatiLive", "getSociPerFattura",
  "getStartData", "getStoricoFirmeUtente", "inviaCandidatura", "inviaLinkReset",
  "inviaRichiestaDimissioni", "login", "logout", "modificaPropostaEvento",
  "ottieniUrlVerbale", "proponiNuovoSocio", "registraAccettazioneRegolamento",
  "registraMovimentoCompostoServer", "registraTokenPush", "requestSignOTP",
  "riceviVoto", "riceviVotoCampagna", "richiediIscrizioneSocio",
  "ricompilaVotoPropostaEvento", "rigeneraCertificatoUtente", "rimuoviTokenPush",
  "salvaDatiUtente", "salvaFatturaRicevuta", "salvaReportArchivio", "sostieniCandidato",
  "spostaFileUtente", "uploadFile", "utenteAggiungeOdG", "votaAmmissioneSocio",
  "votaPropostaEvento"
];

function creaSessione_(email) {
  var token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    SESSION_CACHE_PREFIX + token,
    JSON.stringify({
      email: String(email || "").trim().toLowerCase(),
      expiresAt: Date.now() + SESSION_TTL_SECONDS * 1000
    })
  );
  return token;
}

function getEmailSessione_(token) {
  if (typeof token !== "string" || !token || token.length > 200) return "";
  var chiave = SESSION_CACHE_PREFIX + token;
  var proprieta = PropertiesService.getScriptProperties();
  var sessione = proprieta.getProperty(chiave);
  if (!sessione) return "";
  try {
    var datiSessione = JSON.parse(sessione);
    if (!datiSessione.email || !datiSessione.expiresAt || Date.now() >= datiSessione.expiresAt) {
      proprieta.deleteProperty(chiave);
      return "";
    }
    return datiSessione.email;
  } catch (err) {
    proprieta.deleteProperty(chiave);
    return "";
  }
}

function revocaSessione_(token) {
  if (typeof token === "string" && token && token.length <= 200) {
    CacheService.getScriptCache().remove(SESSION_CACHE_PREFIX + token);
    PropertiesService.getScriptProperties().deleteProperty(SESSION_CACHE_PREFIX + token);
  }
}

function vincolaIdentitaAdmin_(valore, email) {
  if (Array.isArray(valore)) {
    return valore.map(function (elemento) { return vincolaIdentitaAdmin_(elemento, email); });
  }
  if (!valore || typeof valore !== "object") return valore;

  var risultato = {};
  Object.keys(valore).forEach(function (chiave) {
    if (chiave === "adminEmail" || chiave === "emailAdmin") {
      risultato[chiave] = email;
    } else {
      risultato[chiave] = vincolaIdentitaAdmin_(valore[chiave], email);
    }
  });
  return risultato;
}

function creaHash(input) {
  // 1. Assicurati che input sia una stringa (evita errori se è null/undefined)
  var text = (input || "").toString();

  // 2. Aggiunge il SALT (la tua frase segreta definita sopra)
  var textSalted = text + SECRET_SALT;

  // 3. Calcola l'hash SHA-256 (restituisce un array di byte con segno)
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, textSalted);

  // 4. Converte i byte in una stringa esadecimale pulita
  var txtHash = "";
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256; // Correzione per i byte negativi
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0'; // Aggiunge lo zero iniziale se serve (padding)
    }
    txtHash += hashVal.toString(16);
  }

  return txtHash;
}

function doPost(e) {
  try {
    var richiesta = JSON.parse(e.postData.contents);
    var azione = richiesta.azione;

    // Login and explicitly public workflows are the only anonymous entry points.
    if (azione === "login") {
      var payloadLogin = richiesta.payload || {};
      var emailLogin = richiesta.email || payloadLogin.email;
      var passwordLogin = richiesta.password || payloadLogin.password;
      var esito = verificaLogin({ email: emailLogin, password: passwordLogin, info: "Accesso Mobile" });
      if (esito === "OK_LOGIN") {
        var tokenSessione = creaSessione_(emailLogin);
        var datiUtente = getDatiUtente(emailLogin);
        return ContentService.createTextOutput(JSON.stringify({
          status: "SUCCESS",
          token: tokenSessione,
          dati: datiUtente,
          data: { token: tokenSessione, utente: datiUtente }
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", messaggio: "Credenziali errate." })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    var azioniPubbliche = [
      "completaResetPassword",
      "getPropostaEventoPubblica",
      "votaPropostaEvento",
      "richiediIscrizioneSocio"
    ];
    var emailSessione = getEmailSessione_(richiesta.token);
    if (azioniPubbliche.indexOf(azione) < 0 && !emailSessione) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "AUTH_ERROR",
        messaggio: "Sessione non valida o scaduta. Effettua di nuovo l'accesso."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (azione === "logout") {
      revocaSessione_(richiesta.token);
      return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: true })).setMimeType(ContentService.MimeType.JSON);
    }

    if (AZIONI_API_CONSENTITE.indexOf(azione) < 0) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ERROR",
        messaggio: "Azione non consentita tramite API."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Results for event proposals are public and contain aggregated counts only.
    if (azione === "getProposteEventi") {
      var payloadProposte = richiesta.payload;
      var emailProposte = Array.isArray(payloadProposte) ? payloadProposte[0] : (typeof payloadProposte === "string" ? payloadProposte : payloadProposte && payloadProposte.email);
      return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: getProposteEventi(emailProposte || richiesta.email || "") })).setMimeType(ContentService.MimeType.JSON);
    }

    if (azione === "getPropostaEventoPubblica") {
      return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: getPropostaEventoPubblica(richiesta.payload) })).setMimeType(ContentService.MimeType.JSON);
    }

    if (azione === "creaPropostaEvento") {
      var payloadEvento = richiesta.payload;
      if (Array.isArray(payloadEvento)) {
        var risultatoCreazione = payloadEvento.length >= 3
          ? creaPropostaEvento(payloadEvento[0], payloadEvento[1], payloadEvento[2] || {})
          : { ok: false, messaggio: "Accedi al portale per proporre un evento." };
      } else {
        var risultatoCreazione = creaPropostaEvento(payloadEvento.email || richiesta.email || "", payloadEvento.password, payloadEvento.proposta || payloadEvento);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: risultatoCreazione })).setMimeType(ContentService.MimeType.JSON);
    }

    if (azione === "votaPropostaEvento") {
      var payloadVoto = richiesta.payload;
      if (Array.isArray(payloadVoto)) {
        var risultatoVoto = votaPropostaEvento(payloadVoto[0], payloadVoto[1], payloadVoto[2]);
      } else {
        var risultatoVoto = votaPropostaEvento(payloadVoto.email || richiesta.email || "", payloadVoto.propostaId || payloadVoto.id, payloadVoto.risposta || payloadVoto.voto || payloadVoto);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "SUCCESS", data: risultatoVoto })).setMimeType(ContentService.MimeType.JSON);
    }

    if (azione === "processaScadenzeAmmissioni" || azione === "configuraNotificheCaricheIstituzionali" || azione === "gestisciModificaCaricaIstituzionale") {
      return ContentService.createTextOutput(JSON.stringify({ status: "ERROR", messaggio: "Azione interna non invocabile via API." })).setMimeType(ContentService.MimeType.JSON);
    }

    var azioneAmministrativa = /^admin/i.test(azione || "") || /Admin$/.test(azione || "") || azione === "exportBilancioExcel" || azione === "configuraTriggerAmmissioni";
    if (azioneAmministrativa) {
      var utenteSessione = getDatiUtente(emailSessione);
      if (!utenteSessione || !isRuoloAmministrativo(utenteSessione.ruolo)) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "ERROR",
          messaggio: "Operazione riservata agli amministratori."
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Sensitive identity parameters must always come from the validated session.
    if (azione === "getDatiUtente" || azione === "getDashboardAdmin" || azione === "getPreferenzeNotifiche" || azione === "ripristinaPreferenzeNotifiche") {
      richiesta.payload = emailSessione;
    } else if (azione === "salvaPreferenzeNotifiche") {
      var payloadPreferenze = richiesta.payload && typeof richiesta.payload === "object" ? richiesta.payload : {};
      richiesta.payload = { email: emailSessione, preferenze: payloadPreferenze.preferenze };
    } else if (/^admin/i.test(azione || "") && richiesta.payload && typeof richiesta.payload === "object") {
      richiesta.payload = vincolaIdentitaAdmin_(richiesta.payload, emailSessione);
    }

    // Dynamic dispatch is retained for authenticated application actions only.
    if (typeof this[azione] === 'function' && !/_$/.test(azione)) {
      // Estrae i parametri in modo flessibile a seconda di come li invia l'app
      var param = richiesta.payload !== undefined ? richiesta.payload :
        richiesta.email !== undefined ? richiesta.email :
          richiesta.dati !== undefined ? richiesta.dati : richiesta;

      var indiceIdentitaAdmin = {
        adminGetListaPagamenti: 0,
        adminRegistraRinnovo: 0,
        adminGetListaElezioni: 0,
        adminEliminaElezione: 0,
        adminCambiaStatoVoto: 0,
        adminGetCandidaturePendenti: 0,
        adminProcessaCandidatura: 0,
        adminGetRisultatiLive: 0,
        getAmmissioniAdmin: 0,
        adminGestisciAmmissione: 0,
        adminGetStatisticheFirme: 0,
        adminEseguiControfirma: 1,
        adminChiudiEGeneraRegistroFirme: 1,
        adminEliminaGruppoRichieste: 1,
        adminGetUltimiMovimentiPD: 0,
        adminChiudiEGeneraVerbale: 0,
        adminRigeneraVerbale: 0,
        configuraTriggerAmmissioni: 0,
        exportBilancioExcel: 0
      };
      var indiceIdentita = indiceIdentitaAdmin[azione];
      if (indiceIdentita !== undefined) {
        if (Array.isArray(param)) {
          param = param.slice();
          param[indiceIdentita] = emailSessione;
        } else if (indiceIdentita === 0) {
          param = emailSessione;
        }
      } else if (/^admin/i.test(azione || "") && param && typeof param === "object") {
        param = vincolaIdentitaAdmin_(param, emailSessione);
        if (azione === "adminPubblicaNews" || azione === "adminCancellaNews") param.email = emailSessione;
      }

      var risultato;
      if (Array.isArray(param)) {
        risultato = this[azione].apply(this, param); // Spacchetta l'array nei vari argomenti
      } else {
        risultato = this[azione](param);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "SUCCESS",
        data: risultato
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. SE LA FUNZIONE NON ESISTE
    return ContentService.createTextOutput(JSON.stringify({
      status: "ERROR",
      messaggio: "Azione (" + azione + ") non trovata sul server."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "FATAL_ERROR", error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
