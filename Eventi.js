var EVENTI_SHEET_NAME = 'ProposteEventi';
var EVENTI_CALENDAR_ID = 'gens.ssshhh@gmail.com';

function getFoglioProposteEventi_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EVENTI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EVENTI_SHEET_NAME);
    sheet.appendRow(['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato']);
  }
  var headers = sheet.getDataRange().getValues()[0] || [];
  [['EventoCalendarId', 8], ['OpzioneConfermataJson', 9], ['PromemoriaNonVotantiInviatoIl', 10]].forEach(function (item) {
    if (String(headers[item[1]] || '') !== item[0]) sheet.getRange(1, item[1] + 1).setValue(item[0]);
  });
  return sheet;
}

function verificaSocioEvento_(email) {
  var target = String(email || '').trim().toLowerCase();
  if (!target) return false;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('soci');
  if (!sheet) return false;

  var rows = sheet.getDataRange().getValues();
  var headers = rows.length ? rows[0].map(function (value) { return String(value).trim().toLowerCase(); }) : [];
  var statusColumn = headers.indexOf('stato socio');
  return rows.slice(1).some(function (row) {
    var matches = String(row[2] || '').trim().toLowerCase() === target;
    if (!matches || statusColumn < 0) return matches;
    return String(row[statusColumn] || '').trim().toLowerCase() === 'attivo';
  });
}

function normalizzaOpzioniEvento_(opzioni) {
  if (!Array.isArray(opzioni)) return [];
  return opzioni.map(function (opzione) {
    return {
      data: String(opzione.data || '').trim(),
      ora: String(opzione.ora || '').trim(),
      luogo: String(opzione.luogo || '').trim()
    };
  }).filter(function (opzione) {
    return opzione.data && opzione.ora;
  });
}

function dataOraEvento_(opzione) {
  var matchData = String(opzione.data || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  var matchOra = String(opzione.ora || '').match(/^(\d{2}):(\d{2})$/);
  if (!matchData || !matchOra) return null;
  var data = new Date(Number(matchData[3]), Number(matchData[2]) - 1, Number(matchData[1]), Number(matchOra[1]), Number(matchOra[2]), 0, 0);
  if (data.getFullYear() !== Number(matchData[3]) || data.getMonth() !== Number(matchData[2]) - 1 || data.getDate() !== Number(matchData[1]) || data.getHours() !== Number(matchOra[1]) || data.getMinutes() !== Number(matchOra[2])) return null;
  return data;
}

function hashEmailEvento_(email) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(email).trim().toLowerCase());
  return bytes.map(function (byte) {
    var valore = byte < 0 ? byte + 256 : byte;
    return ('0' + valore.toString(16)).slice(-2);
  }).join('');
}

function leggiVotiEvento_(valore, opzioni, emailSocio) {
  var dati = {};
  try { dati = JSON.parse(valore || '{}'); } catch (e) { }

  var conteggiOpzioni = {};
  if (dati.opzioni && typeof dati.opzioni === 'object') {
    conteggiOpzioni = dati.opzioni;
  } else {
    Object.keys(dati).forEach(function (chiave) {
      if (/^\d+$/.test(chiave)) conteggiOpzioni[chiave] = Number(dati[chiave] || 0);
    });
  }

  var interesse = dati.interesse && typeof dati.interesse === 'object'
    ? { si: Number(dati.interesse.si || 0), no: Number(dati.interesse.no || 0) }
    : { si: Object.keys(conteggiOpzioni).reduce(function (totale, chiave) { return totale + Number(conteggiOpzioni[chiave] || 0); }, 0), no: 0 };

  var votanti = dati._votanti && typeof dati._votanti === 'object' ? dati._votanti : {};
  var mioVoto = null;
  if (emailSocio) {
    var chiaveSocio = 'socio:' + emailSocio;
    var precedente = votanti[chiaveSocio] !== undefined ? votanti[chiaveSocio] : votanti[emailSocio];
    if (precedente !== undefined) {
      mioVoto = typeof precedente === 'number'
        ? { interesse: 'si', opzioni: [precedente] }
        : precedente;
    }
  }

  var ricompilazione = dati._ricompilazione && typeof dati._ricompilazione === 'object' ? dati._ricompilazione : {};

  return {
    riepilogo: {
      interesse: interesse,
      opzioni: opzioni.map(function (_, indice) { return Number(conteggiOpzioni[indice] || 0); })
    },
    votanti: votanti,
    mioVoto: mioVoto,
    ricompilazione: ricompilazione
  };
}

function datiVotiDaSalvare_(riepilogo, votanti, ricompilazione) {
  var dati = { interesse: riepilogo.interesse, opzioni: riepilogo.opzioni, _votanti: votanti };
  if (ricompilazione && Object.keys(ricompilazione).length) dati._ricompilazione = ricompilazione;
  return dati;
}

function rimuoviContributoVotoEvento_(riepilogo, votanti, chiave) {
  var precedente = votanti[chiave];
  if (precedente === undefined) return;
  var risposta = typeof precedente === 'number' ? { interesse: 'si', opzioni: [precedente] } : precedente;
  if (risposta.interesse === 'si' || risposta.interesse === 'no') {
    riepilogo.interesse[risposta.interesse] = Math.max(0, Number(riepilogo.interesse[risposta.interesse] || 0) - 1);
  }
  (Array.isArray(risposta.opzioni) ? risposta.opzioni : []).forEach(function (indice) {
    var numero = Number(indice);
    if (riepilogo.opzioni[numero] !== undefined) riepilogo.opzioni[numero] = Math.max(0, riepilogo.opzioni[numero] - 1);
  });
  delete votanti[chiave];
}

function chiaveOpzioneEvento_(opzione) {
  return [opzione.data, opzione.ora, opzione.luogo || ''].join('|');
}

function slotOpzioneEvento_(opzione) {
  return [opzione.data, opzione.ora].join('|');
}

function opzioniEventoCambiate_(vecchie, nuove) {
  var prima = vecchie.map(chiaveOpzioneEvento_).sort().join('\n');
  var dopo = nuove.map(chiaveOpzioneEvento_).sort().join('\n');
  return prima !== dopo;
}

function haSlotDuplicatiEvento_(opzioni) {
  return opzioni.some(function (opzione, indice) {
    return opzioni.findIndex(function (altra) { return slotOpzioneEvento_(altra) === slotOpzioneEvento_(opzione); }) !== indice;
  });
}

function chiaveRicompilazioneSocio_(ricompilazione, email) {
  var candidate = ['socio:' + email, email];
  for (var i = 0; i < candidate.length; i++) {
    if (ricompilazione[candidate[i]] !== undefined) return candidate[i];
  }
  return null;
}

function contattoVotanteEvento_(chiave, risposta) {
  if (chiave.indexOf('ospite:') === 0) {
    var emailOspite = risposta && typeof risposta === 'object' ? String(risposta.email || '').trim().toLowerCase() : '';
    return { socio: false, email: emailOspite };
  }
  return { socio: true, email: chiave.replace(/^socio:/, '').trim().toLowerCase() };
}

function nomiSociEvento_() {
  var nomi = {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('soci');
  if (!sheet) return nomi;
  var rows = sheet.getDataRange().getValues();
  if (!rows.length) return nomi;
  var headers = rows[0].map(function (value) { return String(value).trim().toLowerCase(); });
  var colonna = function (nome, predefinita) {
    var indice = headers.indexOf(nome);
    return indice >= 0 ? indice : predefinita;
  };
  var emailIndex = colonna('email', 2);
  var nomeIndex = colonna('nome', 0);
  var cognomeIndex = colonna('cognome', 4);
  rows.slice(1).forEach(function (row) {
    var email = String(row[emailIndex] || '').trim().toLowerCase();
    if (email) nomi[email] = (String(row[nomeIndex] || '').trim() + ' ' + String(row[cognomeIndex] || '').trim()).trim();
  });
  return nomi;
}

function nomeVotanteEvento_(chiave, risposta, nomiSoci) {
  if (chiave.indexOf('ospite:') === 0) return String(risposta.nome || '').trim() || 'Ospite';
  var email = chiave.replace(/^socio:/, '').trim().toLowerCase();
  return nomiSoci[email] || email;
}

function votantiPerOpzioneEvento_(votanti, opzioni) {
  var nomiSoci = nomiSociEvento_();
  var perOpzione = opzioni.map(function () { return []; });
  Object.keys(votanti).forEach(function (chiave) {
    var risposta = votanti[chiave];
    if (!risposta || typeof risposta !== 'object' || !Array.isArray(risposta.opzioni)) return;
    var nome = nomeVotanteEvento_(chiave, risposta, nomiSoci);
    risposta.opzioni.forEach(function (indice) {
      var numero = Number(indice);
      if (Number.isInteger(numero) && perOpzione[numero]) perOpzione[numero].push(nome);
    });
  });
  return perOpzione.map(function (nomi) {
    return nomi.sort(function (a, b) { return a.localeCompare(b, 'it'); });
  });
}

function indirizzoVotanteEvento_(chiave, risposta) {
  var emailChiave = chiave.indexOf('socio:') === 0 ? chiave.substring(6) : chiave;
  return String(risposta && typeof risposta === 'object' && risposta.email ? risposta.email : emailChiave).trim().toLowerCase();
}

function votantiSiEvento_(votanti, soloDataSelezionata, indiciDataSelezionata) {
  return Object.keys(votanti).filter(function (chiave) {
    var risposta = votanti[chiave];
    if (typeof risposta === 'number') {
      if (risposta < 0 || soloDataSelezionata) return false;
    } else if (!risposta || risposta.interesse !== 'si') {
      return false;
    }
    if (soloDataSelezionata && !(risposta.opzioni || []).some(function (indice) {
      return indiciDataSelezionata.indexOf(indice) >= 0;
    })) return false;
    return true;
  });
}

function partecipantiSiEvento_(votanti) {
  var nomiSoci = nomiSociEvento_();
  return votantiSiEvento_(votanti, false, []).map(function (chiave) {
    return { chiave: chiave, nome: nomeVotanteEvento_(chiave, votanti[chiave], nomiSoci) };
  });
}

function preparaPropostaEvento_(row, email) {
  var opzioni = [];
  try { opzioni = JSON.parse(row[5] || '[]'); } catch (e) { }
  var socio = verificaSocioEvento_(email);
  var voti = leggiVotiEvento_(row[6], opzioni, socio ? email : '');
  var creatore = String(row[2] || '').trim().toLowerCase() === String(email || '').trim().toLowerCase();
  var chiaveRicompilazione = socio ? chiaveRicompilazioneSocio_(voti.ricompilazione, email) : null;
  var ricompilazioneFacoltativa = chiaveRicompilazione !== null &&
    String(voti.ricompilazione[chiaveRicompilazione]).indexOf('FACOLTATIVA:') === 0;
  return {
    id: String(row[0]),
    titolo: String(row[3] || ''),
    descrizione: String(row[4] || ''),
    opzioni: opzioni,
    creataIl: row[1] instanceof Date ? row[1].toISOString() : String(row[1] || ''),
    voti: voti.riepilogo,
    votantiPerOpzione: socio ? votantiPerOpzioneEvento_(voti.votanti, opzioni) : [],
    partecipantiSi: creatore ? partecipantiSiEvento_(voti.votanti) : [],
    mioVoto: voti.mioVoto,
    ricompilazioneRichiesta: !!(chiaveRicompilazione !== null && !ricompilazioneFacoltativa),
    ricompilazioneFacoltativa: ricompilazioneFacoltativa,
    creatore: creatore
  };
}

function getProposteEventi(email) {
  var socio = String(email || '').trim().toLowerCase();
  if (!verificaSocioEvento_(socio)) return [];

  var values = getFoglioProposteEventi_().getDataRange().getValues();
  return values.slice(1).filter(function (row) {
    return String(row[7] || 'ATTIVA') === 'ATTIVA';
  }).map(function (row) { return preparaPropostaEvento_(row, socio); });
}

function getPropostaEventoPubblica(payload) {
  var id = typeof payload === 'object' ? String(payload.id || payload.propostaId || '') : String(payload || '');
  var values = getFoglioProposteEventi_().getDataRange().getValues();
  var row = values.find(function (item, index) {
    return index > 0 && String(item[0]) === id && String(item[7] || 'ATTIVA') === 'ATTIVA';
  });
  return row ? preparaPropostaEvento_(row, '') : null;
}

function creaPropostaEvento(email, password, proposta) {
  if (!password || typeof verificaLogin !== 'function' || verificaLogin({ email: email, password: password }) !== 'OK_LOGIN') {
    return { ok: false, messaggio: 'Accedi al portale per proporre un evento.' };
  }
  if (!verificaSocioEvento_(email)) return { ok: false, messaggio: 'Solo i soci attivi possono proporre eventi.' };
  proposta = proposta || {};

  var titolo = String(proposta.titolo || '').trim();
  var descrizione = String(proposta.descrizione || '').trim();
  var opzioni = normalizzaOpzioniEvento_(proposta.opzioni);
  if (!titolo) return { ok: false, messaggio: 'Inserisci il titolo dell’evento.' };
  if (!opzioni.length) return { ok: false, messaggio: 'Inserisci almeno una data e un orario per la proposta.' };
  if (opzioni.some(function (opzione) { return !dataOraEvento_(opzione); })) {
    return { ok: false, messaggio: 'Una data o un orario non è valido. Usa gg/mm/aaaa e hh:mm.' };
  }

  var id = Utilities.getUuid();
  var creataIl = new Date();
  getFoglioProposteEventi_().appendRow([
    id,
    creataIl,
    String(email).trim().toLowerCase(),
    titolo,
    descrizione,
    JSON.stringify(opzioni),
    JSON.stringify({}),
    'ATTIVA'
  ]);

  try { configuraTriggerPromemoriaProposteEvento_(); } catch (e) { console.log('Trigger promemoria eventi non configurato: ' + e); }
  try {
    inviaComunicazioniPropostaEvento_(
      sociAttiviProposteEvento_(),
      titolo,
      id,
      false
    );
  } catch (e) { console.log('Notifica nuova proposta evento non completata: ' + e); }

  return { ok: true, proposta: { id: id, titolo: titolo, descrizione: descrizione, opzioni: opzioni, voti: {} } };
}

function modificaPropostaEvento(dati) {
  dati = dati || {};
  var email = String(dati.email || '').trim().toLowerCase();
  if (!dati.password || typeof verificaLogin !== 'function' || verificaLogin({ email: email, password: dati.password }) !== 'OK_LOGIN') {
    return { ok: false, messaggio: 'Accedi al portale per modificare la proposta.' };
  }
  if (!verificaSocioEvento_(email)) return { ok: false, messaggio: 'Solo i soci attivi possono modificare le proposte.' };

  var proposta = dati.proposta || {};
  var titolo = String(proposta.titolo || '').trim();
  var descrizione = String(proposta.descrizione || '').trim();
  var opzioniNuove = normalizzaOpzioniEvento_(proposta.opzioni);
  if (!titolo || !opzioniNuove.length || opzioniNuove.some(function (opzione) { return !dataOraEvento_(opzione); })) {
    return { ok: false, messaggio: 'Inserisci titolo e almeno una data e un’ora valide.' };
  }
  if (opzioniNuove.length > 20) return { ok: false, messaggio: 'La proposta può contenere al massimo 20 opzioni.' };
  if (haSlotDuplicatiEvento_(opzioniNuove)) return { ok: false, messaggio: 'Non puoi inserire due volte la stessa data e ora.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var esito;
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) { return index > 0 && String(row[0]) === String(dati.propostaId); });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata.' };
    var row = values[rowIndex];
    if (String(row[2] || '').trim().toLowerCase() !== email) return { ok: false, messaggio: 'Solo il creatore può modificare questa proposta.' };
    if (String(row[7] || 'ATTIVA') !== 'ATTIVA') return { ok: false, messaggio: 'La proposta non è più modificabile.' };
    esito = applicaModificaPropostaEvento_(sheet, rowIndex, row, email, titolo, descrizione, opzioniNuove, null, false);
  } finally {
    lock.releaseLock();
  }

  inviaRichiestaRicompilazionePropostaEvento_(esito.contatti, titolo, dati.propostaId);
  if (esito.cambiate) inviaConfermaModificaCreatorePropostaEvento_(email, titolo, esito.contatti.length);
  return { ok: true, ricompilazioniRichieste: esito.contatti.length };
}

function inviaConfermaModificaCreatorePropostaEvento_(emailCreatore, titolo, numeroAvvisati) {
  var messaggio = numeroAvvisati
    ? 'Hai modificato date, orari o luoghi. Abbiamo avvisato ' + numeroAvvisati + (numeroAvvisati === 1 ? ' persona' : ' persone') + ' che avevano già risposto, chiedendo di ricompilare il questionario.'
    : 'Hai modificato date, orari o luoghi. Non c\'è nessun altro votante da avvisare.';
  try {
    MailApp.sendEmail({
      to: emailCreatore,
      subject: 'Modifica salvata: ' + titolo,
      body: messaggio + '\n\nGens Ssshhh',
      name: 'Gens Ssshhh'
    });
  } catch (e) {
    console.log('Invio conferma modifica al creatore non completato: ' + e);
  }
}

function applicaModificaPropostaEvento_(sheet, rowIndex, row, creatore, titolo, descrizione, opzioniNuove, chiaveAutore, rispostaFacoltativaNuovaOpzione) {
  var oldOptions = [];
  try { oldOptions = JSON.parse(row[5] || '[]'); } catch (e) { }
  var oldVoti = leggiVotiEvento_(row[6], oldOptions, '');
  var oldSummary = oldVoti.riepilogo;
  var cambiate = opzioniEventoCambiate_(oldOptions, opzioniNuove);

  var mapping = oldOptions.map(function (oldOption) {
    return opzioniNuove.findIndex(function (newOption) {
      return chiaveOpzioneEvento_(newOption) === chiaveOpzioneEvento_(oldOption);
    });
  });
  var counts = opzioniNuove.map(function () { return 0; });
  oldSummary.opzioni.forEach(function (count, oldIndex) {
    if (mapping[oldIndex] >= 0) counts[mapping[oldIndex]] += Number(count || 0);
  });

  var votanti = oldVoti.votanti;
  var votantiConPreferenzeRimosse = {};
  Object.keys(votanti).forEach(function (key) {
    var response = votanti[key];
    if (typeof response === 'number') response = { interesse: 'si', opzioni: [response] };
    if (!response || typeof response !== 'object') return;
    var oldChoices = Array.isArray(response.opzioni) ? response.opzioni : [];
    var preferenzaRimossa = oldChoices.some(function (oldIndex) {
      oldIndex = Number(oldIndex);
      return Number.isInteger(oldIndex) && oldIndex >= 0 && mapping[oldIndex] < 0;
    });
    if (preferenzaRimossa) votantiConPreferenzeRimosse[key] = true;
    response.opzioni = oldChoices.map(Number).filter(function (oldIndex) {
      return Number.isInteger(oldIndex) && mapping[oldIndex] >= 0;
    }).map(function (oldIndex) { return mapping[oldIndex]; });
    votanti[key] = response;
  });

  var ricompilazione = oldVoti.ricompilazione;
  var contatti = [];
  if (cambiate) {
    var richiestaIl = new Date().toISOString();
    Object.keys(votanti).forEach(function (chiave) {
      if (chiave === 'socio:' + creatore || chiave === creatore) return;
      if (rispostaFacoltativaNuovaOpzione) {
        if (ricompilazione[chiave] === undefined) ricompilazione[chiave] = 'FACOLTATIVA:' + richiestaIl;
        return;
      }
      if (!votantiConPreferenzeRimosse[chiave]) return;
      ricompilazione[chiave] = richiestaIl;
      if (chiave !== chiaveAutore) contatti.push(contattoVotanteEvento_(chiave, votanti[chiave]));
    });
  }

  sheet.getRange(rowIndex + 1, 4).setValue(titolo);
  sheet.getRange(rowIndex + 1, 5).setValue(descrizione);
  sheet.getRange(rowIndex + 1, 6).setValue(JSON.stringify(opzioniNuove));
  sheet.getRange(rowIndex + 1, 7).setValue(JSON.stringify(datiVotiDaSalvare_({ interesse: oldSummary.interesse, opzioni: counts }, votanti, ricompilazione)));
  return { contatti: contatti, cambiate: cambiate };
}

function annullaPropostaEvento(dati) {
  dati = dati || {};
  var email = String(dati.email || '').trim().toLowerCase();
  if (!dati.password || typeof verificaLogin !== 'function' || verificaLogin({ email: email, password: dati.password }) !== 'OK_LOGIN') {
    return { ok: false, messaggio: 'Accedi al portale per annullare la proposta.' };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) { return index > 0 && String(row[0]) === String(dati.propostaId); });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata.' };
    var row = values[rowIndex];
    if (String(row[2] || '').trim().toLowerCase() !== email) return { ok: false, messaggio: 'Solo il creatore può annullare questa proposta.' };
    if (String(row[7] || 'ATTIVA') !== 'ATTIVA') return { ok: false, messaggio: 'La proposta non è più aperta.' };
    sheet.getRange(rowIndex + 1, 8).setValue('ANNULLATA');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function identitaVotanteEvento_(email, dati) {
  var socio = String(email || '').trim().toLowerCase();
  if (verificaSocioEvento_(socio)) return { chiave: 'socio:' + socio, email: socio, nome: '', socio: true };

  var nome = String(dati.nome || '').trim();
  var emailOspite = String(dati.email || '').trim().toLowerCase();
  if (!nome || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailOspite)) {
    return { errore: 'Per votare come ospite inserisci nome ed e-mail validi.' };
  }
  if (verificaSocioEvento_(emailOspite)) {
    return { errore: 'Questa e-mail appartiene a un socio: accedi al portale per votare.' };
  }
  return { chiave: 'ospite:' + hashEmailEvento_(emailOspite), email: emailOspite, nome: nome, socio: false };
}

function votaPropostaEvento(email, propostaId, risposta) {
  var socio = String(email || '').trim().toLowerCase();
  var rispostaVoto = typeof risposta === 'number'
    ? { interesse: 'si', opzioni: [risposta] }
    : (risposta || {});
  var interessato = String(rispostaVoto.interesse || '').toLowerCase();
  var opzioniScelte = Array.isArray(rispostaVoto.opzioni) ? rispostaVoto.opzioni : [];
  var identita = identitaVotanteEvento_(socio, rispostaVoto);
  if (identita.errore) return { ok: false, messaggio: identita.errore };
  var chiaveVotante = identita.chiave;

  if (interessato !== 'si' && interessato !== 'no') {
    return { ok: false, messaggio: 'Indica se ti interessa partecipare.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) {
      return index > 0 && String(row[0]) === String(propostaId);
    });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata.' };

    var row = values[rowIndex];
    var opzioni = JSON.parse(row[5] || '[]');
    var scelte = interessato === 'si' ? opzioniScelte.map(Number) : [];
    if (scelte.some(function (indice) {
      return !Number.isInteger(indice) || indice < 0 || indice >= opzioni.length;
    })) {
      return { ok: false, messaggio: 'Opzione non valida.' };
    }
    scelte = scelte.filter(function (indice, index) { return scelte.indexOf(indice) === index; });
    if (interessato === 'si' && !scelte.length) {
      return { ok: false, messaggio: 'Se scegli Sì, seleziona almeno una data e un orario preferiti.' };
    }

    var votiEsistenti = leggiVotiEvento_(row[6], opzioni, socio);
    var votanti = votiEsistenti.votanti;
    var ricompilazione = votiEsistenti.ricompilazione;
    var chiaveEsistente = votanti[chiaveVotante] !== undefined ? chiaveVotante : (socio && votanti[socio] !== undefined ? socio : null);
    if (chiaveEsistente !== null) {
      if (ricompilazione[chiaveEsistente] === undefined && !identita.socio) {
        return { ok: false, messaggio: 'Hai già votato questa proposta.' };
      }
      rimuoviContributoVotoEvento_(votiEsistenti.riepilogo, votanti, chiaveEsistente);
      delete ricompilazione[chiaveEsistente];
    }

    var interesse = votiEsistenti.riepilogo.interesse;
    interesse[interessato]++;
    scelte.forEach(function (indice) { votiEsistenti.riepilogo.opzioni[indice]++; });
    votanti[chiaveVotante] = { interesse: interessato, opzioni: scelte, email: identita.email, nome: identita.nome };

    sheet.getRange(rowIndex + 1, 7).setValue(JSON.stringify(datiVotiDaSalvare_(votiEsistenti.riepilogo, votanti, ricompilazione)));
    return { ok: true, voti: votiEsistenti.riepilogo };
  } finally {
    lock.releaseLock();
  }
}

function ricompilaVotoPropostaEvento(email, propostaId, ricompila) {
  var socio = String(email || '').trim().toLowerCase();
  if (!verificaSocioEvento_(socio)) return { ok: false, messaggio: 'Solo i soci attivi possono rispondere a questa richiesta.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) {
      return index > 0 && String(row[0]) === String(propostaId) && String(row[7] || 'ATTIVA') === 'ATTIVA';
    });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata o già chiusa.' };

    var row = values[rowIndex];
    var opzioni = [];
    try { opzioni = JSON.parse(row[5] || '[]'); } catch (e) { }
    var voti = leggiVotiEvento_(row[6], opzioni, socio);
    var chiaveRichiesta = chiaveRicompilazioneSocio_(voti.ricompilazione, socio);
    if (chiaveRichiesta === null) return { ok: false, messaggio: 'Non risulta una richiesta di ricompilazione per questo evento.' };

    var vuoleRicompilare = ricompila === true || ricompila === 'true';
    delete voti.ricompilazione[chiaveRichiesta];
    if (vuoleRicompilare) rimuoviContributoVotoEvento_(voti.riepilogo, voti.votanti, chiaveRichiesta);

    sheet.getRange(rowIndex + 1, 7).setValue(JSON.stringify(datiVotiDaSalvare_(voti.riepilogo, voti.votanti, voti.ricompilazione)));
    return { ok: true, ricompilato: vuoleRicompilare };
  } finally {
    lock.releaseLock();
  }
}

function aggiungiOpzionePropostaEvento(email, propostaId, opzione) {
  opzione = opzione || {};
  var votante = identitaVotanteEvento_(email, opzione);
  if (votante.errore) return { ok: false, messaggio: votante.errore };
  var nuova = normalizzaOpzioniEvento_([{ data: opzione.data, ora: opzione.ora, luogo: opzione.luogo }]);
  if (!nuova.length || !dataOraEvento_(nuova[0])) return { ok: false, messaggio: 'Inserisci una data e un orario validi.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var esito;
  var titolo = '';
  var creatore = '';
  var indice = 0;
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) {
      return index > 0 && String(row[0]) === String(propostaId) && String(row[7] || 'ATTIVA') === 'ATTIVA';
    });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata o già confermata.' };
    var row = values[rowIndex];
    titolo = String(row[3] || '');
    creatore = String(row[2] || '').trim().toLowerCase();

    var options = [];
    try { options = JSON.parse(row[5] || '[]'); } catch (e) { }
    if (options.length >= 20) return { ok: false, messaggio: 'La proposta ha già raggiunto il limite di 20 opzioni.' };
    if (options.some(function (existing) { return slotOpzioneEvento_(existing) === slotOpzioneEvento_(nuova[0]); })) {
      return { ok: false, messaggio: 'Questa data e ora sono già state proposte.' };
    }

    indice = options.length;
    esito = applicaModificaPropostaEvento_(sheet, rowIndex, row, creatore, titolo, String(row[4] || ''), options.concat(nuova), votante.chiave, true);
  } finally {
    lock.releaseLock();
  }

  inviaRichiestaRicompilazionePropostaEvento_(esito.contatti, titolo, propostaId);
  if (votante.chiave !== 'socio:' + creatore) {
    var nomeAutore = votante.nome || nomiSociEvento_()[votante.email] || votante.email;
    inviaAggiuntaOpzioneAlCreatorePropostaEvento_(creatore, titolo, nomeAutore, nuova[0]);
  }
  return { ok: true, indice: indice, opzione: nuova[0], ricompilazioniRichieste: esito.contatti.length };
}

function inviaAggiuntaOpzioneAlCreatorePropostaEvento_(emailCreatore, titolo, nomeAutore, opzione) {
  var luogo = opzione.luogo ? ' · ' + opzione.luogo : '';
  try {
    MailApp.sendEmail({
      to: emailCreatore,
      subject: 'Nuova data proposta: ' + titolo,
      body: nomeAutore + ' ha proposto una nuova data per "' + titolo + '": ' + opzione.data + ' alle ' + opzione.ora + luogo + '.\n\nLe risposte già inviate restano valide; chi lo desidera può aggiornare le proprie preferenze dal link della proposta.\n\nGens Ssshhh',
      name: 'Gens Ssshhh'
    });
  } catch (e) {
    console.log('Invio avviso nuova opzione al creatore non completato: ' + e);
  }
}

function getRispostaRicompilazioneOspite(dati) {
  dati = dati || {};
  var id = String(dati.id || dati.propostaId || '');
  var emailOspite = String(dati.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailOspite) || verificaSocioEvento_(emailOspite)) return null;

  var chiave = 'ospite:' + hashEmailEvento_(emailOspite);
  var values = getFoglioProposteEventi_().getDataRange().getValues();
  var row = values.find(function (item, index) {
    return index > 0 && String(item[0]) === id && String(item[7] || 'ATTIVA') === 'ATTIVA';
  });
  if (!row) return null;

  var opzioni = [];
  try { opzioni = JSON.parse(row[5] || '[]'); } catch (e) { }
  var voti = leggiVotiEvento_(row[6], opzioni, '');
  var risposta = voti.votanti[chiave];
  if (voti.ricompilazione[chiave] === undefined || !risposta || typeof risposta !== 'object') return null;

  var indici = (Array.isArray(risposta.opzioni) ? risposta.opzioni : []).map(Number).filter(function (indice) {
    return Number.isInteger(indice) && indice >= 0 && indice < opzioni.length;
  });
  return {
    nome: String(risposta.nome || ''),
    interesse: risposta.interesse === 'no' ? 'no' : 'si',
    opzioni: indici,
    facoltativa: String(voti.ricompilazione[chiave]).indexOf('FACOLTATIVA:') === 0
  };
}

function confermaEventoProposto(dati) {
  dati = dati || {};
  var email = String(dati.email || '').trim().toLowerCase();
  if (!dati.password || typeof verificaLogin !== 'function' || verificaLogin({ email: email, password: dati.password }) !== 'OK_LOGIN') {
    return { ok: false, messaggio: 'Accedi al portale per confermare l’evento.' };
  }
  if (!verificaSocioEvento_(email)) return { ok: false, messaggio: 'Solo i soci attivi possono confermare l’evento.' };
  var propostaId = dati.propostaId;
  var opzioneFinale = dati.opzione || {};
  var opzioniFinali = normalizzaOpzioniEvento_([opzioneFinale]);
  if (!opzioniFinali.length || !dataOraEvento_(opzioniFinali[0])) return { ok: false, messaggio: 'La data o l’orario dell’evento non sono validi.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getFoglioProposteEventi_();
    var values = sheet.getDataRange().getValues();
    var rowIndex = values.findIndex(function (row, index) {
      return index > 0 && String(row[0]) === String(propostaId);
    });
    if (rowIndex < 1) return { ok: false, messaggio: 'Proposta non trovata.' };
    var row = values[rowIndex];
    if (String(row[2] || '').trim().toLowerCase() !== email) return { ok: false, messaggio: 'Solo il creatore può confermare l’evento.' };
    if (String(row[7] || 'ATTIVA') !== 'ATTIVA') return { ok: false, messaggio: 'Questa proposta è già stata chiusa.' };

    var votes = {};
    try { votes = JSON.parse(row[6] || '{}'); } catch (e) { }
    var confirmed = votes._votanti || {};
    var esclusi = Array.isArray(dati.esclusi) ? dati.esclusi.map(String) : [];
    var soloDataSelezionata = dati.soloDataSelezionata === true;
    var indiciDataSelezionata = [];
    if (soloDataSelezionata) {
      var opzioniProposta = [];
      try { opzioniProposta = JSON.parse(row[5] || '[]'); } catch (e) { }
      var slotScelto = slotOpzioneEvento_(opzioneFinale);
      opzioniProposta.forEach(function (opzioneProposta, indice) {
        if (slotOpzioneEvento_({
          data: String(opzioneProposta.data || '').trim(),
          ora: String(opzioneProposta.ora || '').trim()
        }) === slotScelto) indiciDataSelezionata.push(indice);
      });
    }
    var attendees = [];
    votantiSiEvento_(confirmed, soloDataSelezionata, indiciDataSelezionata).forEach(function (key) {
      if (esclusi.indexOf(key) >= 0) return;
      var address = indirizzoVotanteEvento_(key, confirmed[key]);
      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address) && attendees.indexOf(address) < 0) attendees.push(address);
    });
    if (!attendees.length) {
      return {
        ok: false,
        messaggio: soloDataSelezionata
          ? 'Nessun partecipante ha risposto Sì a questa data e ora.'
          : 'Nessun partecipante ha risposto Sì.'
      };
    }

    var start = dataOraEvento_(opzioniFinali[0]);
    if (start.getTime() <= new Date().getTime()) return { ok: false, messaggio: 'Scegli una data e un orario futuri.' };
    var end = new Date(start.getTime() + 2 * 60 * 60 * 1000);
    var calendar = CalendarApp.getCalendarById(EVENTI_CALENDAR_ID);
    if (!calendar) return { ok: false, messaggio: 'Il calendario associativo non è accessibile all’account Apps Script.' };
    var event = calendar.createEvent(String(row[3] || 'Evento associativo'), start, end, {
      description: String(row[4] || ''),
      location: opzioniFinali[0].luogo,
      guests: attendees.join(','),
      sendInvites: true
    });
    if (event.setGuestsCanSeeGuests) event.setGuestsCanSeeGuests(false);

    sheet.getRange(rowIndex + 1, 8).setValue('CONFERMATA');
    sheet.getRange(rowIndex + 1, 9).setValue(event.getId());
    sheet.getRange(rowIndex + 1, 10).setValue(JSON.stringify(opzioniFinali[0]));
    return { ok: true, invitati: attendees.length, idEvento: event.getId(), opzione: opzioniFinali[0] };
  } finally {
    lock.releaseLock();
  }
}

function sociAttiviProposteEvento_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('soci');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  var headers = rows[0].map(function (value) { return String(value).trim().toLowerCase(); });
  var emailIndex = headers.indexOf('email');
  if (emailIndex < 0) emailIndex = 2;
  var statusIndex = headers.indexOf('stato socio');
  return rows.slice(1).filter(function (row) {
    var email = String(row[emailIndex] || '').trim();
    var attivo = statusIndex < 0 || String(row[statusIndex] || '').trim().toLowerCase() === 'attivo';
    return attivo && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }).map(function (row) { return String(row[emailIndex]).trim().toLowerCase(); }).filter(function (email, index, emails) {
    return emails.indexOf(email) === index;
  });
}

function escapeHtmlEvento_(valore) {
  return String(valore == null ? '' : valore).replace(/[&<>"']/g, function (carattere) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[carattere];
  });
}

function inviaComunicazioniPropostaEvento_(emails, titolo, propostaId, promemoria) {
  if (!emails.length) return;
  var oggetto = promemoria ? 'Promemoria: questionario evento' : 'Nuovo questionario evento';
  var testoEvento = escapeHtmlEvento_(titolo);
  var url = 'https://gensssshhh-collab.github.io/Sito-web-Gens-Ssshhh/index.html';
  var intro = promemoria
    ? 'Sono passati tre giorni e non risulta ancora una tua risposta.'
    : 'È stato pubblicato un nuovo questionario per un evento.';
  var body = intro + '\n\n' + titolo + '\n\nAccedi al portale soci per rispondere: ' + url;
  var htmlBody = '<div style="font-family:Arial,sans-serif;color:#243247;line-height:1.5">' +
    '<p>' + intro + '</p><h2>' + testoEvento + '</h2><p><a href="' + url + '">Accedi al portale soci</a></p>' +
    '<p>Gens Ssshhh</p></div>';

  try {
    for (var inizio = 0; inizio < emails.length; inizio += 40) {
      var gruppo = emails.slice(inizio, inizio + 40);
      MailApp.sendEmail({
        to: 'undisclosed-recipients@gmail.com',
        bcc: gruppo.join(','),
        subject: oggetto + ': ' + titolo,
        body: body,
        htmlBody: htmlBody,
        name: 'Gens Ssshhh'
      });
    }
  } catch (e) {
    console.log('Invio email questionario evento non completato: ' + e);
  }

  try {
    if (typeof inviaNotificaPush === 'function') {
      inviaNotificaPush(oggetto, 'Partecipa al questionario: ' + titolo, emails, 'viewCalendario');
    }
  } catch (e) {
    console.log('Invio push questionario evento non completato: ' + e);
  }
}

function inviaEmailRicompilazioneEvento_(destinatari, titolo, intro, url, etichettaLink) {
  var oggetto = 'Ricompila il questionario evento';
  var testoEvento = escapeHtmlEvento_(titolo);
  var body = intro + '\n\n' + titolo + '\n\n' + etichettaLink + ': ' + url;
  var htmlBody = '<div style="font-family:Arial,sans-serif;color:#243247;line-height:1.5">' +
    '<p>' + intro + '</p><h2>' + testoEvento + '</h2><p><a href="' + url + '">' + etichettaLink + '</a></p>' +
    '<p>Gens Ssshhh</p></div>';

  try {
    for (var inizio = 0; inizio < destinatari.length; inizio += 40) {
      MailApp.sendEmail({
        to: 'undisclosed-recipients@gmail.com',
        bcc: destinatari.slice(inizio, inizio + 40).join(','),
        subject: oggetto + ': ' + titolo,
        body: body,
        htmlBody: htmlBody,
        name: 'Gens Ssshhh'
      });
    }
  } catch (e) {
    console.log('Invio email ricompilazione evento non completato: ' + e);
  }
}

function inviaRichiestaRicompilazionePropostaEvento_(contatti, titolo, propostaId) {
  if (!contatti || !contatti.length) return;
  var emailSocie = [];
  var emailEsterni = [];
  contatti.forEach(function (contatto) {
    if (!contatto || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contatto.email)) return;
    var elenco = contatto.socio ? emailSocie : emailEsterni;
    if (elenco.indexOf(contatto.email) < 0) elenco.push(contatto.email);
  });
  if (!emailSocie.length && !emailEsterni.length) return;

  var introModifica = 'Il creatore ha modificato date, orari o luoghi di questo evento. Ti chiediamo di confermare di nuovo la tua disponibilità.';
  if (emailSocie.length) {
    inviaEmailRicompilazioneEvento_(emailSocie, titolo,
      introModifica + ' Accedi al portale soci per aggiornare la tua risposta.',
      'https://gensssshhh-collab.github.io/Sito-web-Gens-Ssshhh/index.html',
      'Accedi al portale soci');
  }
  if (emailEsterni.length) {
    inviaEmailRicompilazioneEvento_(emailEsterni, titolo, introModifica,
      'https://gensssshhh-collab.github.io/Sito-web-Gens-Ssshhh/index.html?evento=' + encodeURIComponent(propostaId),
      'Apri il questionario evento');
  }

  try {
    if (emailSocie.length && typeof inviaNotificaPush === 'function') {
      inviaNotificaPush('Ricompila il questionario evento', 'Ricompila la tua risposta per: ' + titolo, emailSocie, 'viewCalendario');
    }
  } catch (e) {
    console.log('Invio push ricompilazione evento non completato: ' + e);
  }
}

function configuraTriggerPromemoriaProposteEvento_() {
  var presente = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'inviaPromemoriaProposteEvento';
  });
  if (!presente) ScriptApp.newTrigger('inviaPromemoriaProposteEvento').timeBased().everyDays(1).atHour(9).create();
}

function inviaPromemoriaProposteEvento() {
  var sheet = getFoglioProposteEventi_();
  var rows = sheet.getDataRange().getValues();
  var sociAttivi = sociAttiviProposteEvento_();
  var adesso = new Date();
  var promemoriaInviati = 0;

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (String(row[7] || 'ATTIVA') !== 'ATTIVA' || row[10]) continue;
    var creataIl = row[1] instanceof Date ? row[1] : new Date(row[1]);
    if (isNaN(creataIl.getTime()) || adesso.getTime() - creataIl.getTime() < 3 * 86400000) continue;

    var opzioni = [];
    var votanti = {};
    try { opzioni = JSON.parse(row[5] || '[]'); } catch (e) { }
    votanti = leggiVotiEvento_(row[6], opzioni, '').votanti;
    var senzaRisposta = sociAttivi.filter(function (email) {
      return votanti['socio:' + email] === undefined && votanti[email] === undefined;
    });
    inviaComunicazioniPropostaEvento_(senzaRisposta, String(row[3] || 'Evento proposto'), String(row[0]), true);
    sheet.getRange(i + 1, 11).setValue(adesso);
    promemoriaInviati++;
  }
  return promemoriaInviati;
}