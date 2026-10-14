const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function loadScript(file, sandbox) {
  const source = fs.readFileSync(path.join(root, file), 'utf8');
  vm.runInNewContext(source, sandbox, { filename: file });
  return sandbox;
}

function dataSheet(values) {
  return {
    getDataRange() {
      return { getValues: () => values };
    },
    appendRow(row) {
      values.push(row.slice());
    },
    getRange() {
      return {
        setValue() { }
      };
    }
  };
}

test('login: accetta credenziali corrette e blocca dopo cinque errori', () => {
  const cacheValues = new Map();
  const cache = {
    get: key => cacheValues.get(key) || null,
    put: (key, value) => cacheValues.set(key, value),
    remove: key => cacheValues.delete(key)
  };
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email'],
    ['Mario', 'hash:segreta', 'Mario@Example.it']
  ]);
  const logRows = [['Data', 'Email', 'Azione', 'Info']];
  const logSheet = dataSheet(logRows);
  const spreadsheet = {
    getSheetByName: name => name === 'Log' ? logSheet : sociSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    CacheService: { getScriptCache: () => cache },
    Utilities: { base64Encode: value => Buffer.from(value).toString('base64') },
    creaHash: value => 'hash:' + value
  };

  loadScript('Auth.js', context);

  assert.equal(context.verificaLogin({ email: ' mario@example.it ', password: 'segreta', info: 'test' }), 'OK_LOGIN');
  assert.equal(context.verificaLogin({ email: 'mario@example.it', password: 'errata' }), 'ERR_CREDENZIALI');
  assert.equal(cacheValues.size, 1);
  cacheValues.set('block_login_bWFyaW9AZXhhbXBsZS5pdA==', '5');
  assert.equal(context.verificaLogin({ email: 'mario@example.it', password: 'segreta' }), 'BLOCKED');
  assert.ok(logRows.some(row => row[2] === 'LOGIN_OK'));
});

test('account ammissione: crea socio non attivo e invia link per impostare password', () => {
  const headers = ['Nome', 'Password', 'Email', 'ID Cartella Personale', 'Cognome', 'Telefono', 'Indirizzo', 'ResetToken', 'ResetTime', 'Numero Tessera', 'Data Scadenza', 'Stato Socio', 'Carica sociale', 'Sesso', 'Codice Fiscale', 'Data Nascita', 'Comune Nascita', 'Versione Accettata', 'ID_Univoco', 'CAP res', 'Comune res'];
  const rows = [headers, ['Mario', 'old-hash', 'mario@example.it']];
  const sociSheet = dataSheet(rows);
  sociSheet.getRange = (row, column) => ({
    setValue(value) { rows[row - 1][column - 1] = value; },
    getValue() { return rows[row - 1][column - 1]; }
  });
  const messages = [];
  let id = 0;
  const logSheet = dataSheet([['Data', 'Email', 'Azione', 'Info']]);
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'Log' ? logSheet : null }) },
    Utilities: { getUuid: () => 'uuid-' + (++id) },
    MailApp: { sendEmail: message => messages.push(message) },
    creaHash: value => 'hash:' + value,
    scriviLog() { }
  };
  loadScript('Auth.js', context);

  assert.equal(context.creaAccountSocioDaAmmissione_({ email: 'giulia@example.net', nome: 'Giulia', cognome: 'Neri', telefono: '123' }), 'LINK_INVIATO');
  const newMember = rows[2];
  assert.equal(newMember[0], 'Giulia');
  assert.equal(newMember[1], 'hash:uuid-1');
  assert.equal(newMember[2], 'giulia@example.net');
  assert.equal(newMember[7], 'uuid-3');
  assert.equal(newMember[11], 'non attivo');
  assert.equal(newMember[18], 'uuid-2');
  assert.match(messages[0].subject, /Attiva il tuo account/);
  assert.equal(context.creaAccountSocioDaAmmissione_({ email: 'GIULIA@example.net', nome: 'Giulia', cognome: 'Neri' }), 'ACCOUNT_ESISTENTE');
  assert.equal(rows.length, 3);
});

test('ammissione: crea account solo dopo esito favorevole dei 2/3', () => {
  const headersSoci = ['Nome', 'Password', 'Email', 'ID Cartella Personale', 'Cognome', 'Telefono', 'Indirizzo', 'ResetToken', 'ResetTime', 'Numero Tessera', 'Data Scadenza', 'Stato Socio', 'Carica sociale', 'Sesso', 'Codice Fiscale', 'Data Nascita', 'Comune Nascita', 'Versione Accettata', 'ID_Univoco', 'CAP res', 'Comune res'];
  const sociRows = [headersSoci, ['Luca', 'hash:admin', 'admin@example.it', '', 'Germandi', '', '', '', '', '1', '', 'ATTIVO', 'Presidente', '', '', '', '', '', 'uuid-admin']];
  const ammissioniRows = [
    ['Data', 'Nome', 'Cognome', 'Email', 'Telefono', 'Sponsor 1', 'Sponsor 2', 'Stato', 'Esito', 'ID Richiesta', 'Presentazione', 'Scadenza Avvalli', 'Data Voto Prevista', 'Voto Notificato Il'],
    [new Date(), 'Giulia', 'Neri', 'giulia@example.net', '123', '', '', 'IN VOTAZIONE', 'DA VOTARE', 'request-1', 'Presentazione', new Date(), new Date(), '']
  ];
  const votiRows = [
    ['Timestamp', 'HashSocio', 'CandidatoEmail', 'Voto'],
    [new Date(), 'hash-1', 'giulia@example.net', 'FAVOREVOLE'],
    [new Date(), 'hash-2', 'giulia@example.net', 'FAVOREVOLE'],
    [new Date(), 'hash-3', 'giulia@example.net', 'CONTRARIO']
  ];
  const sociSheet = dataSheet(sociRows);
  const ammissioniSheet = dataSheet(ammissioniRows);
  const votiSheet = dataSheet(votiRows);
  [sociSheet, ammissioniSheet].forEach(sheet => {
    sheet.getRange = (row, column) => ({
      getValue: () => (sheet === sociSheet ? sociRows : ammissioniRows)[row - 1][column - 1],
      setValue(value) { (sheet === sociSheet ? sociRows : ammissioniRows)[row - 1][column - 1] = value; }
    });
  });
  const sheets = { soci: sociSheet, Ammissioni: ammissioniSheet, VotiAmmissioni: votiSheet, Log: null };
  const messages = [];
  let uuid = 0;
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] || null }) },
    Utilities: { getUuid: () => 'new-uuid-' + (++uuid) },
    MailApp: { sendEmail: message => messages.push(message) },
    creaHash: value => 'hash:' + value,
    scriviLog() { },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) }
  };
  loadScript('Auth.js', context);
  loadScript('user.js', context);
  loadScript('Admin.js', context);

  const result = context.adminGestisciAmmissione('admin@example.it', 'giulia@example.net', 'CHIUDI_VOTO');
  assert.equal(result, 'OK_SPOGLIO|AMMESSO|2|3|2|LINK_INVIATO');
  assert.equal(sociRows.length, 3);
  assert.equal(sociRows[2][2], 'giulia@example.net');
  assert.equal(sociRows[2][11], 'non attivo');
  assert.ok(sociRows[2][1]);
  assert.ok(sociRows[2][7]);
  assert.match(messages[0].subject, /Attiva il tuo account/);
  assert.equal(ammissioniRows[1][8], 'AMMESSO (2/3) | ACCOUNT: LINK_INVIATO');

  ammissioniRows.push([new Date(), 'Paolo', 'Blu', 'paolo@example.net', '', '', '', 'IN VOTAZIONE', 'DA VOTARE', 'request-2', '', new Date(), new Date(), '']);
  votiRows.push(
    [new Date(), 'hash-4', 'paolo@example.net', 'FAVOREVOLE'],
    [new Date(), 'hash-5', 'paolo@example.net', 'CONTRARIO'],
    [new Date(), 'hash-6', 'paolo@example.net', 'CONTRARIO']
  );
  const rejected = context.adminGestisciAmmissione('admin@example.it', 'paolo@example.net', 'CHIUDI_VOTO');
  assert.equal(rejected, 'OK_SPOGLIO|RESPINTO|1|3|2|');
  assert.equal(sociRows.length, 3);
  assert.equal(messages.length, 1);
});

test('fatture: admin cerca soci attivi e invia XML/PDF con sollecito pagamento', () => {
  const rows = [
    ['Nome', 'Password', 'Email', '', 'Cognome', '', '', '', '', '', '', 'Stato Socio', 'Carica sociale'],
    ['Mario', 'hash', 'mario@example.it', '', 'Rossi', '', '', '', '', '', '', 'ATTIVO'],
    ['Anna', 'hash', 'anna@example.it', '', 'Bianchi', '', '', '', '', '', '', 'attivo'],
    ['Paolo', 'hash', 'paolo@example.it', '', 'Blu', '', '', '', '', '', '', 'non attivo'],
    ['Admin', 'hash', 'admin@example.it', '', 'Zed', '', '', '', '', '', '', 'ATTIVO', 'Tesoriere']
  ];
  const sociSheet = {
    getLastRow: () => rows.length,
    getDataRange: () => ({ getValues: () => rows })
  };
  let sent = null;
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : null }) },
    isRuoloAmministrativo: ruolo => String(ruolo).toLowerCase() === 'tesoriere',
    verificaLogin: dati => dati.password === 'pw' ? 'OK_LOGIN' : 'ERR_AUTH',
    MailApp: { sendEmail: message => { sent = message; } }
  };
  loadScript('Finance.js', context);

  assert.deepEqual(JSON.parse(JSON.stringify(context.getSociPerFattura('admin@example.it', 'pw'))).map(socio => socio.nomeCompleto), ['Anna Bianchi', 'Mario Rossi', 'Admin Zed']);
  assert.equal(context.getSociPerFattura('admin@example.it', 'bad').length, 0);
  assert.equal(context.getSociPerFattura('mario@example.it', 'pw').length, 0);

  context.inviaFatturaSocio_({ email: 'mario@example.it', numero: '12/2026', importo: '45', causale: 'Quota associativa' }, 'Mario Rossi', { name: 'Fattura.pdf' }, { name: 'Fattura.xml' });
  assert.equal(sent.to, 'mario@example.it');
  assert.equal(sent.attachments.length, 2);
  assert.match(sent.body, /45\.00/);
  assert.match(sent.body, /Quota associativa/);
  assert.match(sent.body, /saldare entro la data indicata/);
});

test('fatture: XML e PDF riportano regime, bollo, scadenza e pagamento', () => {
  const sociRows = [
    ['Nome', 'Password', 'Email', '', 'Cognome', '', 'Indirizzo', '', '', 'Numero Tessera', 'Data Scadenza', 'Stato Socio', 'Carica sociale', '', 'Codice Fiscale', '', '', '', 'ID_Univoco', 'CAP res', 'Comune res'],
    ['Mario', 'hash', 'mario@example.it', '', 'Rossi', '', 'Via Roma 1', '', '', 12, '', 'ATTIVO', 'Socio semplice', '', 'RSSMRA80A01F205X', '', '', '', 'uuid-mario', '40100', 'Bologna'],
    ['Admin', 'hash', 'admin@example.it', '', 'Admin', '', '', '', '', '', '', 'ATTIVO', 'Tesoriere']
  ];
  const sociSheet = { getDataRange: () => ({ getValues: () => sociRows }) };
  const invoiceRows = [];
  const invoiceSheet = { appendRow: row => invoiceRows.push(row) };
  const configSheet = { getRange: () => ({ getValue: () => 'segreteria@example.it' }) };
  const archiveSpreadsheet = { getSheetByName: name => name === 'Fatture Emesse' ? invoiceSheet : null };
  const sent = [];
  let blobId = 0;
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'Config' ? configSheet : null }),
      openById: () => archiveSpreadsheet
    },
    getDatiUtente: email => email === 'admin@example.it' ? { ruolo: 'Tesoriere' } : null,
    isRuoloAmministrativo: ruolo => String(ruolo).toLowerCase() === 'tesoriere',
    verificaLogin: data => data.password === 'pw' ? 'OK_LOGIN' : 'ERR_AUTH',
    DriveApp: { getFolderById: () => ({ createFile: blob => ({ getUrl: () => 'https://drive.example/' + blob.name }) }) },
    MailApp: { sendEmail: mail => sent.push(mail) },
    MimeType: { HTML: 'text/html', PDF: 'application/pdf' },
    Utilities: {
      formatDate: (date, zone, format) => {
        const parsedDate = date instanceof Date ? date : new Date(date);
        const year = parsedDate.getFullYear();
        const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
        const day = String(parsedDate.getDate()).padStart(2, '0');
        if (format === 'yyyy-MM-dd') return `${year}-${month}-${day}`;
        if (format === 'dd/MM/yyyy') return `${day}/${month}/${year}`;
        if (format === 'yyyy') return String(year);
        return `${day}/${month}/${year} 12:00:00`;
      },
      newBlob: (content, type, name) => ({
        content: String(content), type, name: name || 'generated',
        setName(value) { this.name = value; },
        getAs: mime => ({ content: String(content), type: mime, name: 'pdf-' + (++blobId), setName(value) { this.name = value; }, getBytes() { return Buffer.from(String(content)); } }),
        getBytes() { return Buffer.from(String(content)); }
      }),
      base64Encode: bytes => Buffer.from(bytes).toString('base64')
    }
  };
  loadScript('Finance.js', context);

  const fattura = context.generaFatturaXML({
    adminEmail: 'admin@example.it', password: 'pw', tipoCliente: 'socio', email: 'mario@example.it',
    numero: '1', importo: '100.00', causale: 'Quota associativa annuale 2026', regimeFiscale: 'RF19',
    scadenzaPagamento: '2026-10-24', metodoPagamento: 'BONIFICO', datiPagamento: 'IT60X0542811101000000123456'
  });
  const xml = Buffer.from(fattura.xmlBase64.split(',')[1], 'base64').toString();
  const pdfHtml = Buffer.from(fattura.pdfBase64.split(',')[1], 'base64').toString();
  assert.equal(fattura.emailInviata, true);
  assert.match(xml, /<Numero>1\/2026<\/Numero>/);
  assert.match(xml, /<RegimeFiscale>RF19<\/RegimeFiscale>/);
  assert.match(xml, /<Natura>N2\.2<\/Natura>/);
  assert.match(xml, /Legge 190\/2014/);
  assert.match(xml, /<ImportoBollo>2\.00<\/ImportoBollo>/);
  assert.match(xml, /<ImportoTotaleDocumento>102\.00<\/ImportoTotaleDocumento>/);
  assert.match(xml, /<ModalitaPagamento>MP05<\/ModalitaPagamento>/);
  assert.match(xml, /<DataScadenzaPagamento>2026-10-24<\/DataScadenzaPagamento>/);
  assert.match(xml, /<ImportoPagamento>102\.00<\/ImportoPagamento>/);
  assert.match(xml, /<IBAN>IT60X0542811101000000123456<\/IBAN>/);
  assert.match(pdfHtml, /Imposta di bollo di 2,00 € assolta sull'originale informatico/);
  assert.match(pdfHtml, /Scadenza:<\/strong> 24\/10\/2026/);
  assert.match(pdfHtml, /IT60X0542811101000000123456/);
  assert.match(pdfHtml, /segreteria@example\.it/);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].attachments.length, 2);
  assert.match(sent[0].body, /Totale da pagare: € 102\.00/);
  assert.equal(invoiceRows.length, 1);

  const senzaRiferimento = context.generaFatturaXML({
    adminEmail: 'admin@example.it', password: 'pw', tipoCliente: 'socio', email: 'mario@example.it',
    numero: '2/2026', importo: '50', causale: 'Quota associativa annuale 2026', regimeFiscale: 'RF01',
    scadenzaPagamento: '2026-10-24', metodoPagamento: 'BONIFICO', datiPagamento: 'IT60X0542811101000000123456'
  });
  assert.match(senzaRiferimento.errore, /riferimento normativo/);

  const datiBase = {
    adminEmail: 'admin@example.it', password: 'pw', tipoCliente: 'socio', email: 'mario@example.it',
    causale: 'Quota associativa annuale 2026', regimeFiscale: 'RF01', riferimentoIVA: 'Articolo 4 & normativa',
    scadenzaPagamento: '2026-10-24', metodoPagamento: 'BONIFICO', datiPagamento: 'IT60X0542811101000000123456'
  };
  const fatturaSoglia = context.generaFatturaXML({ ...datiBase, numero: '3/2026', importo: '77.47' });
  const xmlSoglia = Buffer.from(fatturaSoglia.xmlBase64.split(',')[1], 'base64').toString();
  assert.match(xmlSoglia, /<Natura>N4<\/Natura>/);
  assert.match(xmlSoglia, /Articolo 4 &amp; normativa/);
  assert.doesNotMatch(xmlSoglia, /<DatiBollo>/);
  assert.match(xmlSoglia, /<ImportoTotaleDocumento>77\.47<\/ImportoTotaleDocumento>/);

  const fatturaOltreSoglia = context.generaFatturaXML({ ...datiBase, numero: '4/2026', importo: '77.48' });
  const xmlOltreSoglia = Buffer.from(fatturaOltreSoglia.xmlBase64.split(',')[1], 'base64').toString();
  assert.match(xmlOltreSoglia, /<DatiBollo><BolloVirtuale>SI<\/BolloVirtuale><ImportoBollo>2\.00<\/ImportoBollo><\/DatiBollo>/);
  assert.match(xmlOltreSoglia, /<ImportoTotaleDocumento>79\.48<\/ImportoTotaleDocumento>/);
});

test('notifiche: registra un token una sola volta e invia il canale Android', () => {
  const properties = new Map();
  const requests = [];
  const context = {
    getDatiUtente: () => ({ email: 'mario@example.it' }),
    verificaLogin: dati => dati.password === 'secret' ? 'OK_LOGIN' : 'ERR_CREDENZIALI',
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => properties.get(key) || null,
        setProperty: (key, value) => properties.set(key, value)
      })
    },
    ScriptApp: { getOAuthToken: () => 'oauth-token' },
    UrlFetchApp: {
      fetch: (url, options) => {
        requests.push({ url, options });
        return { getResponseCode: () => 200 };
      }
    }
  };

  loadScript('PushNotifications.js', context);

  const token = { email: 'Mario@Example.it', token: 'token-1', piattaforma: 'android' };
  assert.equal(context.registraTokenPush(token), 'TOKEN_REGISTRATO');
  assert.equal(context.registraTokenPush(token), 'TOKEN_REGISTRATO');
  assert.equal(context.registraTokenPush({ email: 'anna@example.it', token: 'token-2', piattaforma: 'android' }), 'TOKEN_REGISTRATO');
  assert.equal(context.inviaNotificaPush('Titolo', 'Testo'), 'PUSH_INVIATA');
  assert.equal(context.inviaNotificaPushUtente('MARIO@EXAMPLE.IT', 'Firma', 'Documento da firmare', 'viewFirma'), 'PUSH_INVIATA');
  assert.equal(context.inviaNotificaPush('Promemoria', 'Rispondi', ['anna@example.it'], 'viewCalendario'), 'PUSH_INVIATA');

  const saved = JSON.parse(properties.get('PUSH_TOKENS'));
  assert.equal(saved['mario@example.it'].length, 1);
  const payload = JSON.parse(requests[0].options.payload);
  assert.equal(payload.message.android.notification.channel_id, 'gens-notifiche');
  assert.equal(payload.message.notification.body, 'Testo');
  const targetedPayload = JSON.parse(requests[2].options.payload);
  assert.equal(targetedPayload.message.notification.title, 'Firma');
  assert.equal(targetedPayload.message.data.viewId, 'viewFirma');
  const reminderPayload = JSON.parse(requests[3].options.payload);
  assert.equal(reminderPayload.message.notification.title, 'Promemoria');
  assert.equal(reminderPayload.message.data.viewId, 'viewCalendario');
  assert.equal(context.rimuoviTokenPush({ email: 'mario@example.it', password: 'bad' }), 'ERR_AUTH');
  assert.equal(context.rimuoviTokenPush({ email: 'mario@example.it', password: 'secret' }), 'TOKEN_RIMOSSO');
  assert.equal(JSON.parse(properties.get('PUSH_TOKENS'))['mario@example.it'], undefined);
});

test('votazioni: salva il voto e rifiuta un secondo voto dello stesso socio', () => {
  const rows = [['Timestamp', 'HashSocio', 'CandidatoEmail', 'Voto']];
  const voteSheet = dataSheet(rows);
  const ammissioniRows = [
    ['Data', 'Nome', 'Cognome', 'Email', 'Telefono', 'Sponsor 1', 'Sponsor 2', 'Stato', 'Esito'],
    [new Date(), 'Anna', 'Bianchi', 'anna@example.it', '', '', '', 'IN VOTAZIONE', 'DA VOTARE']
  ];
  const ammissioniSheet = dataSheet(ammissioniRows);
  ammissioniSheet.getRange = (row, column) => ({
    getValue: () => ammissioniRows[row - 1] && ammissioniRows[row - 1][column - 1],
    setValue(value) {
      while (ammissioniRows.length < row) ammissioniRows.push([]);
      ammissioniRows[row - 1][column - 1] = value;
    }
  });
  const sociSheet = dataSheet([
    ['Nome', 'Cognome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'Rossi', 'hash', 'mario@example.it', 'ATTIVO']
  ]);
  const context = {
    verificaLogin: () => 'OK_LOGIN',
    getHashUnivoco: () => 'hash-socio',
    LockService: {
      getScriptLock: () => ({ waitLock() { }, releaseLock() { } })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'Ammissioni' ? ammissioniSheet : voteSheet })
    },
    scriviLog() { },
    Date
  };

  loadScript('user.js', context);

  const voto = { email: 'mario@example.it', password: 'pw', emailCandidato: 'anna@example.it', voto: 'SI' };
  assert.equal(context.votaAmmissioneSocio(voto), 'OK');
  assert.equal(context.votaAmmissioneSocio(voto), 'GIA_VOTATO');
  assert.equal(rows.length, 2);
});

test('iscrizioni: richiesta pubblica, due avvalli e voto programmato al decimo giorno', () => {
  const sociRows = [
    ['Nome', 'Cognome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'Rossi', 'hash', 'mario@example.it', 'ATTIVO'],
    ['Anna', 'Bianchi', 'hash', 'anna@example.it', 'ATTIVO'],
    ['Luca', 'Verdi', 'hash', 'luca@example.it', 'NON ATTIVO']
  ];
  const ammissioniRows = [['Data Richiesta', 'Nome Candidato', 'Cognome Candidato', 'Email Candidato', 'Telefono', 'Sponsor 1', 'Sponsor 2', 'Stato Sostegno', 'Esito Assemblea']];
  const sociSheet = dataSheet(sociRows);
  const ammissioniSheet = dataSheet(ammissioniRows);
  ammissioniSheet.getRange = (row, column) => ({
    getValue() { return ammissioniRows[row - 1] ? ammissioniRows[row - 1][column - 1] : undefined; },
    setValue(value) {
      while (ammissioniRows.length < row) ammissioniRows.push([]);
      ammissioniRows[row - 1][column - 1] = value;
    }
  });
  const votiSheet = dataSheet([['Timestamp', 'HashSocio', 'CandidatoEmail', 'Voto']]);
  const sheets = { soci: sociSheet, Ammissioni: ammissioniSheet, VotiAmmissioni: votiSheet };
  let uuid = 0;
  const sentEmails = [];
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] || null, insertSheet: name => (sheets[name] = dataSheet([[]])) })
    },
    Utilities: { getUuid: () => 'richiesta-' + (++uuid) },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    MailApp: { sendEmail: (...args) => sentEmails.push(args) },
    inviaNotificaPush() { },
    getHashUnivoco: email => 'hash:' + email,
    verificaLogin: () => 'OK_LOGIN',
    scriviLog() { }
  };
  loadScript('user.js', context);

  const prima = context.richiediIscrizioneSocio({
    nome: 'Giulia', cognome: 'Neri', email: 'giulia@example.net', telefono: '12345',
    presentazione: 'Vorrei partecipare alle attività', consensoPrivacy: true
  });
  assert.equal(prima.ok, true);
  assert.equal(JSON.stringify(ammissioniRows[1].slice(5, 8)), JSON.stringify(['', '', 'RACCOLTA AVVALLI']));
  assert.equal(context.votaAmmissioneSocio({ email: 'mario@example.it', password: 'pw', emailCandidato: 'giulia@example.net', voto: 'FAVOREVOLE' }), 'ERR_VOTAZIONE_CHIUSA');
  assert.equal(context.getRichiesteIscrizioneDaAvvallare('luca@example.it').length, 0);
  assert.equal(context.getRichiesteIscrizioneDaAvvallare('giulia@example.net').length, 0);
  assert.equal(context.avvallaRichiestaIscrizione('mario@example.it', prima.richiestaId).ok, true);
  assert.equal(ammissioniRows[1][7], 'RACCOLTA AVVALLI');
  assert.ok(new Date(ammissioniRows[1][11]).getTime() > Date.now());
  assert.equal(context.avvallaRichiestaIscrizione('mario@example.it', prima.richiestaId).messaggio, 'Hai già avvallato questa richiesta.');
  assert.equal(context.avvallaRichiestaIscrizione('anna@example.it', prima.richiestaId).avvalli, 2);
  assert.equal(ammissioniRows[1][7], 'AVVALLATA - ATTESA VOTO');

  ammissioniRows[1][12] = new Date(Date.now() - 1000);
  context.processaScadenzeAmmissioni();
  assert.equal(ammissioniRows[1][7], 'IN VOTAZIONE');
  assert.equal(context.getAmmissioniInVoto('mario@example.it').length, 1);
  assert.equal(context.votaAmmissioneSocio({ email: 'mario@example.it', password: 'pw', emailCandidato: 'giulia@example.net', voto: 'FAVOREVOLE' }), 'OK');
  assert.ok(sentEmails.length >= 3);

  const seconda = context.richiediIscrizioneSocio({
    nome: 'Paolo', cognome: 'Blu', email: 'paolo@example.net', consensoPrivacy: true
  });
  assert.equal(seconda.ok, true);
  ammissioniRows[2][11] = new Date(Date.now() - 1000);
  context.processaScadenzeAmmissioni();
  assert.equal(ammissioniRows[2][7], 'SCADUTA');
  assert.equal(context.avvallaRichiestaIscrizione('mario@example.it', seconda.richiestaId).messaggio, 'La raccolta degli avvalli è chiusa.');
});

test('eventi: condivide la proposta e consente ai soci di aggiornare il voto', () => {
  const sociRows = [['Nome', 'Password', 'Email'], ['Mario', 'hash', 'mario@example.it']];
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato']];
  const sociSheet = dataSheet(sociRows);
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: { getUuid: () => 'evento-1' },
    verificaLogin: () => 'OK_LOGIN',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) }
  };

  loadScript('Eventi.js', context);

  const proposta = context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Cena sociale',
    opzioni: [{ data: '10/10/2026', ora: '20:00', luogo: 'Centro' }]
  });
  assert.equal(proposta.ok, true);
  const dataNonValida = context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Prova data',
    opzioni: [
      { data: '23/10/2026', ora: '11:22', luogo: '' },
      { data: '22/21/2027', ora: '23:00', luogo: '' }
    ]
  });
  assert.equal(dataNonValida.ok, false);
  assert.match(dataNonValida.messaggio, /data o un orario non è valido/i);
  assert.equal(context.getProposteEventi('mario@example.it').length, 1);
  assert.equal(context.votaPropostaEvento('mario@example.it', 'evento-1', 0).ok, true);
  const votoAggiornato = context.votaPropostaEvento('mario@example.it', 'evento-1', { interesse: 'no', opzioni: [] });
  assert.equal(votoAggiornato.ok, true);
  assert.equal(votoAggiornato.voti.interesse.si, 0);
  assert.equal(votoAggiornato.voti.interesse.no, 1);
  assert.equal(JSON.stringify(votoAggiornato.voti.opzioni), '[0]');
});

test('eventi: alla pubblicazione avvisa i soci attivi e installa un solo trigger reminder', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'hash', 'mario@example.it', 'ATTIVO'],
    ['Anna', 'hash', 'anna@example.it', 'attivo'],
    ['Paolo', 'hash', 'paolo@example.net', 'NON ATTIVO']
  ]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato']];
  const eventiSheet = dataSheet(eventiRows);
  const triggerList = [];
  const emails = [];
  const pushes = [];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null }) },
    Utilities: { getUuid: () => 'evento-notificato' },
    verificaLogin: () => 'OK_LOGIN',
    MailApp: { sendEmail: message => emails.push(message) },
    inviaNotificaPush: (titolo, testo, destinatari, vista) => pushes.push({ titolo, testo, destinatari, vista }),
    ScriptApp: {
      getProjectTriggers: () => triggerList,
      newTrigger: nome => ({
        timeBased: () => ({
          everyDays: () => ({
            atHour: () => ({
              create: () => {
                const trigger = { getHandlerFunction: () => nome };
                triggerList.push(trigger);
                return trigger;
              }
            })
          })
        })
      })
    }
  };
  loadScript('Eventi.js', context);

  const risultato = context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Cena sociale', opzioni: [{ data: '10/11/2026', ora: '20:00' }]
  });
  context.configuraTriggerPromemoriaProposteEvento_();

  assert.equal(risultato.ok, true);
  assert.equal(emails.length, 1);
  assert.equal(emails[0].bcc, 'mario@example.it,anna@example.it');
  assert.match(emails[0].subject, /Cena sociale/);
  assert.equal(pushes.length, 1);
  assert.deepEqual(pushes[0].destinatari, ['mario@example.it', 'anna@example.it']);
  assert.equal(pushes[0].vista, 'viewCalendario');
  assert.equal(triggerList.length, 1);
});

test('eventi: dopo 72 ore ricorda solo i soci senza voto una sola volta', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'hash', 'mario@example.it', 'ATTIVO'],
    ['Anna', 'hash', 'anna@example.it', 'ATTIVO']
  ]);
  const votes = {
    interesse: { si: 1, no: 0 },
    opzioni: [1],
    _votanti: { 'socio:mario@example.it': { interesse: 'si', opzioni: [0], email: 'mario@example.it' } }
  };
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', 'EventoCalendarId', 'OpzioneConfermataJson', 'PromemoriaNonVotantiInviatoIl'],
    ['evento-reminder', new Date(Date.now() - 4 * 86400000), 'mario@example.it', 'Cena sociale', '', '[{"data":"10/11/2026","ora":"20:00"}]', JSON.stringify(votes), 'ATTIVA', '', '', '']
  ];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({ setValue: value => { eventiRows[row - 1][column - 1] = value; } });
  const emails = [];
  const pushes = [];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null }) },
    MailApp: { sendEmail: message => emails.push(message) },
    inviaNotificaPush: (titolo, testo, destinatari) => pushes.push({ titolo, testo, destinatari })
  };
  loadScript('Eventi.js', context);

  assert.equal(context.inviaPromemoriaProposteEvento(), 1);
  assert.equal(context.inviaPromemoriaProposteEvento(), 0);
  assert.equal(emails.length, 1);
  assert.equal(emails[0].bcc, 'anna@example.it');
  assert.deepEqual(pushes[0].destinatari, ['anna@example.it']);
  assert.ok(eventiRows[1][10]);
});

test('notifiche: mostra avviso evento subito e promemoria dopo 72 ore finché non vota', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'hash', 'mario@example.it', 'ATTIVO']
  ]);
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato'],
    ['evento-in-app', new Date(Date.now() - 4 * 86400000), 'anna@example.it', 'Cena sociale', '', '[{"data":"10/11/2026","ora":"20:00"}]', '{}', 'ATTIVA']
  ];
  const eventiSheet = dataSheet(eventiRows);
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null }) }
  };
  loadScript('Eventi.js', context);
  loadScript('Auth.js', context);
  context.getDatiUtente = () => ({ scadenza: '' });

  const aperte = context.getNotificheUtente('mario@example.it');
  assert.ok(aperte.some(notifica => notifica.titolo === 'Nuovo questionario evento' && notifica.viewId === 'viewCalendario'));
  assert.ok(aperte.some(notifica => notifica.titolo === 'Promemoria questionario evento'));

  eventiRows[1][6] = JSON.stringify({ _votanti: { 'socio:mario@example.it': { interesse: 'no', opzioni: [], email: 'mario@example.it' } } });
  const dopoVoto = context.getNotificheUtente('mario@example.it');
  assert.equal(dopoVoto.some(notifica => /questionario evento/i.test(notifica.titolo)), false);
});

test('eventi: richiede almeno una preferenza per il sì ma non per il no', () => {
  const sociSheet = dataSheet([['Nome', 'Password', 'Email'], ['Mario', 'hash', 'mario@example.it']]);
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato'],
    ['evento-si-rapido', new Date(), 'admin@example.it', 'Cena sociale', '', '[{"data":"10/10/2026","ora":"20:00"}]', '{}', 'ATTIVA']
  ];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({ setValue(value) { eventiRows[row - 1][column - 1] = value; } });
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null }) },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) }
  };
  loadScript('Eventi.js', context);

  const senzaPreferenze = context.votaPropostaEvento('mario@example.it', 'evento-si-rapido', { interesse: 'si', opzioni: [] });
  assert.equal(senzaPreferenze.ok, false);
  assert.match(senzaPreferenze.messaggio, /se scegli Sì.*almeno una/i);

  const risultato = context.votaPropostaEvento('mario@example.it', 'evento-si-rapido', { interesse: 'si', opzioni: [0] });
  assert.equal(risultato.ok, true);
  assert.equal(risultato.voti.interesse.si, 1);
  assert.equal(risultato.voti.opzioni[0], 1);

  const rispostaNo = context.votaPropostaEvento('mario@example.it', 'evento-si-rapido', { interesse: 'no', opzioni: [] });
  assert.equal(rispostaNo.ok, true);
  assert.equal(rispostaNo.voti.interesse.no, 1);
});

test('eventi: ospiti votano interesse e più preferenze senza esporre le email', () => {
  const sociSheet = dataSheet([['Nome', 'Password', 'Email'], ['Mario', 'hash', 'mario@example.it']]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato']];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) { eventiRows[row - 1][column - 1] = value; }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'evento-ospiti',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: () => 'OK_LOGIN',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) }
  };

  loadScript('Eventi.js', context);
  context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Gita',
    descrizione: 'Scegli quando partecipare',
    opzioni: [
      { data: '10/10/2026', ora: '09:00', luogo: 'Centro' },
      { data: '11/10/2026', ora: '10:00', luogo: 'Centro' }
    ]
  });

  const senzaPreferenze = context.votaPropostaEvento('', 'evento-ospiti', {
    nome: 'Giulia', email: 'ospite@example.net', interesse: 'si', opzioni: []
  });
  assert.equal(senzaPreferenze.ok, false);
  assert.match(senzaPreferenze.messaggio, /se scegli Sì.*almeno una/i);

  const risposta = { nome: 'Giulia', email: 'ospite@example.net', interesse: 'si', opzioni: [0, 1] };
  assert.equal(context.votaPropostaEvento('', 'evento-ospiti', risposta).ok, true);
  assert.equal(context.votaPropostaEvento('', 'evento-ospiti', Object.assign({}, risposta, { email: 'OSPITE@example.net' })).messaggio, 'Hai già votato questa proposta.');
  assert.equal(context.votaPropostaEvento('', 'evento-ospiti', { nome: 'Paolo', email: 'paolo@example.net', interesse: 'no', opzioni: [] }).ok, true);

  const propostaPubblica = context.getPropostaEventoPubblica('evento-ospiti');
  assert.deepEqual(JSON.parse(JSON.stringify(propostaPubblica.voti)), {
    interesse: { si: 1, no: 1 },
    opzioni: [1, 1]
  });
  assert.equal(JSON.stringify(propostaPubblica).includes('ospite@example.net'), false);
  assert.equal(propostaPubblica.mioVoto, null);
  assert.deepEqual(context.getProposteEventi('').length, 0);
});

test('eventi: mostra ai soci chi ha scelto ogni data senza esporre le email', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email', 'Cognome'],
    ['Mario', 'hash', 'mario@example.it', 'Rossi']
  ]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato']];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) { eventiRows[row - 1][column - 1] = value; }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'evento-nomi',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: () => 'OK_LOGIN',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) }
  };

  loadScript('Eventi.js', context);
  context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Pizza',
    descrizione: '',
    opzioni: [
      { data: '10/10/2026', ora: '20:00', luogo: 'Centro' },
      { data: '11/10/2026', ora: '20:00', luogo: 'Centro' }
    ]
  });

  context.votaPropostaEvento('mario@example.it', 'evento-nomi', { interesse: 'si', opzioni: [0] });
  context.votaPropostaEvento('', 'evento-nomi', { nome: 'Giulia', email: 'ospite@example.net', interesse: 'si', opzioni: [0, 1] });

  const propostaSocio = context.getProposteEventi('mario@example.it')[0];
  assert.deepEqual(JSON.parse(JSON.stringify(propostaSocio.votantiPerOpzione)), [
    ['Giulia', 'Mario Rossi'],
    ['Giulia']
  ]);
  assert.equal(JSON.stringify(propostaSocio).includes('ospite@example.net'), false);

  const propostaPubblica = context.getPropostaEventoPubblica('evento-nomi');
  assert.deepEqual(JSON.parse(JSON.stringify(propostaPubblica.votantiPerOpzione)), []);
  assert.equal(JSON.stringify(propostaPubblica).includes('Mario Rossi'), false);
  assert.equal(JSON.stringify(propostaPubblica).includes('Giulia'), false);
});

test('eventi: aggiunge opzioni e crea Calendar di due ore per soli interessati', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email'],
    ['Mario', 'hash', 'mario@example.it'],
    ['Anna', 'hash', 'anna@example.it']
  ]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', '', '']];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      while (eventiRows[row - 1].length < column) eventiRows[row - 1].push('');
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  let calendarOptions = null;
  let calendarStart = null;
  let calendarEnd = null;
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'evento-calendar',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: dati => dati.password === 'pw' ? 'OK_LOGIN' : 'ERR_AUTH',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    CalendarApp: {
      getCalendarById: calendarId => {
        assert.equal(calendarId, 'gens.ssshhh@gmail.com');
        return {
          createEvent: (title, start, end, options) => {
            calendarStart = start;
            calendarEnd = end;
            calendarOptions = options;
            return { getId: () => 'calendar-event-1', setGuestsCanSeeGuests(value) { this.guestsVisible = value; } };
          }
        };
      }
    }
  };
  loadScript('Eventi.js', context);

  context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Sushi', descrizione: '', opzioni: [{ data: '20/11/2026', ora: '12:00', luogo: '' }]
  });
  assert.equal(context.aggiungiOpzionePropostaEvento('', 'evento-calendar', {
    data: '21/11/2026', ora: '13:00', luogo: ''
  }).ok, false);
  assert.equal(context.aggiungiOpzionePropostaEvento('mario@example.it', 'evento-calendar', {
    data: '21/11/2026', ora: '13:00', luogo: ''
  }).ok, true);
  const aggiuntaSocio = context.aggiungiOpzionePropostaEvento('anna@example.it', 'evento-calendar', {
    data: '22/11/2026', ora: '09:00', luogo: ''
  });
  assert.equal(aggiuntaSocio.ok, true);
  assert.equal(aggiuntaSocio.indice, 2);
  assert.equal(aggiuntaSocio.ricompilazioniRichieste, 0);
  assert.equal(context.aggiungiOpzionePropostaEvento('mario@example.it', 'evento-calendar', {
    data: '22/11/2026', ora: '09:00', luogo: ''
  }).ok, false);
  assert.equal(context.votaPropostaEvento('mario@example.it', 'evento-calendar', { interesse: 'si', opzioni: [0] }).ok, true);
  assert.equal(context.votaPropostaEvento('', 'evento-calendar', {
    nome: 'Paolo', email: 'paolo@example.net', interesse: 'si', opzioni: [1]
  }).ok, true);
  assert.equal(context.votaPropostaEvento('', 'evento-calendar', {
    nome: 'Lucia', email: 'lucia@example.net', interesse: 'no', opzioni: []
  }).ok, true);
  assert.equal(context.confermaEventoProposto({
    email: 'anna@example.it', password: 'pw', propostaId: 'evento-calendar', opzione: {
      data: '21/11/2026', ora: '14:30', luogo: 'Centro'
    }
  }).messaggio, 'Solo il creatore può confermare l’evento.');

  const conferma = context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-calendar', opzione: {
      data: '21/11/2026', ora: '14:30', luogo: 'Centro'
    }
  });
  assert.equal(conferma.ok, true);
  assert.equal(conferma.invitati, 2);
  assert.equal(calendarEnd.getTime() - calendarStart.getTime(), 2 * 60 * 60 * 1000);
  assert.deepEqual(calendarOptions.guests.split(','), ['mario@example.it', 'paolo@example.net']);
  assert.equal(calendarOptions.sendInvites, true);
  assert.equal(eventiRows[1][7], 'CONFERMATA');
  assert.equal(eventiRows[1][8], 'calendar-event-1');
});

test('eventi: aggiungere una data non invia notifiche e rende facoltativo aggiornare le preferenze', () => {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email'],
    ['Mario', 'hash', 'mario@example.it'],
    ['Anna', 'hash', 'anna@example.it']
  ]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', '', '']];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      while (eventiRows[row - 1].length < column) eventiRows[row - 1].push('');
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const posta = [];
  const push = [];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'evento-voto-aggiunta',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: dati => dati.password === 'pw' ? 'OK_LOGIN' : 'ERR_AUTH',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    MailApp: { sendEmail: messaggio => posta.push(messaggio) },
    inviaNotificaPush: (...args) => push.push(args),
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create() { } }) }) }) }) }
  };
  loadScript('Eventi.js', context);

  context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Cena', descrizione: '', opzioni: [{ data: '20/11/2026', ora: '12:00', luogo: '' }]
  });
  assert.equal(context.votaPropostaEvento('', 'evento-voto-aggiunta', {
    nome: 'Paolo', email: 'paolo@example.net', interesse: 'si', opzioni: [0]
  }).ok, true);
  assert.equal(context.votaPropostaEvento('anna@example.it', 'evento-voto-aggiunta', { interesse: 'no', opzioni: [] }).ok, true);

  posta.length = 0;
  push.length = 0;
  const aggiunta = context.aggiungiOpzionePropostaEvento('anna@example.it', 'evento-voto-aggiunta', {
    data: '22/11/2026', ora: '09:00', luogo: 'Bar'
  });
  assert.equal(aggiunta.ok, true);
  assert.equal(aggiunta.ricompilazioniRichieste, 0);

  const destinatari = posta.map(messaggio => messaggio.to || messaggio.bcc);
  assert.equal(destinatari.some(to => String(to).includes('mario@example.it')), true);
  assert.equal(posta.length, 1);
  assert.equal(posta.some(messaggio => String(messaggio.bcc || '').includes('paolo@example.net')), false);
  assert.equal(posta.some(messaggio => String(messaggio.bcc || '').includes('anna@example.it')), false);
  assert.match(posta[0].body, /restano valide/);
  assert.doesNotMatch(posta[0].body, /invitati a rispondere di nuovo/);
  assert.equal(push.length, 0);

  const voti = JSON.parse(eventiRows[1][6]);
  assert.ok(voti._ricompilazione['socio:anna@example.it'].indexOf('FACOLTATIVA:') === 0);
  assert.equal(voti._ricompilazione['socio:mario@example.it'], undefined);
  const chiaveOspite = 'ospite:' + context.hashEmailEvento_('paolo@example.net');
  assert.ok(voti._ricompilazione[chiaveOspite].indexOf('FACOLTATIVA:') === 0);
  const propostaSocio = context.getProposteEventi('anna@example.it')[0];
  assert.equal(propostaSocio.ricompilazioneRichiesta, false);
  assert.equal(propostaSocio.ricompilazioneFacoltativa, true);

  const pre = context.getRispostaRicompilazioneOspite({ id: 'evento-voto-aggiunta', email: 'paolo@example.net' });
  assert.equal(pre.nome, 'Paolo');
  assert.equal(pre.interesse, 'si');
  assert.equal(JSON.stringify(pre.opzioni), '[0]');
  assert.equal(pre.facoltativa, true);
  assert.equal(context.getRispostaRicompilazioneOspite({ id: 'evento-voto-aggiunta', email: 'anna@example.it' }), null);

  const rispostaOspite = context.votaPropostaEvento('', 'evento-voto-aggiunta', {
    nome: 'Paolo', email: 'paolo@example.net', interesse: 'si', opzioni: [0, 1]
  });
  assert.equal(rispostaOspite.ok, true);
  assert.equal(JSON.stringify(rispostaOspite.voti.opzioni), '[1,1]');
  assert.equal(JSON.parse(eventiRows[1][6])._ricompilazione[chiaveOspite], undefined);
  assert.equal(context.getRispostaRicompilazioneOspite({ id: 'evento-voto-aggiunta', email: 'paolo@example.net' }), null);
});

function creaContestoConfermaInviti_() {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email'],
    ['Mario', 'hash', 'mario@example.it']
  ]);
  const eventiRows = [['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', '', '']];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      while (eventiRows[row - 1].length < column) eventiRows[row - 1].push('');
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const creato = { calendarOptions: null };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'evento-inviti',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: dati => dati.password === 'pw' ? 'OK_LOGIN' : 'ERR_AUTH',
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    CalendarApp: {
      getCalendarById: () => ({
        createEvent: (title, start, end, options) => {
          creato.calendarOptions = options;
          return { getId: () => 'calendar-event-inviti', setGuestsCanSeeGuests() { } };
        }
      })
    }
  };
  loadScript('Eventi.js', context);

  context.creaPropostaEvento('mario@example.it', 'pw', {
    titolo: 'Cena', descrizione: '', opzioni: [
      { data: '21/11/2026', ora: '13:00', luogo: '' },
      { data: '22/11/2026', ora: '18:00', luogo: '' }
    ]
  });
  context.votaPropostaEvento('', 'evento-inviti', { nome: 'Paolo', email: 'paolo@example.net', interesse: 'si', opzioni: [0] });
  context.votaPropostaEvento('', 'evento-inviti', { nome: 'Lucia', email: 'lucia@example.net', interesse: 'si', opzioni: [1] });
  context.votaPropostaEvento('', 'evento-inviti', { nome: 'Sara', email: 'sara@example.net', interesse: 'si', opzioni: [0, 1] });
  context.votaPropostaEvento('', 'evento-inviti', { nome: 'Giulia', email: 'giulia@example.net', interesse: 'no', opzioni: [] });
  return { context, creato };
}

test('eventi: conferma invitando solo i Sì della data scelta', () => {
  const { context, creato } = creaContestoConfermaInviti_();
  const opzione = { data: '21/11/2026', ora: '13:00', luogo: '' };

  assert.equal(context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: true,
    opzione: { data: '23/11/2026', ora: '10:00', luogo: '' }
  }).messaggio, 'Nessun partecipante ha risposto Sì a questa data e ora.');

  const risultato = context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: true, opzione
  });
  assert.equal(risultato.ok, true);
  assert.equal(risultato.invitati, 2);
  assert.deepEqual(creato.calendarOptions.guests.split(','), ['paolo@example.net', 'sara@example.net']);
});

test('eventi: conferma invitando tutti i Sì anche se per altre date', () => {
  const { context, creato } = creaContestoConfermaInviti_();

  const risultato = context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: false,
    opzione: { data: '21/11/2026', ora: '13:00', luogo: '' }
  });
  assert.equal(risultato.ok, true);
  assert.equal(risultato.invitati, 3);
  assert.deepEqual(creato.calendarOptions.guests.split(','), ['paolo@example.net', 'lucia@example.net', 'sara@example.net']);
});

test('eventi: il creatore vede i partecipanti Sì e può escluderne alcuni dall\'invito', () => {
  const { context, creato } = creaContestoConfermaInviti_();
  const proposta = context.getProposteEventi('mario@example.it')[0];
  assert.equal(JSON.stringify(proposta.partecipantiSi.map(p => p.nome).sort()), '["Lucia","Paolo","Sara"]');
  const chiaveDi = nome => proposta.partecipantiSi.find(p => p.nome === nome).chiave;
  const opzione = { data: '21/11/2026', ora: '13:00', luogo: '' };

  assert.equal(context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: false,
    esclusi: proposta.partecipantiSi.map(p => p.chiave), opzione
  }).ok, false);

  const risultato = context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: false,
    esclusi: [chiaveDi('Lucia'), chiaveDi('Paolo')], opzione
  });
  assert.equal(risultato.ok, true);
  assert.equal(risultato.invitati, 1);
  assert.deepEqual(creato.calendarOptions.guests.split(','), ['sara@example.net']);
});

test('eventi: esclusione combinata con la sola data scelta', () => {
  const { context, creato } = creaContestoConfermaInviti_();
  const proposta = context.getProposteEventi('mario@example.it')[0];
  const chiaveDi = nome => proposta.partecipantiSi.find(p => p.nome === nome).chiave;

  const risultato = context.confermaEventoProposto({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-inviti', soloDataSelezionata: true,
    esclusi: [chiaveDi('Sara')], opzione: { data: '21/11/2026', ora: '13:00', luogo: '' }
  });
  assert.equal(risultato.ok, true);
  assert.deepEqual(creato.calendarOptions.guests.split(','), ['paolo@example.net']);
});

test('eventi: il creatore modifica la proposta mantenendo solo preferenze ancora valide e può annullarla', () => {
  const sociSheet = dataSheet([['Nome', 'Password', 'Email'], ['Mario', 'hash', 'mario@example.it']]);
  const originalVotes = {
    interesse: { si: 2, no: 0 },
    opzioni: [2, 0],
    _votanti: {
      'socio:mario@example.it': { interesse: 'si', opzioni: [0], email: 'mario@example.it' },
      'ospite:guest-hash': { interesse: 'si', opzioni: [0], email: 'ospite@example.net' }
    }
  };
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', '', ''],
    ['evento-edit', new Date(), 'mario@example.it', 'Cena', 'Vecchia descrizione', JSON.stringify([
      { data: '20/11/2026', ora: '18:00', luogo: 'Centro' },
      { data: '21/11/2026', ora: '18:00', luogo: '' }
    ]), JSON.stringify(originalVotes), 'ATTIVA']
  ];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      while (eventiRows[row - 1].length < column) eventiRows[row - 1].push('');
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const emails = [];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    MailApp: { sendEmail: message => emails.push(message) },
    verificaLogin: () => 'OK_LOGIN'
  };
  loadScript('Eventi.js', context);

  const aggiornamento = context.modificaPropostaEvento({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-edit',
    proposta: {
      titolo: 'Cena soci', descrizione: 'Descrizione aggiornata',
      opzioni: [
        { data: '20/11/2026', ora: '18:00', luogo: 'Ristorante' },
        { data: '22/11/2026', ora: '19:00', luogo: '' }
      ]
    }
  });
  assert.equal(aggiornamento.ok, true);
  assert.equal(aggiornamento.ricompilazioniRichieste, 1);
  assert.equal(eventiRows[1][3], 'Cena soci');
  assert.equal(eventiRows[1][4], 'Descrizione aggiornata');
  assert.deepEqual(JSON.parse(eventiRows[1][5]), [
    { data: '20/11/2026', ora: '18:00', luogo: 'Ristorante' },
    { data: '22/11/2026', ora: '19:00', luogo: '' }
  ]);
  assert.deepEqual(JSON.parse(eventiRows[1][6]).opzioni, [0, 0]);
  assert.ok(JSON.parse(eventiRows[1][6])._ricompilazione['ospite:guest-hash']);
  assert.equal(JSON.parse(eventiRows[1][6])._ricompilazione['socio:mario@example.it'], undefined);
  assert.equal(emails.length, 2);
  assert.equal(emails[0].bcc, 'ospite@example.net');
  assert.equal(emails[1].to, 'mario@example.it');
  assert.match(emails[1].body, /avvisato 1 persona/);

  const rimozioneVotata = context.modificaPropostaEvento({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-edit',
    proposta: { titolo: 'Cena soci', opzioni: [{ data: '22/11/2026', ora: '19:00' }] }
  });
  assert.equal(rimozioneVotata.ok, true);
  assert.equal(rimozioneVotata.ricompilazioniRichieste, 0);
  assert.deepEqual(JSON.parse(eventiRows[1][6]).opzioni, [0]);
  assert.ok(JSON.parse(eventiRows[1][6])._votanti['socio:mario@example.it']);

  assert.equal(context.annullaPropostaEvento({ email: 'other@example.net', password: 'pw', propostaId: 'evento-edit' }).ok, false);
  assert.equal(context.annullaPropostaEvento({ email: 'mario@example.it', password: 'pw', propostaId: 'evento-edit' }).ok, true);
  assert.equal(eventiRows[1][7], 'ANNULLATA');
});

function creaContestoRicompilazione_(votiIniziali, emails, pushes) {
  const sociSheet = dataSheet([
    ['Nome', 'Password', 'Email', 'Stato Socio'],
    ['Mario', 'hash', 'mario@example.it', 'ATTIVO'],
    ['Anna', 'hash', 'anna@example.it', 'ATTIVO']
  ]);
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato', '', ''],
    ['evento-refill', new Date(), 'mario@example.it', 'Pizza', '', JSON.stringify([
      { data: '20/11/2026', ora: '20:00', luogo: 'Centro' },
      { data: '22/11/2026', ora: '20:00', luogo: '' }
    ]), JSON.stringify(votiIniziali), 'ATTIVA']
  ];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({
    setValue(value) {
      while (eventiRows[row - 1].length < column) eventiRows[row - 1].push('');
      eventiRows[row - 1][column - 1] = value;
    }
  });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    Utilities: {
      getUuid: () => 'token-ospite',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    verificaLogin: () => 'OK_LOGIN',
    MailApp: { sendEmail: message => emails.push(message) },
    inviaNotificaPush: (titolo, testo, destinatari, vista) => pushes.push({ titolo, testo, destinatari, vista })
  };
  loadScript('Eventi.js', context);
  loadScript('Auth.js', context);
  context.verificaLogin = () => 'OK_LOGIN';
  context.getDatiUtente = () => ({ scadenza: '' });
  return { context, eventiRows };
}

test('eventi: avvisa solo chi perde una preferenza selezionata', () => {
  const emails = [];
  const pushes = [];
  const { context, eventiRows } = creaContestoRicompilazione_({
    interesse: { si: 4, no: 0 },
    opzioni: [3, 1],
    _votanti: {
      'socio:mario@example.it': { interesse: 'si', opzioni: [0], email: 'mario@example.it' },
      'socio:anna@example.it': { interesse: 'si', opzioni: [0], email: 'anna@example.it' },
      'ospite:invariato': { interesse: 'si', opzioni: [0], email: 'intatto@example.net', nome: 'Intatto' },
      'ospite:rimosso': { interesse: 'si', opzioni: [1], email: 'rimosso@example.net', nome: 'Rimosso' }
    }
  }, emails, pushes);

  const risultato = context.modificaPropostaEvento({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-refill',
    proposta: {
      titolo: 'Pizza', descrizione: '', opzioni: [
        { data: '20/11/2026', ora: '20:00', luogo: 'Centro' }
      ]
    }
  });

  assert.equal(risultato.ok, true);
  assert.equal(risultato.ricompilazioniRichieste, 1);
  assert.equal(emails.length, 2);
  assert.equal(emails[0].bcc, 'rimosso@example.net');
  assert.equal(emails[1].to, 'mario@example.it');
  assert.equal(pushes.length, 0);

  const voti = JSON.parse(eventiRows[1][6]);
  assert.equal(voti._ricompilazione['socio:anna@example.it'], undefined);
  assert.equal(voti._ricompilazione['ospite:invariato'], undefined);
  assert.ok(voti._ricompilazione['ospite:rimosso']);
  assert.equal(voti._votanti['ospite:invariato'].opzioni[0], 0);
  assert.deepEqual(voti.opzioni, [3]);
});

test('eventi: non avvisa i soci che perdono una preferenza se la notifica non è richiesta', () => {
  const emails = [];
  const pushes = [];
  const { context, eventiRows } = creaContestoRicompilazione_({
    interesse: { si: 1, no: 0 },
    opzioni: [1, 0],
    _votanti: {
      'socio:anna@example.it': { interesse: 'si', opzioni: [0], email: 'anna@example.it' }
    }
  }, emails, pushes);

  const risultato = context.modificaPropostaEvento({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-refill',
    proposta: {
      titolo: 'Pizza', descrizione: '', opzioni: [
        { data: '20/11/2026', ora: '20:00', luogo: 'Nuovo locale' },
        { data: '22/11/2026', ora: '20:00', luogo: '' }
      ]
    }
  });

  assert.equal(risultato.ok, true);
  assert.equal(risultato.ricompilazioniRichieste, 0);
  assert.equal(emails.length, 1);
  assert.equal(emails[0].to, 'mario@example.it');
  assert.equal(pushes.length, 0);
  assert.ok(JSON.parse(eventiRows[1][6])._ricompilazione['socio:anna@example.it']);
});

test('eventi: le modifiche non disturbano chi mantiene le date selezionate', () => {
  const emails = [];
  const pushes = [];
  const { context, eventiRows } = creaContestoRicompilazione_({
    interesse: { si: 1, no: 0 },
    opzioni: [1, 0],
    _votanti: { 'socio:anna@example.it': { interesse: 'si', opzioni: [0], email: 'anna@example.it' } }
  }, emails, pushes);
  const voti = () => JSON.parse(eventiRows[1][6]);
  const modifica = {
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-refill',
    proposta: {
      titolo: 'Pizza', descrizione: '', opzioni: [
        { data: '20/11/2026', ora: '20:00', luogo: 'Centro' },
        { data: '21/11/2026', ora: '20:00', luogo: '' }
      ]
    }
  };

  assert.equal(context.modificaPropostaEvento(Object.assign({}, modifica, { email: 'anna@example.it' })).ok, false);

  const primo = context.modificaPropostaEvento(modifica);
  assert.equal(primo.ok, true);
  assert.equal(primo.ricompilazioniRichieste, 0);
  assert.equal(emails.length, 1);
  assert.equal(emails[0].to, 'mario@example.it');
  assert.equal(pushes.length, 0);
  assert.equal(context.getNotificheUtente('anna@example.it').some(notifica => notifica.titolo === 'Ricompila il questionario evento'), false);
  const propostaAggiornata = context.getProposteEventi('anna@example.it')[0];
  assert.equal(propostaAggiornata.ricompilazioneRichiesta, false);
  assert.equal(JSON.stringify(propostaAggiornata.mioVoto.opzioni), '[0]');

  assert.equal(context.ricompilaVotoPropostaEvento('mario@example.it', 'evento-refill', true).ok, false);
  assert.equal(voti()._ricompilazione, undefined);
  const votoModificato = context.votaPropostaEvento('anna@example.it', 'evento-refill', { interesse: 'no', opzioni: [] });
  assert.equal(votoModificato.ok, true);
  assert.equal(votoModificato.voti.interesse.si, 0);
  assert.equal(votoModificato.voti.interesse.no, 1);
  assert.equal(context.votaPropostaEvento('anna@example.it', 'evento-refill', { interesse: 'si', opzioni: [0] }).ok, true);

  const secondo = context.modificaPropostaEvento(Object.assign({}, modifica, {
    inviaNotifica: true,
    proposta: {
      titolo: 'Pizza', descrizione: '', opzioni: [
        { data: '20/11/2026', ora: '20:00', luogo: 'Nuovo locale' },
        { data: '22/11/2026', ora: '21:00', luogo: '' }
      ]
    }
  }));
  assert.equal(secondo.ricompilazioniRichieste, 1);
  assert.ok(emails.some(message => message.bcc === 'anna@example.it'));
  assert.equal(pushes.length, 1);
  assert.equal(JSON.stringify(pushes[0].destinatari), '["anna@example.it"]');
  assert.equal(JSON.stringify(context.getProposteEventi('anna@example.it')[0].mioVoto.opzioni), '[]');
  const ricompila = context.ricompilaVotoPropostaEvento('anna@example.it', 'evento-refill', true);
  assert.equal(ricompila.ok, true);
  assert.equal(ricompila.ricompilato, true);
  assert.equal(voti().interesse.si, 0);
  assert.equal(voti().interesse.no, 0);
  assert.equal(JSON.stringify(voti().opzioni), '[0,0]');
  assert.equal(voti()._votanti['socio:anna@example.it'], undefined);

  assert.equal(context.votaPropostaEvento('anna@example.it', 'evento-refill', { interesse: 'si', opzioni: [1] }).ok, true);
  assert.deepEqual(voti().opzioni, [0, 1]);
  assert.equal(context.votaPropostaEvento('anna@example.it', 'evento-refill', { interesse: 'no', opzioni: [] }).ok, true);
});

test('eventi: la richiesta di ricompilazione porta i soci al sito e gli esterni al questionario', () => {
  const emails = [];
  const pushes = [];
  const context = {
    MailApp: { sendEmail: message => emails.push(message) },
    inviaNotificaPush: (...args) => pushes.push(args)
  };
  loadScript('Eventi.js', context);

  context.inviaRichiestaRicompilazionePropostaEvento_([
    { socio: true, email: 'anna@example.it' },
    { socio: false, email: 'ospite@example.net' },
    { socio: true, email: 'anna@example.it' }
  ], 'Cena soci', 'evento-x');

  assert.equal(emails.length, 2);
  const emailSocio = emails.find(message => message.bcc === 'anna@example.it');
  const emailEsterno = emails.find(message => message.bcc === 'ospite@example.net');
  assert.ok(emailSocio.body.includes('https://Gens-Ssshhh.github.io/Sito-web-Gens-Ssshhh/index.html'));
  assert.equal(emailSocio.body.includes('?evento='), false);
  assert.equal(emailSocio.htmlBody.includes('?evento='), false);
  assert.ok(emailEsterno.body.includes('https://Gens-Ssshhh.github.io/Sito-web-Gens-Ssshhh/index.html?evento=evento-x'));
  assert.ok(emailEsterno.htmlBody.includes('?evento=evento-x'));
  assert.equal(pushes.length, 1);
  assert.equal(JSON.stringify(pushes[0][2]), '["anna@example.it"]');
});

test('eventi: le email di nuovo questionario e promemoria portano i soci al sito, non al questionario esterno', () => {
  const emails = [];
  const context = {
    MailApp: { sendEmail: message => emails.push(message) },
    inviaNotificaPush: () => 'PUSH_INVIATA'
  };
  loadScript('Eventi.js', context);

  context.inviaComunicazioniPropostaEvento_(['anna@example.it'], 'Cena soci', 'evento-x', false);
  context.inviaComunicazioniPropostaEvento_(['anna@example.it'], 'Cena soci', 'evento-x', true);

  assert.equal(emails.length, 2);
  emails.forEach(message => {
    assert.ok(message.body.includes('https://Gens-Ssshhh.github.io/Sito-web-Gens-Ssshhh/index.html'));
    assert.equal(message.body.includes('?evento='), false);
    assert.equal(message.htmlBody.includes('?evento='), false);
    assert.ok(message.htmlBody.includes('Accedi al portale soci'));
  });
});

test('eventi: un ospite con richiesta di ricompilazione sostituisce il voto precedente', () => {
  const emails = [];
  const pushes = [];
  const { context, eventiRows } = creaContestoRicompilazione_({ interesse: { si: 0, no: 0 }, opzioni: [0, 0], _votanti: {} }, emails, pushes);
  const voti = () => JSON.parse(eventiRows[1][6]);
  const risposta = opzioni => ({ nome: 'Lucia', email: 'lucia@example.net', interesse: 'si', opzioni });

  assert.equal(context.votaPropostaEvento('', 'evento-refill', risposta([0])).ok, true);
  assert.equal(emails.length, 1);
  assert.equal(emails[0].to, 'lucia@example.net');
  assert.match(emails[0].body, /Partecipo/);
  assert.match(emails[0].body, /modifica=token-ospite/);
  emails.length = 0;
  const modifica = context.modificaPropostaEvento({
    email: 'mario@example.it', password: 'pw', propostaId: 'evento-refill',
    proposta: {
      titolo: 'Pizza', descrizione: '', opzioni: [
        { data: '20/11/2026', ora: '20:00', luogo: 'Altro locale' },
        { data: '21/11/2026', ora: '20:00', luogo: '' }
      ]
    }
  });
  assert.equal(modifica.ricompilazioniRichieste, 1);
  assert.equal(emails[0].bcc, 'lucia@example.net');

  const nuovoVoto = context.votaPropostaEvento('', 'evento-refill', risposta([1]));
  assert.equal(nuovoVoto.ok, true);
  assert.equal(JSON.stringify(nuovoVoto.voti.opzioni), JSON.stringify([0, 1]));
  assert.equal(nuovoVoto.voti.interesse.si, 1);
  assert.equal(voti()._ricompilazione, undefined);

  const prefill = context.getRispostaRicompilazioneOspite({ id: 'evento-refill', token: 'token-ospite' });
  assert.equal(prefill.email, 'lucia@example.net');
  assert.equal(prefill.interesse, 'si');
  assert.equal(context.getRispostaRicompilazioneOspite({ id: 'evento-refill', token: 'sbagliato' }), null);
  assert.equal(context.votaPropostaEvento('', 'evento-refill', risposta([0])).messaggio, 'Hai già votato questa proposta.');
  const viaLink = context.votaPropostaEvento('', 'evento-refill', Object.assign(risposta([0]), { token: 'token-ospite' }));
  assert.equal(viaLink.ok, true);
  assert.equal(JSON.stringify(viaLink.voti.opzioni), JSON.stringify([1, 0]));
  assert.equal(viaLink.voti.interesse.si, 1);
});

test('eventi: il dispatcher serve il link pubblico e registra il voto esterno', () => {
  const sociSheet = dataSheet([['Nome', 'Password', 'Email'], ['Mario', 'hash', 'mario@example.it']]);
  const eventiRows = [
    ['ID', 'CreatoIl', 'CreatoreEmail', 'Titolo', 'Descrizione', 'OpzioniJson', 'VotiJson', 'Stato'],
    ['evento-web', new Date(), 'mario@example.it', 'Passeggiata', 'Percorso cittadino', JSON.stringify([{ data: '10/11/2026', ora: '10:00', luogo: 'Centro' }]), '{}', 'ATTIVA']
  ];
  const eventiSheet = dataSheet(eventiRows);
  eventiSheet.getRange = (row, column) => ({ setValue(value) { eventiRows[row - 1][column - 1] = value; } });
  const spreadsheet = {
    getSheetByName: name => name === 'soci' ? sociSheet : name === 'ProposteEventi' ? eventiSheet : null,
    insertSheet: () => eventiSheet
  };
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Utilities: {
      getUuid: () => 'token-web',
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, value) => Array.from(Buffer.from(value, 'utf8').subarray(0, 4))
    },
    LockService: { getScriptLock: () => ({ waitLock() { }, releaseLock() { } }) },
    ContentService: {
      MimeType: { JSON: 'JSON' },
      createTextOutput: content => ({ getContent: () => content, setMimeType() { return this; } })
    }
  };
  loadScript('Eventi.js', context);
  let helperInvoked = false;
  context.creaAccountSocioDaAmmissione_ = () => { helperInvoked = true; };
  loadScript('main.js', context);

  const rispostaEndpoint = richiesta => JSON.parse(context.doPost({ postData: { contents: JSON.stringify(richiesta) } }).getContent());
  const azionePrivata = rispostaEndpoint({ azione: 'creaAccountSocioDaAmmissione_', payload: { email: 'evil@example.net' } });
  assert.equal(azionePrivata.status, 'AUTH_ERROR');
  assert.equal(helperInvoked, false);
  const proposta = rispostaEndpoint({ azione: 'getPropostaEventoPubblica', payload: { id: 'evento-web' } });
  assert.equal(proposta.status, 'SUCCESS');
  assert.equal(proposta.data.titolo, 'Passeggiata');

  const voto = rispostaEndpoint({
    azione: 'votaPropostaEvento',
    payload: ['', 'evento-web', { nome: 'Giulia', email: 'giulia@example.net', interesse: 'si', opzioni: [0] }]
  });
  assert.equal(voto.data.ok, true);
});

test('api: crea sessioni dopo il login e vincola le azioni admin all’identità autenticata', () => {
  const cacheValues = new Map();
  let updatePayload = null;
  let datiUtenteRichiesti = null;
  let tokenCounter = 0;
  const context = {
    CacheService: {
      getScriptCache: () => ({
        get: key => cacheValues.get(key) || null,
        put: (key, value) => cacheValues.set(key, value),
        remove: key => cacheValues.delete(key)
      })
    },
    Utilities: { getUuid: () => 'session-token-' + (++tokenCounter) },
    ContentService: {
      MimeType: { JSON: 'JSON' },
      createTextOutput: content => ({ getContent: () => content, setMimeType() { return this; } })
    },
    verificaLogin: ({ email, password }) => ['admin@example.it', 'member@example.it'].includes(email) && password === 'pw' ? 'OK_LOGIN' : 'ERR_CREDENZIALI',
    getDatiUtente: email => {
      datiUtenteRichiesti = email;
      return { email, ruolo: email === 'admin@example.it' ? 'Presidente' : 'Socio semplice' };
    },
    isRuoloAmministrativo: ruolo => ruolo === 'Presidente',
    adminUpdateSocio: dati => { updatePayload = dati; return 'OK'; }
  };
  loadScript('main.js', context);

  const rispostaEndpoint = richiesta => JSON.parse(context.doPost({ postData: { contents: JSON.stringify(richiesta) } }).getContent());
  const senzaSessione = rispostaEndpoint({ azione: 'adminUpdateSocio', payload: { adminEmail: 'admin@example.it' } });
  assert.equal(senzaSessione.status, 'AUTH_ERROR');
  assert.equal(updatePayload, null);

  const azioneInterna = rispostaEndpoint({ azione: 'autorizzaDriveVerbali', token: 'invalid-token' });
  assert.equal(azioneInterna.status, 'AUTH_ERROR');

  const login = rispostaEndpoint({ azione: 'login', email: 'admin@example.it', password: 'pw' });
  assert.equal(login.status, 'SUCCESS');
  assert.equal(login.token, 'session-token-1');

  const azioneNonConsentita = rispostaEndpoint({ azione: 'autorizzaDriveVerbali', token: login.token });
  assert.equal(azioneNonConsentita.status, 'ERROR');

  const modifica = rispostaEndpoint({
    azione: 'adminUpdateSocio',
    token: login.token,
    payload: { adminEmail: 'other-admin@example.it', email: 'socio@example.it' }
  });
  assert.equal(modifica.data, 'OK');
  assert.equal(updatePayload.adminEmail, 'admin@example.it');

  const profilo = rispostaEndpoint({ azione: 'getDatiUtente', token: login.token, payload: 'other@example.it' });
  assert.equal(profilo.data.email, 'admin@example.it');
  assert.equal(datiUtenteRichiesti, 'admin@example.it');

  const memberLogin = rispostaEndpoint({ azione: 'login', email: 'member@example.it', password: 'pw' });
  const forbidden = rispostaEndpoint({
    azione: 'adminUpdateSocio',
    token: memberLogin.token,
    payload: { adminEmail: 'admin@example.it', email: 'victim@example.it' }
  });
  assert.equal(forbidden.status, 'ERROR');
  assert.equal(updatePayload.adminEmail, 'admin@example.it');

  const clientSource = fs.readFileSync(path.join(root, 'www', 'app.js'), 'utf8');
  const clientActions = Array.from(clientSource.matchAll(/chiamaServer\(\s*['"]([^'"]+)['"]/g), match => match[1]);
  assert.deepEqual(
    Array.from(new Set(clientActions)).filter(action => !context.AZIONI_API_CONSENTITE.includes(action)),
    []
  );

  const logout = rispostaEndpoint({ azione: 'logout', token: login.token });
  assert.equal(logout.data, true);
  const sessioneRevocata = rispostaEndpoint({ azione: 'getDatiUtente', token: login.token, payload: 'admin@example.it' });
  assert.equal(sessioneRevocata.status, 'AUTH_ERROR');
});

test('ruoli: riconosce come amministrativi presidente, vicepresidente e segretario ma non consiglieri', () => {
  const context = {};
  loadScript('user.js', context);

  assert.equal(context.isRuoloAmministrativo('Presidente'), true);
  assert.equal(context.isRuoloAmministrativo('Vicepresidente'), true);
  assert.equal(context.isRuoloAmministrativo('Segretario'), true);
  assert.equal(context.isRuoloAmministrativo('Tesoriere'), true);
  assert.equal(context.isRuoloAmministrativo('Consigliere'), false);
  assert.equal(context.isRuoloAmministrativo('Socio semplice'), false);
});

test('firme: invia OTP e lega il codice al documento richiesto', () => {
  const cacheValues = new Map();
  let emailSent = null;
  const context = {
    getDatiUtente: () => ({ email: 'mario@example.it' }),
    CacheService: {
      getScriptCache: () => ({
        put: (key, value, ttl) => cacheValues.set(key, { value, ttl }),
        get: key => (cacheValues.get(key) || {}).value || null,
        remove: key => cacheValues.delete(key)
      })
    },
    MailApp: { sendEmail: message => { emailSent = message; } },
    Math,
    String
  };

  loadScript('Signature.js', context);

  assert.equal(context.requestSignOTP('mario@example.it', 'doc-123'), 'OTP_SENT');
  assert.match(cacheValues.get('SIGN_OTP_mario@example.it').value, /^\d{6}$/);
  assert.equal(cacheValues.get('SIGN_OTP_mario@example.it').ttl, 300);
  assert.equal(cacheValues.get('SIGN_DOC_mario@example.it').value, 'doc-123');
  assert.equal(emailSent.to, 'mario@example.it');
});

test('documenti: elenca categorie e carica un file nella cartella pubblica', () => {
  const files = [
    { getId: () => 'file-1', getName: () => 'verbale.pdf', getUrl: () => 'https://drive/file-1', getDescription: () => '' }
  ];
  let uploaded = null;
  const folder = {
    getName: () => 'Verbali Ufficiali',
    getFiles: () => ({ hasNext: () => files.length > 0, next: () => files.shift() }),
    getFolders: () => ({ hasNext: () => false }),
    createFile: blob => { uploaded = blob; return { setDescription: description => { uploaded.description = description; } }; }
  };
  const config = { getRange: () => ({ getValue: () => 'public-folder' }) };
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'Config' ? config : null })
    },
    DriveApp: { getFolderById: () => folder },
    Utilities: {
      newBlob: (content, mimeType, name) => ({ content, mimeType, name }),
      base64Decode: value => Buffer.from(value, 'base64')
    },
    scriviLog() { }
  };

  loadScript('Documents.js', context);

  const listed = context.getListaDocumenti('PUBBLICO', 'mario@example.it');
  assert.equal(JSON.stringify(listed), JSON.stringify([{ id: 'file-1', nome: 'verbale.pdf', url: 'https://drive/file-1', categoria: 'Verbale' }]));
  assert.equal(context.uploadFile({ content: Buffer.from('pdf').toString('base64'), mimeType: 'application/pdf', filename: 'nuovo.pdf', category: 'Bilancio' }, 'PUBBLICO', 'mario@example.it'), 'UPLOAD_OK');
  assert.equal(uploaded.name, 'nuovo.pdf');
  assert.equal(uploaded.description, 'Bilancio');
});

test('cariche: una modifica diretta alla colonna M notifica il socio aggiornato e il titolare attuale', () => {
  const rows = [
    ['Nome', 'Password', 'Email', '', 'Cognome', '', '', '', '', '', '', 'Stato Socio', 'Carica sociale'],
    ['Giulia', '', 'giulia@example.it', '', 'Neri', '', '', '', '', '', '', 'attivo', 'Tesoriere'],
    ['Mario', '', 'mario@example.it', '', 'Rossi', '', '', '', '', '', '', 'attivo', 'Tesoriere']
  ];
  const sociSheet = {
    getName: () => 'soci',
    getDataRange: () => ({ getValues: () => rows }),
    getRange: () => ({ getValue: () => 'Tesoriere' })
  };
  const messages = [];
  const context = {
    MailApp: { sendEmail: message => messages.push(message) },
    console,
    scriviLog() { }
  };
  loadScript('Cariche.js', context);

  context.gestisciModificaCaricaIstituzionale({
    range: {
      getSheet: () => sociSheet,
      getColumn: () => 13,
      getRow: () => 2,
      getNumRows: () => 1,
      getNumColumns: () => 1,
      getValue: () => 'Tesoriere'
    },
    oldValue: 'Socio semplice'
  });

  assert.deepEqual(messages.map(message => message.to).sort(), ['giulia@example.it', 'mario@example.it']);
  assert.match(messages.find(message => message.to === 'giulia@example.it').body, /Socio semplice.*Tesoriere/);
  assert.match(messages.find(message => message.to === 'mario@example.it').body, /assegnata a Giulia Neri/);
});

test('cariche: adminUpdateSocio notifica chi assume la carica e chi la ricopre', () => {
  const rows = [
    ['Nome', 'Password', 'Email', '', 'Cognome', 'Telefono', '', '', '', 'Numero Tessera', '', 'Stato Socio', 'Carica sociale'],
    ['Giulia', '', 'giulia@example.it', '', 'Neri', '', '', '', '', '2', '', 'attivo', 'Socio semplice'],
    ['Mario', '', 'mario@example.it', '', 'Rossi', '', '', '', '', '3', '', 'attivo', 'Tesoriere'],
    ['Admin', '', 'admin@example.it', '', 'Verdi', '', '', '', '', '1', '', 'attivo', 'Presidente']
  ];
  const sociSheet = {
    getDataRange: () => ({ getValues: () => rows }),
    getRange: (row, column) => ({
      setValue(value) { rows[row - 1][column - 1] = value; }
    })
  };
  const messages = [];
  const context = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'soci' ? sociSheet : null }) },
    MailApp: { sendEmail: message => messages.push(message) },
    getDatiUtente: email => email === 'admin@example.it' ? { ruolo: 'Presidente' } : null,
    scriviLog() { }
  };
  loadScript('Cariche.js', context);
  loadScript('Admin.js', context);

  assert.equal(context.adminUpdateSocio({
    adminEmail: 'admin@example.it',
    tesseraTarget: '2',
    email: 'giulia@example.it',
    telefono: '',
    stato: 'attivo',
    ruolo: 'Tesoriere'
  }), 'OK');

  assert.deepEqual(messages.map(message => message.to).sort(), ['giulia@example.it', 'mario@example.it']);
  assert.equal(rows[1][12], 'Tesoriere');
});
