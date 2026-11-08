const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

test('i promemoria di voto e candidatura rimandano all’area riservata', () => {
  const scadenza = new Date();
  scadenza.setDate(scadenza.getDate() + 3);
  const socio = [];
  socio[2] = 'socio@example.it';
  socio[11] = 'attivo';
  const elezione = [];
  elezione[0] = 'elezione-1';
  elezione[1] = 'Elezione';
  elezione[2] = 'CANDIDATI';
  elezione[4] = scadenza;
  elezione[9] = scadenza;
  const sheets = {
    Database_Elezioni: [[], elezione],
    Soci: [[], socio]
  };
  const messages = [];
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => sheets[name]
          ? { getDataRange: () => ({ getValues: () => sheets[name] }) }
          : null
      })
    },
    getHashUnivoco: email => email,
    Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.it' }) },
    filtraDestinatariNotifica_: (list, type, channel) => channel === 'email' ? list : [],
    notificaAbilitata_: () => false,
    MailApp: { sendEmail: message => messages.push(message) },
    Utilities: { sleep() {} }
  };

  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '..', 'Elezioni.js'), 'utf8'), context);
  context.inviaPromemoriaElezioniAutomatizzati();

  assert.equal(messages.length, 2);
  for (const message of messages) {
    assert.equal(message.bcc, 'socio@example.it');
    assert.match(message.htmlBody, /href="https:\/\/sites\.google\.com\/view\/gens-ssshhh\/area-riservata"[^>]*>ACCEDI AL CENTRO ELETTORALE<\/a>/);
  }
});
