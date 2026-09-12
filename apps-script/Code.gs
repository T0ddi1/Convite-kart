// Apps Script "webhook" para a planilha "seleção kart".
// Como instalar:
// 1. Abra a planilha no navegador.
// 2. Menu Extensões > Apps Script.
// 3. Apague o conteúdo de Code.gs e cole este arquivo inteiro.
// 4. Clique em Implantar > Nova implantação > tipo "App da Web".
//    - Executar como: Eu (sua conta)
//    - Quem pode acessar: Qualquer pessoa
// 5. Copie a URL do App da Web gerada e cole na constante WEBHOOK_URL do site (index.html).
// 6. Sempre que editar o script, gere uma NOVA implantação (ou "Gerenciar implantações" > editar) para a URL continuar valendo.

var SHEET_NAME = 'Página1';

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// GET: devolve a lista de pilotos confirmados.
// Suporta JSONP (?callback=nome) para ser lida direto do navegador sem problema de CORS.
function doGet(e) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < values.length; i++) {
    var nome = values[i][0];
    var indicadoPor = values[i][1];
    if (nome) {
      rows.push({
        nome: String(nome).trim(),
        indicadoPor: indicadoPor ? String(indicadoPor).trim() : ''
      });
    }
  }

  var json = JSON.stringify(rows);

  if (e && e.parameter && e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// POST: recebe { nome, indicadoPor } e adiciona uma linha na planilha.
// Evita duplicar o mesmo nome (comparação sem diferenciar maiúsculas/minúsculas).
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var payload = JSON.parse(e.postData.contents);
    var nome = (payload.nome || '').toString().trim();
    var indicadoPor = (payload.indicadoPor || '').toString().trim();

    if (!nome) {
      return jsonOutput_({ ok: false, error: 'Nome é obrigatório.' });
    }

    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var jaExiste = false;

    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === nome.toLowerCase()) {
        jaExiste = true;
        break;
      }
    }

    if (!jaExiste) {
      sheet.appendRow([nome, indicadoPor]);
    }

    return jsonOutput_({ ok: true, created: !jaExiste });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
