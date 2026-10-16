// ===== Black Friday 2026 · Captação de leads -> Google Sheets =====
// 1. Crie uma planilha no Google Sheets.
// 2. Extensões > Apps Script > apague o que estiver lá, cole este código e salve.
// 3. Implantar > Nova implantação > Tipo: App da Web
//    Executar como: Eu · Quem pode acessar: Qualquer pessoa > Implantar.
//    (o Google pede para autorizar: escolha sua conta > Avançado > Acessar)
// 4. Copie a URL gerada (termina em /exec) e cole em SCRIPT_URL no index.html.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que dois cadastros simultâneos se sobrescrevam
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Black Friday') || ss.insertSheet('Black Friday');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Formulário',
                       'utm_source', 'utm_medium', 'utm_campaign', 'utm_content']);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#151515').setFontColor('#D9B56E');
      sheet.setFrozenRows(1);
    }

    var p = (e && e.parameter) ? e.parameter : {};
    sheet.appendRow([
      new Date(),
      p.nome || '',
      p.whatsapp || '',
      p.email || '',
      p.origem || '',
      p.utm_source || '',
      p.utm_medium || '',
      p.utm_campaign || '',
      p.utm_content || ''
    ]);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('erro: ' + err);
  } finally {
    lock.releaseLock();
  }
}
