// ===== Black Friday 2026 · Captação de leads -> Google Sheets + WhatsApp (API oficial da Meta) =====
// Recebe os cadastros da página humbertovecchi.github.io/blackfriday
//
// Envio de WhatsApp: configure em Configurações do projeto (⚙) > Propriedades do script:
//   WA_TOKEN     token de acesso permanente da Meta (usuário do sistema)
//   WA_PHONE_ID  "Identificação do número de telefone" do WhatsApp Business
//   WA_TEMPLATE  nome do modelo de mensagem aprovado (ex.: boas_vindas_black_friday)
//   WA_LANG      idioma do modelo (ex.: pt_BR)
//   WA_USA_NOME  'sim' se o modelo tiver {{1}} para o primeiro nome (senão, deixe vazio)
// Sem essas propriedades, o cadastro é salvo normalmente e o envio é ignorado.

// Rode esta função uma vez (botão Executar) para autorizar o envio de WhatsApp.
function autorizar() {
  PropertiesService.getScriptProperties().setProperties({
    WA_PHONE_ID: '1117190214800496',      // número +55 62 8515-0061
    WA_TEMPLATE: 'blackfriday_boasvindas',
    WA_LANG: 'pt_BR'
  });
  UrlFetchApp.fetch('https://graph.facebook.com/', { muteHttpExceptions: true });
  Logger.log('Autorizado. Pode publicar a nova versão.');
}

function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var linha = [new Date(), p.nome || '', p.whatsapp || '', p.email || '', p.origem || '',
               p.utm_source || '', p.utm_medium || '', p.utm_campaign || '', p.utm_content || '', ''];
  var sheet, numLinha = 0;

  // 1) Grava o cadastro. Se a fila estiver cheia (pico de acessos), grava mesmo assim.
  var lock = LockService.getScriptLock();
  var temLock = lock.tryLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.getSheetByName('Black Friday') || ss.insertSheet('Black Friday');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Formulário',
                       'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'Mensagem WhatsApp']);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#151515').setFontColor('#D9B56E');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow(linha);
    if (temLock) numLinha = sheet.getLastRow();
  } catch (err) {
    return ContentService.createTextOutput('erro: ' + err);
  } finally {
    if (temLock) lock.releaseLock();
  }

  // 2) Envia a mensagem de boas-vindas e anota o resultado (falha aqui não perde o cadastro).
  try {
    var status = enviaWhatsApp(p.nome || '', p.whatsapp || '');
    if (numLinha) sheet.getRange(numLinha, 10).setValue(status);
  } catch (err) {
    if (numLinha) sheet.getRange(numLinha, 10).setValue('erro: ' + String(err).slice(0, 120));
  }
  return ContentService.createTextOutput('ok');
}

// Envia o modelo aprovado para o número cadastrado. Retorna o status para a planilha.
function enviaWhatsApp(nome, telefone) {
  var cfg = PropertiesService.getScriptProperties().getProperties();
  if (!cfg.WA_TOKEN || !cfg.WA_PHONE_ID || !cfg.WA_TEMPLATE) return 'não configurado';

  var numero = String(telefone).replace(/\D/g, '');
  if (numero.length > 11 && numero.indexOf('55') === 0) { /* já tem 55 */ }
  else if (numero.length === 10 || numero.length === 11) numero = '55' + numero; // DDD + número -> Brasil
  if (numero.length < 12) return 'número inválido';

  var primeiroNome = String(nome).trim().split(/\s+/)[0] || 'tudo bem';
  var corpo = {
    messaging_product: 'whatsapp',
    to: numero,
    type: 'template',
    template: {
      name: cfg.WA_TEMPLATE,
      language: { code: cfg.WA_LANG || 'pt_BR' }
    }
  };
  if (String(cfg.WA_USA_NOME).toLowerCase() === 'sim') {
    corpo.template.components = [{ type: 'body', parameters: [{ type: 'text', text: primeiroNome }] }];
  }

  try {
    var resp = UrlFetchApp.fetch('https://graph.facebook.com/v21.0/' + cfg.WA_PHONE_ID + '/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + cfg.WA_TOKEN },
      payload: JSON.stringify(corpo),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) return 'enviada';
    var msg = '';
    try { msg = JSON.parse(resp.getContentText()).error.message; } catch (_) { msg = resp.getContentText(); }
    return 'erro ' + code + ': ' + String(msg).slice(0, 120);
  } catch (err) {
    return 'erro: ' + String(err).slice(0, 120);
  }
}
