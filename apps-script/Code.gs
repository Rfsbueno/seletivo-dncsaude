// ============================================================
//  DNC Saúde+ — Formulário de Processo Seletivo
//  Google Apps Script — grava no Sheets + envia e-mail
// ============================================================
//
//  COMO ATUALIZAR (você já tem o projeto criado):
//  1. Abra script.google.com e acesse o projeto
//  2. Substitua TODO o código por este arquivo
//  3. Salve (Ctrl+S)
//  4. Clique em "Implantar" > "Gerenciar implantações"
//  5. Clique no lápis (editar) na implantação existente
//  6. Em "Versão" selecione "Nova versão"
//  7. Clique em "Implantar"
//  A URL permanece a mesma — não precisa atualizar o HTML.
// ============================================================

const SPREADSHEET_ID = '';
const SHEET_NAME = 'Inscrições';
const EMAIL_NOTIFICACAO = ''; // preencher no script.google.com (nao versionar o endereco: repo publico)

const HEADERS = [
  'Data de Inscrição',
  'Nome Completo',
  'Telefone/WhatsApp',
  'E-mail',
  'Escolaridade',
  'CEP',
  'Logradouro',
  'Bairro',
  'Cidade',
  'Estado',
  'Já trabalhou com vendas?',
  'Possui computador/notebook?',
  'Possui internet em casa?',
  'Produtos que já vendeu'
];

// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    salvarNaPlanilha(data);
    enviarEmailNotificacao(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────
function getOrCreateSheet() {
  let ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    const files = DriveApp.getFilesByName('DNC Saúde+ — Processo Seletivo');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('DNC Saúde+ — Processo Seletivo');
    }
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const defaultSheet = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
    criarCabecalho(sheet);
  }
  return sheet;
}

// ─────────────────────────────────────────────
function criarCabecalho(sheet) {
  const r = sheet.getRange(1, 1, 1, HEADERS.length);
  r.setValues([HEADERS]);
  r.setBackground('#1a2980')
   .setFontColor('#FFFFFF')
   .setFontWeight('bold')
   .setFontSize(11)
   .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  const widths = [160,200,150,220,100,220,150,150,60,190,200,200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

// ─────────────────────────────────────────────
function salvarNaPlanilha(data) {
  const sheet = getOrCreateSheet();
  const linha = [
    data.timestamp || new Date().toLocaleString('pt-BR'),
    data.nome || '',
    data.telefone || '',
    data.email || '',
    data.escolaridade || '',
    data.cep || '',
    data.rua || '',
    data.bairro || '',
    data.cidade || '',
    data.estado || '',
    data.vendas || '',
    data.computador || '',
    data.internet || '',
    data.produtos || ''
  ];
  sheet.appendRow(linha);
  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, HEADERS.length);
  rowRange.setBackground(lastRow % 2 === 0 ? '#E3F2FD' : '#FFFFFF');
  rowRange.setVerticalAlignment('middle');
  sheet.setRowHeight(lastRow, 28);
}

// ─────────────────────────────────────────────
function enviarEmailNotificacao(data) {
  const ts = data.timestamp || new Date().toLocaleString('pt-BR');
  const assunto = `Nova inscrição: ${data.nome || 'Candidato'} — DNC Saúde+`;

  const corpo = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f0f4f8;margin:0;padding:20px">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.1)">
    <div style="background:linear-gradient(135deg,#1a2980,#00BCD4);padding:24px 28px;text-align:center">
      <h2 style="color:#fff;margin:0;font-size:18px">🩺 DNC Saúde+ — Nova Inscrição</h2>
      <p style="color:rgba(255,255,255,0.85);margin:6px 0 0;font-size:13px">Processo Seletivo 2025</p>
    </div>
    <div style="padding:24px 28px">
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <tr style="background:#f8f9fa">
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980;width:45%">📅 Data de inscrição</td>
          <td style="padding:10px 12px;color:#333">${ts}</td>
        </tr>
        <tr>
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">👤 Nome</td>
          <td style="padding:10px 12px;color:#333">${data.nome || ''}</td>
        </tr>
        <tr style="background:#f8f9fa">
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">📱 WhatsApp</td>
          <td style="padding:10px 12px;color:#333">${data.telefone || ''}</td>
        </tr>
        <tr>
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">📧 E-mail</td>
          <td style="padding:10px 12px;color:#333">${data.email || ''}</td>
        </tr>
        <tr style="background:#f8f9fa">
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">🎓 Escolaridade</td>
          <td style="padding:10px 12px;color:#333">${data.escolaridade || ''}</td>
        </tr>
        <tr style="background:#f8f9fa">
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">📍 Endereço</td>
          <td style="padding:10px 12px;color:#333">${data.rua || ''}, ${data.bairro || ''} — ${data.cidade || ''}/${data.estado || ''} · CEP ${data.cep || ''}</td>
        </tr>
        <tr>
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">💼 Vendas</td>
          <td style="padding:10px 12px;color:#333">${data.vendas || ''}</td>
        </tr>
        <tr style="background:#f8f9fa">
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">💻 Computador</td>
          <td style="padding:10px 12px;color:#333">${data.computador || ''}</td>
        </tr>
        <tr>
          <td style="padding:10px 12px;font-weight:bold;color:#1a2980">🌐 Internet</td>
          <td style="padding:10px 12px;color:#333">${data.internet || '',
    data.produtos || ''}</td>
        </tr>
      </table>
    </div>
    <div style="background:#1a2980;padding:16px 28px;text-align:center">
      <p style="color:rgba(255,255,255,0.7);font-size:12px;margin:0">Este e-mail foi gerado automaticamente pelo sistema de inscrições DNC Saúde+</p>
    </div>
  </div>
</body>
</html>`;

  GmailApp.sendEmail(EMAIL_NOTIFICACAO, assunto, '', { htmlBody: corpo });
}

// ─────────────────────────────────────────────
//  Teste manual — execute para verificar
// ─────────────────────────────────────────────
function testeManual() {
  const dados = {
    timestamp: new Date().toLocaleString('pt-BR'),
    nome: 'Candidato Teste Silva',
    telefone: '(11) 99999-0000',
    escolaridade: 'Ensino Médio completo',
    cep: '06455-000',
    rua: 'Rua Teste, 100',
    bairro: 'Centro',
    cidade: 'Barueri',
    estado: 'SP',
    vendas: 'Sim',
    computador: 'Sim',
    internet: 'Sim'
  };
  salvarNaPlanilha(dados);
  enviarEmailNotificacao(dados);
  Logger.log('Teste concluído — verifique a planilha e o e-mail!');
}
