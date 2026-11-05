/**
 * Recebe as confirmações do convite e grava na planilha.
 * Serve para os dois convites: o campo "aba" diz em qual aba gravar (Mavi ou Kaue).
 *
 * Como instalar (uma vez só):
 * 1. Abra a planilha > Extensões > Apps Script.
 * 2. Apague o que tiver lá e cole este código. Salve.
 * 3. Implantar > Nova implantação > tipo "App da Web".
 *    Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * 4. Autorize, copie a URL que termina em /exec e cole em CONFIG.ENDPOINT no index.html.
 */
const CAMPOS = ['enviado_em', 'presenca', 'nome'];
const ABAS_PERMITIDAS = ['Mavi', 'Kaue'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const nomeAba = ABAS_PERMITIDAS.indexOf(p.aba) >= 0 ? p.aba : 'Mavi';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CAMPOS);
      sh.getRange(1, 1, 1, CAMPOS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    // o apóstrofo evita que o Sheets transforme texto em fórmula
    sh.appendRow(CAMPOS.map(c => {
      const v = String(p[c] || '').slice(0, 1000);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    }));
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Convite no ar ✔');
}
