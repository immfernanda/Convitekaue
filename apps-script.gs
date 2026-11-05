/**
 * Recebe as confirmações dos dois convites (Mavi e Kaue) e grava tudo na MESMA aba.
 * Colunas: enviado_em | presenca | nome | Mavi | Kaue
 * Um "x" na coluna Mavi ou Kaue mostra de qual convite a pessoa veio.
 *
 * Depois de colar este código: Implantar > Gerenciar implantações > lápis (editar)
 * > Versão: "Nova versão" > Implantar. Assim o link /exec continua o mesmo.
 */
const NOME_ABA = 'Aniversáriantes';
const CABECALHO = ['enviado_em', 'presenca', 'nome', 'Mavi', 'Kaue'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
    if (sh.getLastRow() === 0) {
      sh.appendRow(CABECALHO);
      sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const limpa = v => {
      v = String(v || '').slice(0, 1000);
      return /^[=+\-@]/.test(v) ? "'" + v : v;   // evita virar fórmula
    };
    sh.appendRow([
      limpa(p.enviado_em),
      limpa(p.presenca),
      limpa(p.nome),
      p.aba === 'Mavi' ? 'x' : '',
      p.aba === 'Kaue' ? 'x' : ''
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Convite no ar ✔');
}
