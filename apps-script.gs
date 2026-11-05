/**
 * Recebe as confirmações dos dois convites.
 * Convite da Mavi -> aba "Mavi" | Convite do Kaue -> aba "Kaue"
 * Colunas: enviado_em | presenca | nome | Mavi | Kaue  (x na coluna do convite)
 *
 * Depois de colar este código: Implantar > Gerenciar implantações > lápis (editar)
 * > Versão: "Nova versão" > Implantar. Assim o link /exec continua o mesmo.
 */
const ABAS = ['Mavi', 'Kaue'];
const CABECALHO = ['enviado_em', 'presenca', 'nome', 'Mavi', 'Kaue'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const nomeAba = ABAS.indexOf(p.aba) >= 0 ? p.aba : 'Mavi';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);

    // garante o cabeçalho (sem apagar o que já existe)
    const topo = sh.getRange(1, 1, 1, CABECALHO.length);
    const atual = topo.getValues()[0];
    if (atual.some((v, i) => !v)) {
      topo.setValues([CABECALHO.map((c, i) => atual[i] || c)]);
      topo.setFontWeight('bold');
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
      nomeAba === 'Mavi' ? 'x' : '',
      nomeAba === 'Kaue' ? 'x' : ''
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Convite no ar ✔');
}
