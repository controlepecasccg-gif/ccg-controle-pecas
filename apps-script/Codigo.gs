/**
 * Controle de Peças · CCG — script da planilha "Cadastro de usuários"
 *
 * Planilha: Nome completo | Matrícula | E-mail  (o domínio do e-mail define o acesso:
 *   @responsavel.com = almoxarifado, @usuario.com = só pede peças). Coluna "Situação" é opcional.
 *
 * Faz duas coisas para o app (https://ccg-controle-pecas.web.app):
 *  1) entrega as linhas da planilha para o app criar os acessos e devolve a "Situação no app";
 *  2) recebe os termos de retirada (PDF) e guarda no Drive em:
 *     Termos › Empresa › Local › NF 0000
 *
 * Só aceita pedidos de quem está logado no app (o app manda o token do Firebase e o script confere).
 * Como instalar: veja apps-script/LEIA-ME.md
 */
const API_KEY = 'AIzaSyBUEkbbQ0ToFqbVSMOH7Q0jD1IQxl_gS6Q';           // chave pública do app (a mesma do site)
const PASTA_TERMOS_ID = '1qSdA0tOnD_FVCxxOUpqKwVmArAmPwNJ1';         // pasta "Termos" (criada pela Byanca)
const ADMINS = ['controlepecasccg@gmail.com'];                        // podem ler a planilha pelo app
const ABA = 'Usuários';

/* ---------------- planilha ---------------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Controle de Peças')
    .addItem('Preparar planilha (formatação e listas)', 'prepararPlanilha')
    .addToUi();
}
function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ABA) || ss.getSheets()[0];
}
function colunas_(sh) {
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''));
  const acha = (re) => cab.findIndex(h => re.test(h)) + 1;
  return { nome: acha(/^nome/), matricula: acha(/^matr/), email: acha(/^e-?mail/), situacao: acha(/^situa/) };
}
function prepararPlanilha() {
  const sh = aba_(); if (sh.getName() !== ABA) sh.setName(ABA);
  const c = colunas_(sh); const n = sh.getLastColumn(); const max = Math.max(sh.getMaxRows(), 300);
  if (sh.getMaxRows() < max) sh.insertRowsAfter(sh.getMaxRows(), max - sh.getMaxRows());
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#212B2A').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  if (c.nome) sh.setColumnWidth(c.nome, 280);
  if (c.matricula) { sh.setColumnWidth(c.matricula, 110); sh.getRange(2, c.matricula, max - 1, 1).setNumberFormat('@'); }
  if (c.email) {
    sh.setColumnWidth(c.email, 280);
    const col = sh.getRange(2, c.email).getA1Notation().replace(/\d+/, '');
    sh.getRange(2, c.email, max - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireFormulaSatisfied('=REGEXMATCH(LOWER(' + col + '2),"@(responsavel|usuario)\\.com$")')
      .setHelpText('Use nome@responsavel.com (almoxarifado) ou nome@usuario.com (só pede peças)').setAllowInvalid(false).build());
  }
  if (c.situacao) { sh.setColumnWidth(c.situacao, 300); sh.getRange(2, c.situacao, max - 1, 1).setBackground('#F2F4EE').setFontColor('#4E5752'); }
  SpreadsheetApp.getActiveSpreadsheet().toast('Planilha pronta.', 'Controle de Peças');
}

/* ---------------- app da web ---------------- */
function doGet() { return json_({ ok: true, app: 'Controle de Peças · CCG' }); }
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const email = quemE_(d.idToken);
    const admin = ADMINS.indexOf(email) >= 0;
    if (d.acao === 'ping') return json_({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName(), email: email });
    if (d.acao === 'usuarios') { if (!admin) throw new Error('só o escritório pode ler o cadastro'); return json_({ ok: true, linhas: linhas_() }); }
    if (d.acao === 'situacao') { if (!admin) throw new Error('só o escritório pode atualizar o cadastro'); return json_({ ok: true, n: gravarSituacao_(d.linhas || []) }); }
    if (d.acao === 'termo') { if (/@usuario\.com$/.test(email)) throw new Error('sem permissão para salvar termos'); return json_({ ok: true, arquivos: salvarTermo_(d) }); }
    throw new Error('ação desconhecida');
  } catch (err) { return json_({ ok: false, erro: String(err.message || err) }); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function quemE_(idToken) { // confere o login do app com o Firebase
  if (!idToken) throw new Error('sem login');
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + API_KEY,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText() || '{}');
  if (r.getResponseCode() !== 200 || !j.users || !j.users[0]) throw new Error('login expirado, entre de novo no app');
  return String(j.users[0].email || '').toLowerCase();
}
function linhas_() {
  const sh = aba_(); const c = colunas_(sh); const n = sh.getLastRow(); if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getDisplayValues();
  const g = (row, col) => col ? String(row[col - 1] || '').trim() : '';
  return v.map((row, i) => ({ linha: i + 2, nome: g(row, c.nome), matricula: g(row, c.matricula), email: g(row, c.email).toLowerCase() })).filter(l => l.nome && l.email);
}
function gravarSituacao_(lista) {
  const sh = aba_(); const c = colunas_(sh); let n = 0;
  if (!c.situacao) return 0; // a coluna "Situação" é opcional
  lista.forEach(l => { if (!l.linha) return; sh.getRange(l.linha, c.situacao).setValue(l.situacao || ''); n++; });
  return n;
}

/* ---------------- termos no Drive ---------------- */
function pasta_(mae, nome) {
  nome = String(nome || 'Sem nome').replace(/[\\/]/g, '-').trim() || 'Sem nome';
  const it = mae.getFoldersByName(nome); return it.hasNext() ? it.next() : mae.createFolder(nome);
}
function salvarTermo_(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const termos = DriveApp.getFolderById(PASTA_TERMOS_ID); const out = [];
    (d.pastas || []).forEach(p => {
      const nf = pasta_(pasta_(pasta_(termos, p.empresa), p.local), p.nf);
      (d.arquivos || []).forEach(a => {
        const ant = nf.getFilesByName(a.nome); while (ant.hasNext()) ant.next().setTrashed(true); // termo refeito substitui o anterior
        const f = nf.createFile(Utilities.newBlob(Utilities.base64Decode(a.pdf), 'application/pdf', a.nome));
        f.setDescription('Pedido ' + (d.pedido || '') + ' · gerado pelo app Controle de Peças');
        out.push(f.getUrl());
      });
    });
    return out;
  } finally { lock.releaseLock(); }
}
