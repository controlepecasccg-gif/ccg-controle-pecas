/**
 * Controle de Peças · CCG — script da planilha "Cadastro de usuários"
 *
 * Faz duas coisas para o app (https://ccg-controle-pecas.web.app):
 *  1) entrega as linhas da planilha para o app criar os acessos e devolve a "Situação no app";
 *  2) recebe os termos de retirada (PDF) e guarda no Drive em:
 *     CCG · Controle de Peças › Termos › Empresa › Local › NF 0000
 *
 * Só aceita pedidos de quem está logado no app (o app manda o token do Firebase e o script confere).
 * Como instalar: veja apps-script/LEIA-ME.md
 */
const API_KEY = 'AIzaSyBUEkbbQ0ToFqbVSMOH7Q0jD1IQxl_gS6Q';           // chave pública do app (a mesma do site)
const PASTA_RAIZ_ID = '19WQ9woYV5OlE3elpW4pKEkUe2HeljceY';           // pasta "CCG · Controle de Peças"
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
  return { nome: acha(/^nome/), matricula: acha(/^matr/), cargo: acha(/^cargo/), setor: acha(/^setor/), tipo: acha(/^tipo/), login: acha(/^login/),
           email: acha(/^e-?mail/), senha: acha(/^senha/), frente: acha(/^frente/), ativo: acha(/^ativo/), situacao: acha(/^situa/) };
}
function prepararPlanilha() {
  const sh = aba_(); if (sh.getName() !== ABA) sh.setName(ABA);
  const c = colunas_(sh); const n = sh.getLastColumn(); const max = Math.max(sh.getMaxRows(), 300);
  if (sh.getMaxRows() < max) sh.insertRowsAfter(sh.getMaxRows(), max - sh.getMaxRows());
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#212B2A').setWrap(true).setVerticalAlignment('middle');
  sh.setRowHeight(1, 42);
  sh.getRange(2, 1, max - 1, n).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  const lista = (col, itens) => { if (col) sh.getRange(2, col, max - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(itens, true).setAllowInvalid(false).build()); };
  lista(c.tipo, ['Responsável', 'Usuário']); lista(c.ativo, ['Sim', 'Não']);
  [c.email, c.situacao].forEach(col => { if (col) sh.getRange(2, col, max - 1, 1).setBackground('#F2F4EE').setFontColor('#4E5752'); });
  if (c.matricula) sh.getRange(2, c.matricula, max - 1, 1).setNumberFormat('@');
  if (c.senha) sh.getRange(2, c.senha, max - 1, 1).setNumberFormat('@');
  const larg = { nome: 240, matricula: 90, cargo: 160, setor: 130, tipo: 130, login: 140, email: 230, senha: 110, frente: 170, ativo: 70, situacao: 230 };
  Object.keys(larg).forEach(k => { if (c[k]) sh.setColumnWidth(c[k], larg[k]); });
  for (let r = 2; r <= sh.getLastRow(); r++) preencherEmail_(sh, c, r);
  SpreadsheetApp.getActiveSpreadsheet().toast('Planilha pronta.', 'Controle de Peças');
}
function onEdit(e) {
  const sh = e.range.getSheet(); if (sh.getName() !== ABA || e.range.getRow() < 2) return;
  const c = colunas_(sh);
  for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) preencherEmail_(sh, c, r);
}
function emailDe_(login, tipo) {
  login = String(login || '').trim().toLowerCase().replace(/@.*$/, '').replace(/\s+/g, '.');
  if (!login) return '';
  return login + '@' + (/resp/i.test(String(tipo || '')) ? 'responsavel.com' : 'usuario.com');
}
function preencherEmail_(sh, c, r) {
  if (!c.email || !c.login) return;
  const login = sh.getRange(r, c.login).getValue(); const tipo = c.tipo ? sh.getRange(r, c.tipo).getValue() : '';
  sh.getRange(r, c.email).setValue(emailDe_(login, tipo));
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
  return v.map((row, i) => ({ linha: i + 2, nome: g(row, c.nome), matricula: g(row, c.matricula), cargo: g(row, c.cargo), setor: g(row, c.setor), tipo: g(row, c.tipo),
    login: g(row, c.login), senha: g(row, c.senha), frente: g(row, c.frente), ativo: g(row, c.ativo) || 'Sim' })).filter(l => l.nome && l.login);
}
function gravarSituacao_(lista) {
  const sh = aba_(); const c = colunas_(sh); let n = 0;
  lista.forEach(l => { if (!l.linha) return; if (c.email && l.email) sh.getRange(l.linha, c.email).setValue(l.email); if (c.situacao) sh.getRange(l.linha, c.situacao).setValue(l.situacao || ''); n++; });
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
    const termos = pasta_(DriveApp.getFolderById(PASTA_RAIZ_ID), 'Termos'); const out = [];
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
