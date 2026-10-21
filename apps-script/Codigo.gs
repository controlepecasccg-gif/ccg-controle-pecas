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
const PROJETO = 'ccg-controle-pecas';
const FS = 'https://firestore.googleapis.com/v1/projects/' + PROJETO + '/databases/(default)/documents';
const ROBO_EMAIL = 'sincronizador@robo.controlepecas';   // conta interna que o script usa para gravar no app

/* ---------------- planilha ---------------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Controle de Peças')
    .addItem('Preparar planilha (formatação e listas)', 'prepararPlanilha')
    .addItem('Ligar sincronização automática (a cada 10 min)', 'ligarSincronizacao')
    .addItem('Sincronizar agora', 'sincronizarUsuarios')
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

/* ---------------- sincronização automática (planilha → acessos do app) ----------------
   A cada 10 minutos: cria o acesso de quem entrou na planilha (senha inicial Gavea + matrícula),
   atualiza nome e matrícula e desativa quem saiu. Escreve o resultado na coluna "Situação", se existir. */
function ligarSincronizacao() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sincronizarUsuarios').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarUsuarios').timeBased().everyMinutes(10).create();
  sincronizarUsuarios();
  try { SpreadsheetApp.getActiveSpreadsheet().toast('Sincronização automática ligada: a cada 10 minutos.', 'Controle de Peças'); } catch (e) {}
}
function idt_(url, corpo) {
  const r = UrlFetchApp.fetch(url + '?key=' + API_KEY, { method: 'post', contentType: 'application/json', payload: JSON.stringify(corpo), muteHttpExceptions: true });
  return { ok: r.getResponseCode() === 200, j: JSON.parse(r.getContentText() || '{}') };
}
function tokenRobo_() {
  const pr = PropertiesService.getScriptProperties(); let senha = pr.getProperty('ROBO_SENHA');
  if (!senha) {
    senha = Utilities.getUuid() + 'Aa1!'; const c = idt_('https://identitytoolkit.googleapis.com/v1/accounts:signUp', { email: ROBO_EMAIL, password: senha, returnSecureToken: true });
    if (!c.ok) throw new Error('não consegui criar a conta interna: ' + JSON.stringify(c.j.error || c.j));
    pr.setProperty('ROBO_SENHA', senha); return c.j.idToken;
  }
  const r = idt_('https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword', { email: ROBO_EMAIL, password: senha, returnSecureToken: true });
  if (!r.ok) throw new Error('não consegui entrar com a conta interna: ' + JSON.stringify(r.j.error || r.j));
  return r.j.idToken;
}
function fsVal_(v) { return typeof v === 'boolean' ? { booleanValue: v } : { stringValue: String(v == null ? '' : v) }; }
function fsDoc_(d) { const f = {}; (d.fields ? Object.keys(d.fields) : []).forEach(k => { const x = d.fields[k]; f[k] = 'booleanValue' in x ? x.booleanValue : 'stringValue' in x ? x.stringValue : 'integerValue' in x ? Number(x.integerValue) : 'doubleValue' in x ? x.doubleValue : null; }); return f; }
function usuariosApp_(tk) {
  const out = {}; let pg = '';
  do {
    const r = UrlFetchApp.fetch(FS + '/usuarios?pageSize=300' + (pg ? '&pageToken=' + pg : ''), { headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('não consegui ler os usuários do app: ' + r.getContentText().slice(0, 200));
    const j = JSON.parse(r.getContentText() || '{}'); (j.documents || []).forEach(d => { out[d.name.split('/').pop()] = fsDoc_(d); }); pg = j.nextPageToken || '';
  } while (pg);
  return out;
}
function gravarUsuario_(tk, id, dados) {
  const campos = Object.keys(dados); const fields = {}; campos.forEach(k => fields[k] = fsVal_(dados[k]));
  const url = FS + '/usuarios/' + encodeURIComponent(id) + '?' + campos.map(k => 'updateMask.fieldPaths=' + k).join('&');
  const r = UrlFetchApp.fetch(url, { method: 'patch', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, payload: JSON.stringify({ fields }), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('falha ao gravar no app (' + r.getResponseCode() + ')');
}
function perfilDoEmail_(e) { return /@responsavel\.com$/.test(e) ? 'responsavel' : /@usuario\.com$/.test(e) ? 'usuario' : ''; }
function hoje_() { return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy'); }
function sincronizarUsuarios() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    const tk = tokenRobo_(); const app = usuariosApp_(tk); const linhas = linhas_(); const volta = []; const naPlanilha = {};
    linhas.forEach(l => {
      const email = l.email; const id = email.replace(/[^a-z0-9]/g, '_'); const perfil = perfilDoEmail_(email); naPlanilha[id] = true;
      if (!perfil) { volta.push({ linha: l.linha, situacao: 'Erro: o e-mail precisa terminar em @responsavel.com ou @usuario.com' }); return; }
      const ja = app[id]; const mat = String(l.matricula || '').replace(/\s+/g, '');
      try {
        if (!ja || ja.ativo === false) {
          if (!ja) {
            if (!mat) { volta.push({ linha: l.linha, situacao: 'Erro: falta a matrícula (ela forma a senha do primeiro acesso)' }); return; }
            const c = idt_('https://identitytoolkit.googleapis.com/v1/accounts:signUp', { email: email, password: 'Gavea' + mat, returnSecureToken: false });
            const m = (c.j.error && c.j.error.message) || '';
            if (!c.ok && !/EMAIL_EXISTS/.test(m)) { volta.push({ linha: l.linha, situacao: 'Erro: ' + (/INVALID_EMAIL/.test(m) ? 'e-mail inválido' : /WEAK_PASSWORD/.test(m) ? 'matrícula curta demais para a senha' : m) }); return; }
          }
          gravarUsuario_(tk, id, { id: id, email: email, nome: l.nome, matricula: mat, perfil: perfil, ativo: true });
          volta.push({ linha: l.linha, situacao: (ja ? 'Reativado em ' : 'Acesso criado em ') + hoje_() + ' · senha inicial Gavea' + mat });
        } else {
          const muda = {}; if (!ja.senhaRedefinida && ja.nome !== l.nome) muda.nome = l.nome; if (mat && ja.matricula !== mat) muda.matricula = mat; if (ja.perfil !== perfil) muda.perfil = perfil;
          if (Object.keys(muda).length) gravarUsuario_(tk, id, muda);
          volta.push({ linha: l.linha, situacao: ja.senhaRedefinida ? 'Ativo · já criou a própria senha' : 'Ativo · ainda não fez o primeiro acesso (senha Gavea' + mat + ')' });
        }
      } catch (err) { volta.push({ linha: l.linha, situacao: 'Erro: ' + err.message }); }
    });
    Object.keys(app).forEach(id => { const u = app[id]; if (perfilDoEmail_(u.email || '') && u.ativo !== false && !naPlanilha[id]) gravarUsuario_(tk, id, { ativo: false }); });
    gravarSituacao_(volta);
    PropertiesService.getScriptProperties().setProperty('ULTIMA_SINC', new Date().toISOString());
  } finally { lock.releaseLock(); }
}
