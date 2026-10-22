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
    .addSeparator()
    .addItem('Sincronizar usuários agora', 'sincronizarUsuariosAgora')
    .addItem('Enviar termos e relatórios ao Drive agora', 'enviarTermosAgora')
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
  ScriptApp.getProjectTriggers().filter(t => ['sincronizarUsuarios', 'rodarAutomatico'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('rodarAutomatico').timeBased().everyMinutes(10).create();
  rodarAutomatico();
  try { SpreadsheetApp.getActiveSpreadsheet().toast('Ligado: usuários e termos a cada 10 minutos.', 'Controle de Peças'); } catch (e) {}
}
function rodarAutomatico() {
  try { sincronizarUsuarios(); } catch (e) { console.error('usuários: ' + e.message); }
  try { enviarTermos(); } catch (e) { console.error('termos: ' + e.message); }
  try { enviarRelatorios(); } catch (e) { console.error('relatórios: ' + e.message); }
  try { enviarRecebimentos(); } catch (e) { console.error('recebimentos: ' + e.message); }
  try { enviarMensais(); } catch (e) { console.error('mensais: ' + e.message); }
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

/* ---------------- termos no Drive (gerados aqui, a partir do que o app gravou) ----------------
   O app marca o pedido com drivePendente = true quando registra uma retirada ou devolução.
   A cada 10 minutos (ou pelo menu) o script monta os dois termos em PDF e guarda em Termos › Empresa › Local › NF.
   Numa devolução o termo é refeito e substitui o anterior. */
function enviarTermosAgora() {
  const ui = SpreadsheetApp.getUi(); const msg = [];
  try { const t = enviarTermos(true); msg.push(t.lidos ? 'Termos: ' + t.enviados + ' de ' + t.lidos + ' enviados.' : 'Termos: nenhum pendente.'); if (t.erros.length) msg.push('Erros nos termos:\n' + t.erros.join('\n')); }
  catch (e) { msg.push('Termos: não consegui enviar (' + e.message + ')'); }
  [['Coleta', enviarRelatorios], ['Recebimento no canteiro', enviarRecebimentos], ['Conferência mensal', enviarMensais]].forEach(([nome, fn]) => {
    try { const r = fn(true); msg.push(r.lidos ? nome + ': ' + r.enviados + ' de ' + r.lidos + ' enviados.' : nome + ': nada pendente.'); if (r.erros.length) msg.push('Erros (' + nome + '):\n' + r.erros.join('\n')); }
    catch (e) { msg.push(nome + ': não consegui enviar (' + e.message + ')'); } });
  ui.alert('Drive', msg.join('\n\n'), ui.ButtonSet.OK);
}
function fsLer_(v) {
  if (!v) return null;
  if ('stringValue' in v) return v.stringValue; if ('booleanValue' in v) return v.booleanValue;
  if ('integerValue' in v) return Number(v.integerValue); if ('doubleValue' in v) return v.doubleValue;
  if ('nullValue' in v) return null; if ('timestampValue' in v) return v.timestampValue;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fsLer_);
  if ('mapValue' in v) { const o = {}; const f = v.mapValue.fields || {}; Object.keys(f).forEach(k => o[k] = fsLer_(f[k])); return o; }
  return null;
}
function fsEsc_(x) {
  if (x === null || x === undefined) return { nullValue: null };
  if (typeof x === 'boolean') return { booleanValue: x };
  if (typeof x === 'number') return Number.isInteger(x) ? { integerValue: String(x) } : { doubleValue: x };
  if (Array.isArray(x)) return { arrayValue: { values: x.map(fsEsc_) } };
  if (typeof x === 'object') { const f = {}; Object.keys(x).forEach(k => f[k] = fsEsc_(x[k])); return { mapValue: { fields: f } }; }
  return { stringValue: String(x) };
}
function pendentesDrive_(tk, colecao) {
  const r = UrlFetchApp.fetch(FS + ':runQuery', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: colecao || 'remessas' }], where: { fieldFilter: { field: { fieldPath: 'drive.ok' }, op: 'EQUAL', value: { booleanValue: false } } }, limit: 15 } }) });
  if (r.getResponseCode() !== 200) throw new Error('não consegui ler os pedidos: ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText() || '[]').filter(x => x.document).map(x => { const o = {}; const f = x.document.fields || {}; Object.keys(f).forEach(k => o[k] = fsLer_(f[k])); o._id = x.document.name.split('/').pop(); return o; });
}
function marcarDrive_(tk, id, drive, colecao) {
  const url = FS + '/' + (colecao || 'remessas') + '/' + encodeURIComponent(id) + '?updateMask.fieldPaths=drive&updateMask.fieldPaths=drivePendente';
  UrlFetchApp.fetch(url, { method: 'patch', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: JSON.stringify({ fields: { drive: fsEsc_(drive), drivePendente: { booleanValue: !drive.ok } } }) });
}
function enviarTermos(detalhe) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) { if (detalhe) throw new Error('outra sincronização está rodando, tente em 1 minuto'); return 0; }
  let n = 0, lidos = 0; const erros = [];
  try {
    const tk = tokenRobo_(); const raiz = DriveApp.getFolderById(PASTA_TERMOS_ID);
    const lista = pendentesDrive_(tk); lidos = lista.length;
    lista.forEach(r => {
      if (!r.termo) { marcarDrive_(tk, r._id, { ok: true, versao: 0, em: new Date().toISOString(), obs: 'sem termo' }); return; }
      const versao = (r.devolucoes || []).length;
      try {
        const arquivos = ['colab', 'resp'].map(via => ({ nome: nomeTermo_(r, via), blob: pdfTermo_(r, via) }));
        const pastas = (r.termo.pastas && r.termo.pastas.length) ? r.termo.pastas : [{ empresa: 'Sem nota', local: r.destino, nf: 'Sem nota' }];
        pastas.forEach(p => {
          const pasta = pasta_(pasta_(pasta_(pasta_(raiz, p.empresa), p.local), p.nf), '3 · Termos');
          arquivos.forEach(a => { const ant = pasta.getFilesByName(a.nome); while (ant.hasNext()) ant.next().setTrashed(true); pasta.createFile(a.blob.copyBlob().setName(a.nome)).setDescription('Pedido ' + (r.numero || '') + ' · gerado pelo Controle de Peças'); });
        });
        marcarDrive_(tk, r._id, { ok: true, versao: versao, em: new Date().toISOString() }); n++;
      } catch (err) { erros.push((r.numero || r._id) + ': ' + err.message); marcarDrive_(tk, r._id, { ok: false, erro: String(err.message || err).slice(0, 120), em: new Date().toISOString() }); }
    });
  } finally { lock.releaseLock(); }
  return detalhe ? { enviados: n, lidos: lidos, erros: erros } : n;
}
function dataBR_(iso, fmt) { return Utilities.formatDate(new Date(iso), 'America/Sao_Paulo', fmt || 'dd/MM/yyyy'); }
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function semAcento_(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function nomeTermo_(r, via) {
  const c = r.termo.colaborador || {};
  return semAcento_(dataBR_(r.termo.em, 'yyyy-MM-dd') + ' - ' + r.numero + ' - ' + String(c.nome || '').toUpperCase() + (c.matricula ? ' (' + c.matricula + ')' : '') + ' - Termo ' + (via === 'colab' ? 'FOR-SUP-001 via colaborador' : 'FOR-SUP-002 via responsavel') + '.pdf').replace(/[\\/:*?"<>|]/g, '-').replace(/[^\x20-\x7E]/g, '');
}
function linhasTermo_(r) { // uma linha por peça e nota fiscal; cada devolução vira linha com data e visto
  const out = [];
  (r.termo.itens || []).forEach(i => {
    const desc = String(i.desc || i.key).toUpperCase();
    const segs = (i.porNF && i.porNF.length ? i.porNF : [{ nf: '', q: Number(i.qtd) || 0 }]).map(x => ({ nf: x.nf || '', resta: Number(x.q) || 0 }));
    (r.devolucoes || []).forEach(d => { let q = Number((d.itens || {})[i.key]) || 0;
      segs.forEach(sg => { if (q <= 0 || sg.resta <= 0) return; const x = Math.min(q, sg.resta); out.push({ qtd: x, desc: desc, nf: sg.nf, dev: d }); sg.resta -= x; q -= x; }); });
    segs.forEach(sg => { if (sg.resta > 0) out.push({ qtd: sg.resta, desc: desc, nf: sg.nf, dev: null }); });
  });
  return out;
}
function pdfTermo_(r, via) {
  const t = r.termo, c = t.colaborador || {}, dRet = dataBR_(t.em), linhas = linhasTermo_(r);
  while (linhas.length < 20) linhas.push(null);
  const B = 'border:1px solid #13152B;';
  const th = 'style="' + B + 'background:#E6EAE0;font-size:7.5pt;font-weight:bold;text-align:center;padding:3px"';
  const td = (x, extra) => '<td style="' + B + 'font-size:8pt;padding:2px 4px;height:19px;' + (extra || '') + '">' + (x == null ? '' : x) + '</td>';
  const img = (src, h) => src ? '<img src="' + src + '" style="height:' + h + 'px">' : '';
  const dd = (d) => d ? dataBR_(d.em) : '';
  const cod = via === 'colab' ? 'FOR_SUP-001' : 'FOR_SUP-002';
  const titulo = via === 'colab' ? 'CONTROLE DE FERRAMENTAS/EQUIPAMENTOS - COLABORADOR' : 'CONTROLE DE FERRAMENTAS/EQUIPAMENTOS - DIÁRIA';
  let h = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm 12mm 10mm}body{font-family:Arial,sans-serif;color:#13152B;margin:0}table{border-collapse:collapse;width:100%}</style></head><body>';
  h += '<table><tr><td style="' + B + 'width:20%;text-align:center;padding:6px"><img src="data:image/jpeg;base64,' + LOGO_CCG + '" style="height:34px"></td><td style="' + B + 'text-align:center;font-weight:bold;font-size:11pt;padding:6px">' + titulo + '</td><td style="' + B + 'width:28%;padding:0"><table style="font-size:7pt"><tr><td style="padding:3px 5px;border-bottom:1px solid #13152B;border-right:1px solid #13152B"><b>TIPO DE DOCUMENTO:</b><br>FORMULÁRIO</td><td style="padding:3px 5px;border-bottom:1px solid #13152B"><b>CÓDIGO:</b><br>' + cod + '</td></tr><tr><td style="padding:3px 5px;border-right:1px solid #13152B"><b>REVISÃO:</b><br>01</td><td style="padding:3px 5px"><b>DATA:</b><br>26/06/2023</td></tr></table></td></tr></table>';
  const lb = (l, v, w) => '<td style="' + B + 'font-size:8pt;padding:2px 4px;height:24px;' + (w ? 'width:' + w + ';' : '') + '"><b style="font-size:6.5pt">' + l + '</b><br>' + esc_(v) + '</td>';
  if (via === 'colab') {
    h += '<table style="margin-top:5px"><tr>' + lb('NOME DO FUNCIONÁRIO:', c.nome, '48%') + lb('MATRÍCULA:', c.matricula, '16%') + lb('OBRA:', t.obra) + '</tr><tr>' + lb('CARGO/FUNÇÃO:', c.cargo) + lb('RESPONSÁVEL DO SETOR:', String(t.responsavelSetor || '').toUpperCase()) + lb('SETOR:', String(c.setor || '').toUpperCase()) + '</tr></table>';
    h += '<table style="margin-top:6px"><tr><th ' + th + ' width="8%">QUANTIDADE</th><th ' + th + '>DESCRIÇÃO DO ITEM</th><th ' + th + ' width="11%">Nº DA NOTA FISCAL</th><th ' + th + ' width="11%">DATA DA RETIRADA</th><th ' + th + ' width="12%">VISTO DO INTEGRANTE</th><th ' + th + ' width="11%">DATA DA DEVOLUÇÃO</th><th ' + th + ' width="12%">VISTO DO ALMOXARIFADO</th></tr>';
    linhas.forEach(l => { h += l ? '<tr>' + td(l.qtd, 'text-align:center') + td(esc_(l.desc)) + td(esc_(l.nf), 'text-align:center') + td(dRet, 'text-align:center') + td(img(t.rubrica, 16), 'text-align:center;padding:0') + td(dd(l.dev), 'text-align:center') + td(l.dev ? img(l.dev.visto, 16) : '', 'text-align:center;padding:0') + '</tr>' : '<tr>' + td('') + td('') + td('') + td('') + td('') + td('') + td('') + '</tr>'; });
  } else {
    h += '<table style="margin-top:6px"><tr><th ' + th + ' rowspan="2" width="7%">QUANT.</th><th ' + th + ' rowspan="2">DESCRIÇÃO DO ITEM</th><th ' + th + ' rowspan="2" width="8%">Nº DA NOTA</th><th ' + th + ' colspan="4">RETIRADA</th><th ' + th + ' colspan="2">DEVOLUÇÃO</th></tr><tr><th ' + th + ' width="8.5%">DATA</th><th ' + th + ' width="15%">NOME / MATRÍCULA</th><th ' + th + ' width="9.5%">VISTO DO INTEGRANTE</th><th ' + th + ' width="6.5%">OBRA</th><th ' + th + ' width="8.5%">DATA</th><th ' + th + ' width="9.5%">VISTO DO ALMOX.</th></tr>';
    const nm = esc_(c.nome) + (c.matricula ? ' / ' + esc_(c.matricula) : '');
    linhas.forEach(l => { h += l ? '<tr>' + td(l.qtd, 'text-align:center') + td(esc_(l.desc), 'font-size:7pt') + td(esc_(l.nf), 'text-align:center;font-size:7pt') + td(dRet, 'text-align:center;font-size:7pt') + td(nm, 'font-size:6.5pt') + td(img(t.rubrica, 16), 'text-align:center;padding:0') + td('GÁVEA', 'text-align:center;font-size:6.5pt') + td(dd(l.dev), 'text-align:center;font-size:7pt') + td(l.dev ? img(l.dev.visto, 16) : '', 'text-align:center;padding:0') + '</tr>' : '<tr>' + td('') + td('') + td('') + td('') + td('') + td('') + td('') + td('') + td('') + '</tr>'; });
  }
  h += '</table>';
  h += '<div style="' + B + 'margin-top:6px;padding:5px 8px;font-size:7.5pt;line-height:1.4"><div style="text-align:center;font-weight:bold;margin-bottom:2px">DECLARAÇÃO DO RESPONSÁVEL</div>Como signatário do presente Termo de Responsabilidade, declaro para os devidos fins, que os bens aqui relacionados foram inventariados e localizados fisicamente, encontrando-se em bom estado de conservação. Reconhecendo a exatidão das informações aqui prestadas comprometo-me pela guarda e proteção desses bens, conservando-os adequadamente ficando assim, responsável pelo ônus da não observância destes cuidados. Autorizando desde já a empregadora o desconto em folha de pagamento do valor dos itens que por culpa ou dolo, sejam extraviados ou danificados.</div>';
  const ass = (src, nome, papel) => '<td style="width:50%;text-align:center;vertical-align:bottom;padding:0 24px"><div style="height:44px">' + img(src, 42) + '</div><div style="border-top:1px solid #13152B;font-size:8pt;font-weight:bold;padding-top:2px">' + esc_(nome) + '</div><div style="font-size:7pt;color:#4E5752">' + papel + '</div></td>';
  h += '<table style="margin-top:28px"><tr>' + ass(t.assColab, (c.nome || '') + (c.matricula ? ' · MAT. ' + c.matricula : ''), 'COLABORADOR') + ass(t.assResp, String(t.respNome || '').toUpperCase() + (t.respMat ? ' · MAT. ' + t.respMat : ''), 'RESPONSÁVEL PELO TERMO') + '</tr></table>';
  h += '<div style="margin-top:14px;border-top:1px solid #D7DBD0;padding-top:3px;font-size:6.5pt;color:#8A9289">' + (via === 'colab' ? 'VIA DO COLABORADOR' : 'VIA DO RESPONSÁVEL') + ' · ' + esc_(r.numero) + ' · ' + esc_(r.destino) + (t.nfsTxt ? ' · ' + esc_(t.nfsTxt) : '') + (r.status === 'quitado' ? ' · TERMO QUITADO' : '') + ' · gerado em ' + dataBR_(new Date().toISOString(), 'dd/MM/yyyy HH:mm') + '</div>';
  h += '</body></html>';
  return Utilities.newBlob(h, 'text/html', 'termo.html').getAs('application/pdf');
}

// logo da CCG (usada no cabeçalho do termo)
const LOGO_CCG = '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';

function sincronizarUsuariosAgora() {
  const ui = SpreadsheetApp.getUi();
  try { sincronizarUsuarios(); ui.alert('Usuários', 'Sincronizado. ' + (colunas_(aba_()).situacao ? 'Veja a coluna "Situação".' : 'Crie uma coluna "Situação" para ver o resultado de cada linha.'), ui.ButtonSet.OK); }
  catch (e) { ui.alert('Usuários', 'Não consegui sincronizar: ' + e.message, ui.ButtonSet.OK); }
}

/* ---------------- relatórios de conferência no Drive ----------------
   Cada viagem fechada, encerramento ou correção marca a nota (drive.ok = false). O script gera o PDF
   do relatório e guarda na pasta da nota: Termos › Empresa › Local › NF 0000, junto com os termos.
   Revisões antigas são refeitas com a marca OBSOLETA no nome e no documento. */
function grupoEmpresa_(f) { f = String(f || '').trim(); if (/mills/i.test(f)) return 'Mills'; if (/^sh\b|sh f[oô]rmas/i.test(f)) return 'SH'; if (/^peri\b/i.test(f)) return 'Peri'; return f || 'Sem empresa'; }
function brl_(v) { return 'R$ ' + (Number(v) || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function nfmt_(v) { return String(Math.round(Number(v) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function enviarRelatorios(detalhe) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) { if (detalhe) throw new Error('outra sincronização está rodando, tente em 1 minuto'); return 0; }
  let n = 0, lidos = 0; const erros = [];
  try {
    const tk = tokenRobo_(); const raiz = DriveApp.getFolderById(PASTA_TERMOS_ID);
    const lista = pendentesDrive_(tk, 'notas'); lidos = lista.length;
    lista.forEach(nota => {
      try {
        if (['em_conferencia', 'concluida'].indexOf(nota.status) < 0) { marcarDrive_(tk, nota._id, { ok: true, em: new Date().toISOString(), obs: 'sem conferência' }, 'notas'); return; }
        const emp = grupoEmpresa_(nota.fornecedor);
        const pasta = pasta_(pasta_(pasta_(pasta_(raiz, emp), nota.local || 'Sem local'), 'NF ' + (nota.numero || 's-n')), '1 · Coleta');
        const base = 'Coleta NF ' + (nota.numero || '') + ' - ' + emp;
        const salvar = (nome, blob) => { const ant = pasta.getFilesByName(nome); while (ant.hasNext()) ant.next().setTrashed(true); pasta.createFile(blob.setName(nome)).setDescription('Relatório de conferência · gerado pelo Controle de Peças'); };
        (nota.revisoes || []).forEach(v => { // versões anteriores: marcadas como obsoletas
          const velha = Object.assign({}, nota, { viagens: v.viagens || [], contagem: {}, assinaturas: v.assinaturas || [], status: 'concluida', concluidaEm: v.concluidaEm, conferente: v.conferente, revisao: v.rev, obsoleta: true, substituidaEm: v.substituidaEm, substituidaPor: v.substituidaPor, motivoRevisao: v.motivo });
          const ant = pasta.getFilesByName(base + ' - rev' + v.rev + '.pdf'); while (ant.hasNext()) ant.next().setTrashed(true);
          salvar(semAcento_(base + ' - rev' + v.rev + ' - OBSOLETA.pdf'), pdfRelatorio_(velha));
        });
        salvar(semAcento_(base + ' - rev' + (nota.revisao || 1) + '.pdf'), pdfRelatorio_(nota));
        marcarDrive_(tk, nota._id, { ok: true, em: new Date().toISOString(), revisao: nota.revisao || 1 }, 'notas'); n++;
      } catch (err) { erros.push('NF ' + (nota.numero || nota._id) + ': ' + err.message); marcarDrive_(tk, nota._id, { ok: false, erro: String(err.message || err).slice(0, 120), em: new Date().toISOString() }, 'notas'); }
    });
  } finally { lock.releaseLock(); }
  return detalhe ? { enviados: n, lidos: lidos, erros: erros } : n;
}
function pdfRelatorio_(nt) {
  const concluida = nt.status === 'concluida'; const rev = nt.revisao || 1;
  const rec = {}; (nt.viagens || []).forEach(v => Object.keys(v.recebido || {}).forEach(k => rec[k] = (rec[k] || 0) + (Number(v.recebido[k]) || 0)));
  if (concluida) Object.keys(nt.contagem || {}).forEach(k => rec[k] = (rec[k] || 0) + (Number(nt.contagem[k]) || 0));
  const itens = (nt.itens || []).map(i => { const r = rec[i.key] || 0; return { codigo: i.codigo || '', desc: i.desc || i.key, qtd: Number(i.qtd) || 0, rec: r, dif: r - (Number(i.qtd) || 0), vu: Number(i.vu) || 0 }; });
  const totQ = itens.reduce((t, i) => t + i.qtd, 0), totR = itens.reduce((t, i) => t + i.rec, 0);
  const faltas = itens.filter(i => i.dif < 0), sobras = itens.filter(i => i.dif > 0);
  const faltaQ = faltas.reduce((t, i) => t - i.dif, 0), faltaV = faltas.reduce((t, i) => t - i.dif * i.vu, 0), sobraQ = sobras.reduce((t, i) => t + i.dif, 0);
  let viagens = (nt.viagens || []).map(v => ({ rot: v.correcao ? 'Correção (rev. ' + v.correcao + ')' : 'Viagem ' + v.n, em: v.em, por: v.por || '', q: Object.keys(v.recebido || {}).reduce((t, k) => t + (Number(v.recebido[k]) || 0), 0) }));
  if (!viagens.length && concluida) viagens = [{ rot: 'Viagem 1', em: nt.concluidaEm || new Date().toISOString(), por: nt.conferente || '', q: totR }];
  const sit = !concluida ? ['RECEBIMENTO PARCIAL', '#5C4300'] : faltaQ ? ['CONFERIDA COM FALTA', '#7A1556'] : sobraQ ? ['CONFERIDA COM SOBRA', '#243B8F'] : ['CONFERIDA SEM DIFERENÇA', '#10543A'];
  const datas = viagens.map(v => dataBR_(v.em, 'yyyy-MM-dd')).sort();
  const periodo = !datas.length ? '—' : datas[0] === datas[datas.length - 1] ? dataBR_(viagens[0].em) : dataBR_(datas[0] + 'T12:00:00', 'dd/MM') + ' a ' + dataBR_(datas[datas.length - 1] + 'T12:00:00');
  const emp = grupoEmpresa_(nt.fornecedor);
  const ass = (nt.assinaturas || []).filter(a => (a.conf || a.mot) && (nt.obsoleta || (a.revisao || 1) === rev)).slice(-3);
  const revTxt = nt.obsoleta ? 'DOCUMENTO OBSOLETO · revisão ' + rev + ' substituída pela revisão ' + nt.substituidaPor + ' em ' + dataBR_(nt.substituidaEm) : rev > 1 ? 'REVISÃO ' + rev + ' · substitui a conferência anterior' + (nt.motivoRevisao ? ' · motivo: ' + nt.motivoRevisao : '') : '';
  const LG = { mills: LOGO_MILLS, sh: LOGO_SH, peri: LOGO_PERI }[emp.toLowerCase()];
  const th = (al) => 'style="padding:4px 5px;font-size:7pt;text-transform:uppercase;color:#4E5752;text-align:' + al + ';border-bottom:1.5px solid #13152B"';
  const td = 'padding:3px 5px;font-size:8pt;border-bottom:1px solid #EBEEE6';
  let h = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;color:#13152B;font-size:8.5pt;margin:0}table{border-collapse:collapse;width:100%}</style></head><body>';
  if (nt.obsoleta) h += '<div style="border:3px solid #7A1556;color:#7A1556;text-align:center;font-size:26pt;font-weight:bold;letter-spacing:6px;padding:4px;margin-bottom:6px">OBSOLETO</div>';
  h += '<table><tr><td style="width:40%"><img src="data:image/jpeg;base64,' + LOGO_CCG + '" style="height:36px"></td><td style="text-align:right"><div style="font-size:6.5pt;letter-spacing:2px;color:#5E7020;font-weight:bold">LINHA 4 · EXTENSÃO GÁVEA · CONTROLE DE PEÇAS</div><div style="font-size:14pt;font-weight:bold">Relatório da coleta (conferência na fábrica)</div></td></tr></table>';
  h += '<div style="border-top:3px solid #AEC155;margin:6px 0"></div>';
  if (revTxt) h += '<div style="padding:5px 8px;background:' + (nt.obsoleta ? '#FBE4F1' : '#FBF3DD') + ';color:' + (nt.obsoleta ? '#7A1556' : '#5C4300') + ';font-size:7.5pt;font-weight:bold;margin-bottom:6px">' + esc_(revTxt) + '</div>';
  h += '<table><tr><td style="width:46px">' + (LG ? '<img src="data:image/png;base64,' + LG + '" style="width:40px;height:40px">' : '') + '</td><td><div style="font-size:16pt;font-weight:bold">NF ' + esc_(nt.numero) + ' · ' + esc_(emp) + '</div><div style="font-size:7.5pt;color:#4E5752">' + esc_(nt.fornecedor || '') + (nt.contrato ? ' · contrato ' + esc_(nt.contrato) : '') + '</div></td><td style="text-align:right"><span style="border:1.5px solid ' + sit[1] + ';color:' + sit[1] + ';font-weight:bold;font-size:8pt;padding:4px 8px">' + sit[0] + '</span>' + (nt.emissao ? '<div style="font-size:7pt;color:#4E5752;margin-top:4px">Emitida em ' + esc_(nt.emissao) + '</div>' : '') + '</td></tr></table>';
  const info = (l, v) => '<td style="border:1px solid #D7DBD0;padding:5px 7px"><div style="font-size:6pt;color:#4E5752;text-transform:uppercase">' + l + '</div><div style="font-size:8.5pt;font-weight:bold">' + v + '</div></td>';
  h += '<table style="margin-top:8px"><tr>' + info('Local da aplicação', esc_(nt.local || '—')) + info('Conferência', esc_(periodo)) + info('Conferente', esc_((viagens[viagens.length - 1] || {}).por || nt.conferente || '—')) + info('Valor da nota', nt.totalNota ? brl_(nt.totalNota) : '—') + '</tr></table>';
  const kpi = (v, l, c) => '<td style="border:1px solid #D7DBD0;padding:6px 8px"><div style="font-size:13pt;font-weight:bold;' + (c ? 'color:' + c : '') + '">' + v + '</div><div style="font-size:7pt;color:#4E5752">' + l + '</div></td>';
  h += '<table style="margin-top:6px"><tr>' + kpi(nfmt_(totQ), 'peças na nota') + kpi(nfmt_(totR), concluida ? 'peças recebidas' : 'recebidas até agora') + (concluida ? kpi(nfmt_(faltaQ), 'peças em falta', faltaQ ? '#7A1556' : '') + kpi(nfmt_(sobraQ), 'peças a mais', sobraQ ? '#243B8F' : '') + kpi(faltaV ? brl_(faltaV) : 'R$ 0', 'valor em falta', faltaV ? '#7A1556' : '') : kpi(nfmt_(totQ - totR), 'ainda por vir', '#5C4300')) + '</tr></table>';
  if (viagens.length) h += '<div style="font-weight:bold;margin:10px 0 3px">Viagens</div><table><tr><th ' + th('left') + '>Viagem</th><th ' + th('left') + '>Data e hora</th><th ' + th('left') + '>Conferente</th><th ' + th('right') + '>Peças recebidas</th></tr>' + viagens.map(v => '<tr><td style="' + td + '">' + esc_(v.rot) + '</td><td style="' + td + '">' + dataBR_(v.em, 'dd/MM/yyyy · HH:mm') + '</td><td style="' + td + '">' + esc_(v.por) + '</td><td style="' + td + ';text-align:right;font-weight:bold">' + nfmt_(v.q) + '</td></tr>').join('') + '</table>';
  h += '<div style="font-weight:bold;margin:10px 0 3px">Itens da nota</div><table><tr><th ' + th('left') + '>Código</th><th ' + th('left') + '>Descrição</th><th ' + th('right') + '>Nota</th><th ' + th('right') + '>Recebido</th><th ' + th('left') + '>Situação</th><th ' + th('right') + '>Valor unit.</th><th ' + th('right') + '>Diferença</th></tr>';
  itens.forEach((it, ix) => { const f = it.dif < 0, so = it.dif > 0; const bg = f ? '#FBF3F7' : so ? '#F3F5FA' : ix % 2 ? '#FAFBF8' : '#FFFFFF';
    const st = f ? '<b style="color:' + (concluida ? '#7A1556' : '#5C4300') + '">' + (concluida ? 'falta ' : 'a receber ') + (-it.dif) + '</b>' : so ? '<b style="color:#243B8F">sobra ' + it.dif + '</b>' : '<b style="color:#3E7A5E">ok</b>';
    h += '<tr style="background:' + bg + '"><td style="' + td + ';color:#4E5752">' + esc_(it.codigo) + '</td><td style="' + td + '">' + esc_(it.desc) + '</td><td style="' + td + ';text-align:right">' + nfmt_(it.qtd) + '</td><td style="' + td + ';text-align:right;font-weight:bold">' + nfmt_(it.rec) + '</td><td style="' + td + '">' + st + '</td><td style="' + td + ';text-align:right;color:#4E5752">' + brl_(it.vu) + '</td><td style="' + td + ';text-align:right">' + (f && concluida ? '<b style="color:#7A1556">−' + brl_(-it.dif * it.vu) + '</b>' : '—') + '</td></tr>'; });
  h += '<tr><td colspan="2" style="padding:5px;font-weight:bold;border-top:1.5px solid #13152B">Total · ' + itens.length + ' itens</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B">' + nfmt_(totQ) + '</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B">' + nfmt_(totR) + '</td><td style="border-top:1.5px solid #13152B"></td><td style="border-top:1.5px solid #13152B"></td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B;color:#7A1556">' + (faltaV && concluida ? '−' + brl_(faltaV) : '') + '</td></tr></table>';
  const oc = [];
  if (faltas.length && concluida) oc.push('<b>Falta:</b> ' + faltas.map(i => esc_(i.desc) + ' (nota ' + i.qtd + ', recebidas ' + i.rec + ')').join('; ') + '. Pendência aberta para contestação com a ' + esc_(emp) + ' no valor de ' + brl_(faltaV) + '.');
  if (sobras.length) oc.push('<b>Sobra:</b> ' + sobras.map(i => esc_(i.desc) + ' (+' + i.dif + ')').join('; ') + ' — peças fora da nota.');
  if (!concluida) oc.push('<b>Recebimento parcial:</b> ainda faltam chegar ' + nfmt_(totQ - totR) + ' peças desta nota.');
  if (oc.length) h += '<div style="margin-top:8px;padding:6px 8px;background:#F7F8F4;border-left:3px solid ' + sit[1] + ';font-size:8pt;line-height:1.5">' + oc.join('<br>') + '</div>';
  if (ass.length) h += '<table style="margin-top:22px">' + ass.map(a => '<tr>' + [[a.conf, a.confNome || '', 'Conferente · CCG'], [a.mot, a.motNome || '', 'Representante ' + emp]].map(x => '<td style="width:42%;text-align:center;vertical-align:bottom;padding:0 14px 8px"><div style="height:40px">' + (x[0] ? '<img src="' + x[0] + '" style="height:38px">' : '<span style="font-size:7pt;color:#B9BFB2">sem assinatura</span>') + '</div><div style="border-top:1px solid #13152B;font-size:8pt;font-weight:bold;padding-top:2px">' + esc_(x[1]) + '</div><div style="font-size:7pt;color:#4E5752">' + x[2] + '</div></td>').join('') + '<td style="font-size:7pt;color:#4E5752;text-align:right;vertical-align:bottom;padding-bottom:12px">' + ((a.revisao || 1) > 1 ? 'Revisão ' + a.revisao : a.encerramento ? 'Encerramento' : a.viagem ? 'Viagem ' + a.viagem : '') + '<br>' + dataBR_(a.em, 'dd/MM/yyyy · HH:mm') + '</td></tr>').join('') + '</table>';
  h += '<div style="margin-top:14px;border-top:1px solid #D7DBD0;padding-top:3px;font-size:6.5pt;color:#8A9289">Gerado pelo Controle de Peças · CCG · ' + dataBR_(new Date().toISOString(), 'dd/MM/yyyy HH:mm') + (rev > 1 || nt.obsoleta ? ' · revisão ' + rev : '') + '</div></body></html>';
  return Utilities.newBlob(h, 'text/html', 'relatorio.html').getAs('application/pdf');
}
const LOGO_MILLS = 'iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAIAAABt+uBvAAAH5klEQVR42u2cXWxcRxXH/2dm7r37ZceO7ThxsFM7iohTYpKWNIoA9aFR1JZ+qCqUAhE8FIRAAvGCFPWpEkIVQkVVEBIVQVTwVkTaPBSRlqhVS4qqJsVtTQN2o3wHJ7az3u/de+/M4eFee9eJXdmu3bjuHMl+OOP1nf3NnDP/OWdtYmZYm9uERWABWUAWkAVkAVlAFpAFZM0CsoAsIAvIArKALCALyJoFZAFZQBaQBWQBWUAWkDULyAKygCwgC2gFm1oRs5j+CA7FXxbQTCOyO2hu0yEHFZAAG1IelGsBTUeWAYnakZ8Hr/2BMu1cysrP7U0+9gzAKyfQbi4ggMC5US5PQjpcyXNudDoVWUBT5niIIku5cLwZ+4tNHZaQs/kZJJc1hd18QOyXuVqEm+RqEbVSQ+YWoNlUyFz+VQiIAEBu2GZu2UnptVzJye7tAMAaJPWZE+G7R8lLc1AV7ZucPd+Icpa5dCo4+QI5HgAOfeeOh0Xnlmho9QESANx7fuLe/SOAwBzHEQMEPfhi7S9P0JpOruRlz4Cz+5GIqR4+XnvucWrqAJiL10Rrl+jcAuZlSlwrRAfJWVKzm6LmDsq0Q7mUamnMWdTcQel2gCEElLeSQoxNPTyi1Mg85SSAp3IENSRUnnP0uvPsutOdDXQIE0KHMHqG7I78YOgQN34Qvj6lG+Qo0UKPSLWIoIjNaAgJonj9G80YgCEkSMwyn1nyBS3Z6R7PSt6UHcRmdASBDxNSew+l1wIwox/oMyfM6DAXJ0h5Yv0WtfM+atkQi+SLQ+bsv8zV01zKQnmya6va+RVq7ly2nMoQkovj+r/H9ZXTPHkZOiSpqHWjaOsWG7eJrq0Lfe78ADGDiGvlysGvcWmS/Yraca9z+4PB63/SF97lchbM8YONpqMHE/t/hUrBf/WQuXyKK/np0cBo+tvTye8ekpvvWGpGHE3SP3owePWQmRyFCUEEroMjNyV6BpLff5YybdEPL/UOmoptSjbr91/Rg39lISmRoeZ1cVhF77k8Wf3dYwh8JqJEUzzKBsZASM6NVv/449SBlyjZvJQ3CgYE1Z573D96kDJtlG6BkA3piUCE0DcXh7g8SZm2+d9mFnqKTelXqZBsoqDGpWxMzUnATcKE8W1TOhTUuDQRz9JJkJtEEFKq1Vz5IBw65ux6KNI7M+e6WGZS6ZE3/FcOUcsGMCOocrUE5cS/zWgYDTZwkx/XMc8MoUT3dtG1VbR1g0Q49Hd9bpCcRHyCCCVu2Sm7+ql1Axjh0Mv6/HvkeNG5xheHsOuhmdDxEfN0+N5LMBokEFaRbvXuPyA7t0TCimtFLk2aaxc4P0ap1gU9aFGAhORi1tnz9cS3fz0dyc7uR0pP7oVfhnS5dM358ncS33qqrl32PFp6ci9qRUgPQujspek6mf/yb4K3j1CqBZWC7PuC99WfYVF/wmayl6JzkysF94v73b0/XJIKlFp80CsvCuxI2lDzOtG8Tl8ZIeUBTNIBUB/NrBWpNaacix/YoGv0+UH9n9epqYNL2Y+UtjlK1QZuwpwbNGNnRFtP/YoLQAcAQaqPJcTi9WcICSEAgvZnbNtIyMVSiBBUYcys60ZuipLNlMzAaHipxavxTBsMgw25KX3+ncpT99O6zaKrX35mm+joFRv7KdM+hwpbxqvGh6SPmZr1Q7Y0GxgNY+I8umhF13t7cOy3EArah5PgSsGcflMP/yNgwE2ITLvo2+Xd91Ox4bPzP+OxeroaRssd96r+O/nahTh+HY9SLdTUQU3t5CS5nAvfOlx5+mG+ehpE81+J1QKImbx04nu/d+/6ASXXcK3EhXEujHE5h6ASHSzUst5kL9aOPTOVsD5Bt/klqAcQmKmp3fvmL93igShJ60vvm9ERnrhgCuPkeAhr5GXM2ZMwIYSap+ZaLYCiXaHDKFvLW++SgANAh1zO1g4/Eb75ZyRbQMTlHIcBuepTFmJRpUE6iORFg8Kmpg7Zu4sDP6518EICbPXsICGD157l4oTo6qfmDko2QSiAoAOTGw3fOkxeOi5OKZeE/PQBAoLBF8OTRyjTRsqBcGJ1xpr9ChHBTQDE2pfr+qC85bnNAxASQsXfb5RbQkIqCAU5x6iYbZRE/VWNazvtF9f5KfZHd+ap90nSoXQrZdZCB3FpgRlE5GVABDDCAGHN+dL+SBcAcukBcSnLxSzcBBfGUSvfMDrJ+TEkqpwf48YGzvRrC2Pw0lwcR61Q99fKnL8KgItZVPJ1v1/m3BhIcDnH5cm6P6hxfixKI1ycQOjH/DfeSsPHOT8OE0RYiQhgZoYJYTSl1yYe/YUauBvM8y850vz+uQkDBKODE8+jUoBy2K/ITTtk3666cmcOTr6A/BiUy35F9t4mN+8Gm6nUaIITz6MwAeVyUJEbt8mtd8IYCKFH/qnPDZKXZL8qOvrUwL7Yf/4dPfwGuUkOfdGyXt32QNz2GR0J/32MHA8MDn01sE909MYlvYkL+uzb5vIp/b9hzl9BGIAITkK0dcuez8vt+0T7pgXJ6PkD+kRKx3h5GsN54WXMBQIyut6BuLHDubjReiuZAZqtxXydn8ENTY7G1vN0P+O63kmcj8QimtSrdwfNEDuLr8MprFpbmnK3/YyiBWQBWUAWkAVkAVlAFpA1C8gCsoAsIAvIArKALCBrFpAFZAFZQBaQBWQBWUDWLCALyAKygCyglWz/BxjOlgTOCLO0AAAAAElFTkSuQmCC';
const LOGO_SH = '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';
const LOGO_PERI = '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';

/* ---------------- Recebimento no canteiro (com as fotos das avarias) ---------------- */
function lerFoto_(tk, id) {
  try { const r = UrlFetchApp.fetch(FS + '/fotos/' + encodeURIComponent(id), { headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null; const j = JSON.parse(r.getContentText()); return j.fields && j.fields.dados ? j.fields.dados.stringValue : null; } catch (e) { return null; }
}
function cabecalho_(titulo, sub) {
  return '<table><tr><td style="width:40%"><img src="data:image/jpeg;base64,' + LOGO_CCG + '" style="height:36px"></td><td style="text-align:right"><div style="font-size:6.5pt;letter-spacing:2px;color:#5E7020;font-weight:bold">LINHA 4 · EXTENSÃO GÁVEA · CONTROLE DE PEÇAS</div><div style="font-size:14pt;font-weight:bold">' + esc_(titulo) + '</div>' + (sub ? '<div style="font-size:7.5pt;color:#4E5752">' + esc_(sub) + '</div>' : '') + '</td></tr></table><div style="border-top:3px solid #AEC155;margin:6px 0 8px"></div>';
}
const CSS_PDF_ = '<style>@page{size:A4;margin:12mm}body{font-family:Arial,sans-serif;color:#13152B;font-size:8.5pt;margin:0}table{border-collapse:collapse;width:100%}</style>';
function fotosHtml_(lista, fotos) { // [{rot, ids}]
  let h = ''; lista.forEach(x => (x.ids || []).forEach((id, k) => { if (fotos[id]) h += '<div style="display:inline-block;width:31%;margin:0 1.5% 10px 0;vertical-align:top"><img src="' + fotos[id] + '" style="width:100%;border:1px solid #D7DBD0"><div style="font-size:7pt;color:#4E5752">' + esc_(x.rot) + ' · ' + (k + 1) + '/' + x.ids.length + '</div></div>'; }));
  return h;
}
function enviarRecebimentos(detalhe) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) { if (detalhe) throw new Error('outra sincronização está rodando, tente em 1 minuto'); return 0; }
  let n = 0, lidos = 0; const erros = [];
  try {
    const tk = tokenRobo_(); const raiz = DriveApp.getFolderById(PASTA_TERMOS_ID);
    const lista = pendentesCampo_(tk, 'notas', 'recebDrive.ok'); lidos = lista.length;
    lista.forEach(nota => {
      const marcar = (d) => UrlFetchApp.fetch(FS + '/notas/' + encodeURIComponent(nota._id) + '?updateMask.fieldPaths=recebDrive', { method: 'patch', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true, payload: JSON.stringify({ fields: { recebDrive: fsEsc_(d) } }) });
      try {
        if (!nota.receb || nota.receb.status !== 'concluido') { marcar({ ok: true, obs: 'sem recebimento' }); return; }
        const emp = grupoEmpresa_(nota.fornecedor);
        const pasta = pasta_(pasta_(pasta_(pasta_(raiz, emp), nota.local || 'Sem local'), 'NF ' + (nota.numero || 's-n')), '2 · Recebimento');
        const nome = semAcento_('Recebimento NF ' + (nota.numero || '') + ' - ' + emp + '.pdf');
        const ant = pasta.getFilesByName(nome); while (ant.hasNext()) ant.next().setTrashed(true);
        pasta.createFile(pdfRecebimento_(nota, tk).setName(nome)).setDescription('Recebimento no canteiro · gerado pelo Controle de Peças');
        marcar({ ok: true, em: new Date().toISOString() }); n++;
      } catch (err) { erros.push('NF ' + (nota.numero || nota._id) + ': ' + err.message); marcar({ ok: false, erro: String(err.message || err).slice(0, 120), em: new Date().toISOString() }); }
    });
  } finally { lock.releaseLock(); }
  return detalhe ? { enviados: n, lidos: lidos, erros: erros } : n;
}
function pendentesCampo_(tk, colecao, campo) {
  const r = UrlFetchApp.fetch(FS + ':runQuery', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: colecao }], where: { fieldFilter: { field: { fieldPath: campo }, op: 'EQUAL', value: { booleanValue: false } } }, limit: 15 } }) });
  if (r.getResponseCode() !== 200) throw new Error('não consegui ler ' + colecao + ': ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText() || '[]').filter(x => x.document).map(x => { const o = {}; const f = x.document.fields || {}; Object.keys(f).forEach(k => o[k] = fsLer_(f[k])); o._id = x.document.name.split('/').pop(); return o; });
}
function pdfRecebimento_(nt, tk) {
  const it = nt.receb.itens || {}; const emp = grupoEmpresa_(nt.fornecedor);
  let col = 0, q = 0, av = 0, avV = 0, fa = 0, qn = 0; const lista = [];
  (nt.itens || []).forEach(i => { const x = it[i.key] || {}; qn += Number(i.qtd) || 0; col += x.col || 0; q += x.q || 0; av += x.avq || 0; avV += (x.avq || 0) * (i.vu || 0); fa += Math.max(0, (x.col || 0) - (x.q || 0)); if ((x.fotos || []).length) lista.push({ rot: i.desc, ids: x.fotos }); });
  const fotos = {}; lista.forEach(x => x.ids.forEach(id => { fotos[id] = lerFoto_(tk, id); }));
  const sit = av ? ['RECEBIDA COM AVARIA', '#7A1556'] : fa ? ['RECEBIDA COM FALTA', '#7A1556'] : ['RECEBIDA SEM DIFERENÇA', '#10543A'];
  const LG = { mills: LOGO_MILLS, sh: LOGO_SH, peri: LOGO_PERI }[emp.toLowerCase()];
  let h = '<html><head><meta charset="utf-8">' + CSS_PDF_ + '</head><body>' + cabecalho_('Relatório de recebimento no canteiro', '');
  h += '<table><tr><td style="width:46px">' + (LG ? '<img src="data:image/png;base64,' + LG + '" style="width:40px;height:40px">' : '') + '</td><td><div style="font-size:16pt;font-weight:bold">NF ' + esc_(nt.numero) + ' · ' + esc_(emp) + '</div><div style="font-size:7.5pt;color:#4E5752">' + esc_(nt.local || '') + ' · recebida ' + dataBR_(nt.receb.em, 'dd/MM/yyyy HH:mm') + ' · conferente ' + esc_(nt.receb.por || '') + '</div></td><td style="text-align:right"><span style="border:1.5px solid ' + sit[1] + ';color:' + sit[1] + ';font-weight:bold;font-size:8pt;padding:4px 8px">' + sit[0] + '</span></td></tr></table>';
  const kpi = (v, l, c) => '<td style="border:1px solid #D7DBD0;padding:6px 8px"><div style="font-size:13pt;font-weight:bold;' + (c ? 'color:' + c : '') + '">' + v + '</div><div style="font-size:7pt;color:#4E5752">' + l + '</div></td>';
  h += '<table style="margin-top:8px"><tr>' + kpi(nfmt_(qn), 'na nota') + kpi(nfmt_(col), 'na coleta') + kpi(nfmt_(q), 'no canteiro') + kpi(nfmt_(av), 'avariadas', av ? '#7A1556' : '') + kpi(brl_(avV), 'valor avariado', avV ? '#7A1556' : '') + '</tr></table>';
  const th = (t, a) => '<th style="padding:4px 5px;font-size:7pt;text-transform:uppercase;color:#4E5752;text-align:' + a + ';border-bottom:1.5px solid #13152B">' + t + '</th>';
  h += '<table style="margin-top:10px"><tr>' + th('Descrição', 'left') + th('Nota', 'right') + th('Coleta', 'right') + th('Canteiro', 'right') + th('Avariadas', 'right') + th('Observação', 'left') + '</tr>';
  (nt.itens || []).forEach(i => { const x = it[i.key] || {}; const ruim = (x.avq || 0) > 0 || (x.q || 0) !== (x.col || 0); const td = 'padding:3px 5px;font-size:8pt;border-bottom:1px solid #EBEEE6';
    const obs = [x.obs, (x.q || 0) < (x.col || 0) ? ((x.col || 0) - (x.q || 0)) + ' a menos que a coleta' : '', (x.q || 0) > (x.col || 0) ? '+' + ((x.q || 0) - (x.col || 0)) + ' a mais' : ''].filter(Boolean).join(' · ');
    h += '<tr style="background:' + (ruim ? '#FBF3F7' : '#fff') + '"><td style="' + td + '">' + esc_(i.desc) + '</td><td style="' + td + ';text-align:right">' + i.qtd + '</td><td style="' + td + ';text-align:right">' + (x.col || 0) + '</td><td style="' + td + ';text-align:right;font-weight:bold">' + (x.q || 0) + '</td><td style="' + td + ';text-align:right;color:#7A1556;font-weight:bold">' + (x.avq || '—') + '</td><td style="' + td + ';color:#4E5752">' + esc_(obs) + '</td></tr>'; });
  h += '</table>';
  if (lista.length) h += '<div style="font-weight:bold;font-size:10pt;margin:14px 0 6px">Fotos das avarias</div>' + fotosHtml_(lista, fotos);
  const a = nt.receb.assinaturas;
  if (a) h += '<table style="margin-top:20px"><tr>' + [[a.conf, a.confNome, 'Conferente · CCG'], [a.rep, a.repNome, 'Representante ' + emp]].map(x => '<td style="width:50%;text-align:center;padding:0 18px"><div style="height:42px">' + (x[0] ? '<img src="' + x[0] + '" style="height:40px">' : '') + '</div><div style="border-top:1px solid #13152B;font-size:8pt;font-weight:bold;padding-top:2px">' + esc_(x[1] || '') + '</div><div style="font-size:7pt;color:#4E5752">' + x[2] + '</div></td>').join('') + '</tr></table>';
  h += '<div style="margin-top:12px;border-top:1px solid #D7DBD0;padding-top:3px;font-size:6.5pt;color:#8A9289">Gerado pelo Controle de Peças · CCG · ' + dataBR_(new Date().toISOString(), 'dd/MM/yyyy HH:mm') + '</div></body></html>';
  return Utilities.newBlob(h, 'text/html', 'recebimento.html').getAs('application/pdf');
}

/* ---------------- Conferência mensal (dia 10): um PDF por frente e um resumo do mês ---------------- */
function enviarMensais(detalhe) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) { if (detalhe) throw new Error('outra sincronização está rodando, tente em 1 minuto'); return 0; }
  let n = 0, lidos = 0; const erros = [];
  try {
    const tk = tokenRobo_(); const raiz = DriveApp.getFolderById(PASTA_TERMOS_ID);
    const lista = pendentesCampo_(tk, 'conferencias', 'drive.ok'); lidos = lista.length; const meses = {};
    lista.forEach(d => {
      try {
        const pasta = pasta_(pasta_(raiz, 'Conferências mensais'), d.mes + ' · ' + dataBR_(d.data + 'T12:00:00', 'dd-MM-yyyy'));
        const nome = semAcento_('Conferencia mensal ' + dataBR_(d.data + 'T12:00:00', 'dd-MM-yyyy') + ' - ' + d.local + '.pdf').replace(/[\\/:*?"<>|]/g, '-');
        const ant = pasta.getFilesByName(nome); while (ant.hasNext()) ant.next().setTrashed(true);
        pasta.createFile(pdfMensal_([d], tk).setName(nome)).setDescription('Conferência mensal · gerado pelo Controle de Peças');
        marcarDrive_(tk, d._id, { ok: true, em: new Date().toISOString() }, 'conferencias'); meses[d.mes] = pasta; n++;
      } catch (err) { erros.push(d.local + ': ' + err.message); marcarDrive_(tk, d._id, { ok: false, erro: String(err.message || err).slice(0, 120), em: new Date().toISOString() }, 'conferencias'); }
    });
    Object.keys(meses).forEach(mes => { // resumo do mês com todas as frentes já conferidas
      try {
        const r = UrlFetchApp.fetch(FS + ':runQuery', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
          payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: 'conferencias' }], where: { fieldFilter: { field: { fieldPath: 'mes' }, op: 'EQUAL', value: { stringValue: mes } } } } }) });
        const docs = JSON.parse(r.getContentText() || '[]').filter(x => x.document).map(x => { const o = {}; const f = x.document.fields || {}; Object.keys(f).forEach(k => o[k] = fsLer_(f[k])); return o; }).sort((a, b) => a.local.localeCompare(b.local));
        if (!docs.length) return;
        const nome = semAcento_('Conferencia mensal ' + dataBR_(docs[0].data + 'T12:00:00', 'dd-MM-yyyy') + ' - Resumo.pdf');
        const ant = meses[mes].getFilesByName(nome); while (ant.hasNext()) ant.next().setTrashed(true);
        meses[mes].createFile(pdfMensal_(docs, tk, true).setName(nome));
      } catch (err) { erros.push('resumo ' + mes + ': ' + err.message); }
    });
  } finally { lock.releaseLock(); }
  return detalhe ? { enviados: n, lidos: lidos, erros: erros } : n;
}
function pdfMensal_(docs, tk, resumo) {
  const tit = 'Conferência mensal · ' + dataBR_(docs[0].data + 'T12:00:00');
  const th = (t, a) => '<th style="padding:4px 5px;font-size:7pt;text-transform:uppercase;color:#4E5752;text-align:' + a + ';border-bottom:1.5px solid #13152B">' + t + '</th>';
  const td = 'padding:3px 5px;font-size:8pt;border-bottom:1px solid #EBEEE6';
  const res = (d) => { let esp = 0, cont = 0, perd = 0, dan = 0, val = 0; (d.itens || []).forEach(i => { esp += i.esperado || 0; cont += i.contado || 0; perd += i.perdidas || 0; dan += i.danificadas || 0; val += ((i.perdidas || 0) + (i.danificadas || 0)) * (i.vu || 0); }); return { esp, cont, perd, dan, val }; };
  let h = '<html><head><meta charset="utf-8">' + CSS_PDF_ + '</head><body>';
  if (resumo) {
    h += cabecalho_(tit, 'Resumo de todas as frentes conferidas');
    h += '<table><tr>' + th('Frente', 'left') + th('Conferido por', 'left') + th('Esperadas', 'right') + th('Contadas', 'right') + th('Perdidas', 'right') + th('Danificadas', 'right') + th('Valor', 'right') + '</tr>';
    const t = { esp: 0, cont: 0, perd: 0, dan: 0, val: 0 };
    docs.forEach(d => { const r = res(d); Object.keys(t).forEach(k => t[k] += r[k]); h += '<tr><td style="' + td + ';font-weight:bold">' + esc_(d.local) + '</td><td style="' + td + '">' + esc_(d.por) + ' · ' + dataBR_(d.em, 'dd/MM HH:mm') + '</td><td style="' + td + ';text-align:right">' + r.esp + '</td><td style="' + td + ';text-align:right">' + r.cont + '</td><td style="' + td + ';text-align:right;color:#7A1556">' + (r.perd || '—') + '</td><td style="' + td + ';text-align:right;color:#7A1556">' + (r.dan || '—') + '</td><td style="' + td + ';text-align:right">' + (r.val ? brl_(r.val) : '—') + '</td></tr>'; });
    h += '<tr><td style="padding:5px;font-weight:bold;border-top:1.5px solid #13152B" colspan="2">Total</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B">' + t.esp + '</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B">' + t.cont + '</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B;color:#7A1556">' + t.perd + '</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B;color:#7A1556">' + t.dan + '</td><td style="padding:5px;text-align:right;font-weight:bold;border-top:1.5px solid #13152B">' + brl_(t.val) + '</td></tr></table>';
  } else {
    const d = docs[0]; const r = res(d); const lista = (d.itens || []).filter(i => (i.fotos || []).length).map(i => ({ rot: i.desc, ids: i.fotos }));
    const fotos = {}; lista.forEach(x => x.ids.forEach(id => { fotos[id] = lerFoto_(tk, id); }));
    h += cabecalho_(tit, d.local + ' · conferido por ' + d.por + ' em ' + dataBR_(d.em, 'dd/MM/yyyy HH:mm'));
    h += '<div style="font-size:10pt;margin-bottom:6px"><b>' + nfmt_(r.cont) + '</b> de <b>' + nfmt_(r.esp) + '</b> peças · <b style="color:#7A1556">' + r.perd + ' perdidas</b> · <b style="color:#7A1556">' + r.dan + ' danificadas</b>' + (r.val ? ' · ' + brl_(r.val) : '') + '</div>';
    h += '<table><tr>' + th('Peça', 'left') + th('Esperadas', 'right') + th('Contadas', 'right') + th('Perdidas', 'right') + th('Danificadas', 'right') + th('Observação', 'left') + '</tr>';
    (d.itens || []).forEach(i => { h += '<tr style="background:' + ((i.perdidas || i.danificadas) ? '#FBF3F7' : '#fff') + '"><td style="' + td + '">' + esc_(i.desc) + '</td><td style="' + td + ';text-align:right">' + i.esperado + '</td><td style="' + td + ';text-align:right;font-weight:bold">' + i.contado + '</td><td style="' + td + ';text-align:right;color:#7A1556;font-weight:bold">' + (i.perdidas || '—') + '</td><td style="' + td + ';text-align:right;color:#7A1556;font-weight:bold">' + (i.danificadas || '—') + '</td><td style="' + td + ';color:#4E5752">' + esc_(i.obs || '') + '</td></tr>'; });
    h += '</table>';
    if (lista.length) h += '<div style="font-weight:bold;font-size:10pt;margin:14px 0 6px">Fotos das peças danificadas</div>' + fotosHtml_(lista, fotos);
  }
  h += '<div style="margin-top:12px;border-top:1px solid #D7DBD0;padding-top:3px;font-size:6.5pt;color:#8A9289">Gerado pelo Controle de Peças · CCG · ' + dataBR_(new Date().toISOString(), 'dd/MM/yyyy HH:mm') + '</div></body></html>';
  return Utilities.newBlob(h, 'text/html', 'mensal.html').getAs('application/pdf');
}
