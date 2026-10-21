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
    .addItem('Sincronizar usuários agora', 'sincronizarUsuarios')
    .addItem('Enviar termos ao Drive agora', 'enviarTermosAgora')
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
  const n = enviarTermos();
  SpreadsheetApp.getActiveSpreadsheet().toast(n ? n + ' termo(s) enviado(s) ao Drive.' : 'Nenhum termo pendente.', 'Controle de Peças');
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
function pendentesDrive_(tk) {
  const r = UrlFetchApp.fetch(FS + ':runQuery', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: 'remessas' }], where: { fieldFilter: { field: { fieldPath: 'drivePendente' }, op: 'EQUAL', value: { booleanValue: true } } }, limit: 15 } }) });
  if (r.getResponseCode() !== 200) throw new Error('não consegui ler os pedidos: ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText() || '[]').filter(x => x.document).map(x => { const o = {}; const f = x.document.fields || {}; Object.keys(f).forEach(k => o[k] = fsLer_(f[k])); o._id = x.document.name.split('/').pop(); return o; });
}
function marcarDrive_(tk, id, drive) {
  const url = FS + '/remessas/' + encodeURIComponent(id) + '?updateMask.fieldPaths=drive&updateMask.fieldPaths=drivePendente';
  UrlFetchApp.fetch(url, { method: 'patch', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tk }, muteHttpExceptions: true,
    payload: JSON.stringify({ fields: { drive: fsEsc_(drive), drivePendente: { booleanValue: !drive.ok } } }) });
}
function enviarTermos() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return 0;
  let n = 0;
  try {
    const tk = tokenRobo_(); const raiz = DriveApp.getFolderById(PASTA_TERMOS_ID);
    pendentesDrive_(tk).forEach(r => {
      if (!r.termo) { marcarDrive_(tk, r._id, { ok: true, versao: 0, em: new Date().toISOString(), obs: 'sem termo' }); return; }
      const versao = (r.devolucoes || []).length;
      try {
        const arquivos = ['colab', 'resp'].map(via => ({ nome: nomeTermo_(r, via), blob: pdfTermo_(r, via) }));
        const pastas = (r.termo.pastas && r.termo.pastas.length) ? r.termo.pastas : [{ empresa: 'Sem nota', local: r.destino, nf: 'Sem nota' }];
        pastas.forEach(p => {
          const pasta = pasta_(pasta_(pasta_(raiz, p.empresa), p.local), p.nf);
          arquivos.forEach(a => { const ant = pasta.getFilesByName(a.nome); while (ant.hasNext()) ant.next().setTrashed(true); pasta.createFile(a.blob.copyBlob().setName(a.nome)).setDescription('Pedido ' + (r.numero || '') + ' · gerado pelo Controle de Peças'); });
        });
        marcarDrive_(tk, r._id, { ok: true, versao: versao, em: new Date().toISOString() }); n++;
      } catch (err) { marcarDrive_(tk, r._id, { ok: false, erro: String(err.message || err).slice(0, 120), em: new Date().toISOString() }); }
    });
  } finally { lock.releaseLock(); }
  return n;
}
function dataBR_(iso, fmt) { return Utilities.formatDate(new Date(iso), 'America/Sao_Paulo', fmt || 'dd/MM/yyyy'); }
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function semAcento_(t) { return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function nomeTermo_(r, via) {
  const c = r.termo.colaborador || {};
  return semAcento_(dataBR_(r.termo.em, 'yyyy-MM-dd') + ' - ' + r.numero + ' - ' + String(c.nome || '').toUpperCase() + (c.matricula ? ' (' + c.matricula + ')' : '') + ' - Termo ' + (via === 'colab' ? 'FOR-SUP-001 via colaborador' : 'FOR-SUP-002 via responsavel') + '.pdf').replace(/[\\/:*?"<>|]/g, '-').replace(/[^\x20-\x7E]/g, '');
}
function linhasTermo_(r) {
  const out = [];
  (r.termo.itens || []).forEach(i => {
    const desc = (String(i.desc || i.key) + (i.nfTxt ? ' · ' + i.nfTxt : '')).toUpperCase(); let resta = Number(i.qtd) || 0;
    (r.devolucoes || []).forEach(d => { const q = Math.min(resta, Number((d.itens || {})[i.key]) || 0); if (q > 0) { out.push({ qtd: q, desc: desc, dev: d }); resta -= q; } });
    if (resta > 0) out.push({ qtd: resta, desc: desc, dev: null });
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
    h += '<table style="margin-top:6px"><tr><th ' + th + ' width="8%">QUANTIDADE</th><th ' + th + '>DESCRIÇÃO DO ITEM</th><th ' + th + ' width="11%">Nº DO PATRIMÔNIO</th><th ' + th + ' width="11%">DATA DA RETIRADA</th><th ' + th + ' width="12%">VISTO DO INTEGRANTE</th><th ' + th + ' width="11%">DATA DA DEVOLUÇÃO</th><th ' + th + ' width="12%">VISTO DO ALMOXARIFADO</th></tr>';
    linhas.forEach(l => { h += l ? '<tr>' + td(l.qtd, 'text-align:center') + td(esc_(l.desc)) + td('') + td(dRet, 'text-align:center') + td(img(t.rubrica, 16), 'text-align:center;padding:0') + td(dd(l.dev), 'text-align:center') + td(l.dev ? img(l.dev.visto, 16) : '', 'text-align:center;padding:0') + '</tr>' : '<tr>' + td('') + td('') + td('') + td('') + td('') + td('') + td('') + '</tr>'; });
  } else {
    h += '<table style="margin-top:6px"><tr><th ' + th + ' rowspan="2" width="7%">QUANT.</th><th ' + th + ' rowspan="2">DESCRIÇÃO DO ITEM</th><th ' + th + ' rowspan="2" width="8%">Nº DO PATRIM.</th><th ' + th + ' colspan="4">RETIRADA</th><th ' + th + ' colspan="2">DEVOLUÇÃO</th></tr><tr><th ' + th + ' width="8.5%">DATA</th><th ' + th + ' width="15%">NOME / MATRÍCULA</th><th ' + th + ' width="9.5%">VISTO DO INTEGRANTE</th><th ' + th + ' width="6.5%">OBRA</th><th ' + th + ' width="8.5%">DATA</th><th ' + th + ' width="9.5%">VISTO DO ALMOX.</th></tr>';
    const nm = esc_(c.nome) + (c.matricula ? ' / ' + esc_(c.matricula) : '');
    linhas.forEach(l => { h += l ? '<tr>' + td(l.qtd, 'text-align:center') + td(esc_(l.desc), 'font-size:7pt') + td('') + td(dRet, 'text-align:center;font-size:7pt') + td(nm, 'font-size:6.5pt') + td(img(t.rubrica, 16), 'text-align:center;padding:0') + td('GÁVEA', 'text-align:center;font-size:6.5pt') + td(dd(l.dev), 'text-align:center;font-size:7pt') + td(l.dev ? img(l.dev.visto, 16) : '', 'text-align:center;padding:0') + '</tr>' : '<tr>' + td('') + td('') + td('') + td('') + td('') + td('') + td('') + td('') + td('') + '</tr>'; });
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
const LOGO_CCG = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBAUEBAYFBQUGBgYHCQ4JCQgICRINDQoOFRIWFhUSFBQXGiEcFxgfGRQUHScdHyIjJSUlFhwpLCgkKyEkJST/2wBDAQYGBgkICREJCREkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCT/wAARCABUAO8DASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6poorF8XeMNF8EaNNq+u3qWtrGOM/ekbsqjqSfSgDazWDr/jzwv4WUnWte06wI/hlmAb8utfMniP43/EX4w6tJoXw/wBPu7GyJ2n7P/rmX+9JJ0QewrFuPgz4Y8NN9q+JfxEtba+b5nsrJvtE+fQsc81oqfdiufQV1+0x8MLWQxnxCZCOMxW0jD8wK0NJ+Pvw21h1SDxTZxuxwFuN0RJ/4EK+alP7P1sNkdt4x1AD/lrtK59+lNOgfAXX28m08SeIPD1w33Wv4d0Y/Sq5EK59o2V/aajAJ7O5huYW6SROHU/iKnr4xT4cfET4bRf8JF8PfEaa7pafMZNKl3jb/twnIPvivVvhD+03YeLLiLQvFkcWlaux8uOf7sE7ehz9xvY8VDh1Q7nvFFAIIyKKgYUUUZoAKKKKACiimrLG0jRh1LqAWUHkA9MigB1FFIWAxkgZ6e9AC0UUZoAKKAQaKACiijNABRRRQAUUUUAFFFFAGb4j8Q6f4V0S81nVJxBZ2cZkkc+g7D1JPAr5Bd9f/aS8ZXerareHSPCelAvLK5xHZw9cDsZGHU1137UHi2/8WeLNJ+GehEyMZI2uEQ/fmf7in2UfMag1Dw7FrGp2Xwb8O3X2Tw7oUYvPE2pIcedIOX3H9P8A9VbRVlclhpd/qnjCyuPDfwtij8IeBrDK33iCf5JLnHVi55JP5/SuWuPFfwk+HcrRaFoMnjfWFOH1PU2/cs/+yv8AEPwp+t6vq3xm1+1+Hnw9tfsHhSw+SONMrGyKcGeYjqD2HevoH4c/s/eD/ANtFI9lHquqYHmXl2gbnvsU8KKG0twPBIfjn8TLlN2h+DNNtbQfdjg0hmUD6nFQXHx0mnb7L4++G2jX0T8O/wBlNtL/AMBJHX8a+zkjSNQqIqgcAAYFUNY8PaT4gtWtNV021voGBBSeIMMfj0qeddh2PljwxoukavO2ufBHxPdaJrcQ8yXw7fyYEoHULnhh+dQ6voOlfG6O9ji0xPDPxL05SbrTyPLj1Hb1Kjs3v/Tmtv4ufs3T+F93i74dSXMLWZ897GNz5kWOd8R68f3axrDWZfjT4eGt6e62PxM8LoLhJofkOpQr3x3Yen+NWn1Qjs/2cPjPeXtyfAHi2SRNUtd0dnNPw77esL5/iHb1FfRNfFvxJZfFvh3SPjF4dX7FqkM6W+sxRDBgu0+7LjtnGD9RX1P8L/G0PxB8EaZr8eBLPHtnQfwSrw4/P+dRNdUNGP8AEe91DVvEvh3wTYX9xpsWq+dcXt1bttl8iIDKI38JYkDPpSXtlpfwe02fUtPbWdTuL+SKzttPnvWlEs7HChd/3Se59BWj4/8ACeq6te6P4h8OzW0et6NI7RR3ORFcxOMPExHIz1B7EVh61o3jrxxpqtqGlaVol7pd1Df6aBdGcSTxk5V8AYUqSPXmpQyeX4p6poUes2nibQI7TVdO05tUhitrjzIruEHB2sQMEHAOR3pLD4oaxY3sFt4o8PxWYvtPm1Cyksrjz/MWNN7RkYBD7SCO1ZuqeBvF3jX+3NX1u30/Tb6fR5NJsLKGcyKu9gzyO+O5AwB0FdBqnhHV5Nb8I6naLayf2JZXEMqSOQHkeFVUD2yOaNBFPRvilfyeErzxjrWm2NvoqWoubc2d358rEn5Y2GMBzkD2NUfDHiS88O+I5f8AhIfDR0278SJJeQXIvPPLtGm77O/A2EL0A461kR/CbW/EU2vNdaZp/hS21Kx8hrWyuDNHNdCQOk5XAVcEY45INbsPhjxn4q1nTLzxPa6Zp8OhwzeQlrMZDd3DxmPzDx8iAE8deaegDbD4t603hSPxRqHh2BLPUfKj0m3t7rfPczSPtWNgQAvrmptV1bU2vvC//CYaBb295JrCx2n2K/ZkjJjY734G7HIx0pn/AArXV2+Ffh7QI7i2h1zQ3guoHbLQtNExIVu+0gkZqxc6P4w8W3WgXetaVp+ltpWqrcvHFdeb5kXlsrHOBg5PA9KWgyq3xf1P7K/iRPDm7wel0bZr77QPP2h9hnEeP9WG9845q38Rr291vxH4a8GWWoXFha6x51zeXNs22VoIlB2I38O4kZI7Vgn4e+M18PS/DyNdNPhyS4Y/2qZT5wtTJ5hi8vH3/wCHdnGK67xz4M1HUpdF1nw3c29trWhOxthcgmKeJl2vE+OQCAOR0Io0EUNQstI+D+lz6npUOp3txfSQ2VvYSXjyLLOzYXG8nbnqT6Com+J+s6F/bFj4o0O2ttUsdMk1W3W0uDJDdRLwyhiAQwOM/Wma1oPjrxvpZTU7LR9Fu7C4hvdOMc7T7p42ziTgfIwyOOearX/gTxZ4wk1rVteXTbG9m0ebSbC0t5TIiGTlpHfHc4AA6CjTqBb0j4meILe+sbfxPoNpbR6pYy31jJYXBl3eWgcxsCBhip4xxWReeP8AxL4k+FviHxCsOm2Fo2mPPZTWN4ZJ4n7LIMfKwHXHQ10994N1V77wdeWz2rNoNtLHKsjEB3MIQY9sjn2rlB8K9f129165ubLSfDMep6VNYy2+nytIl1OxBWZ1wANuPqcmjQDYm8c+MNE8O6MLvQ9Mk1TVLqCys1F4xjbdGWLyNtyOnQVu+HfGGqX3iO88O6vp9tbXtlp0N7M1vKXQtIzjaMjoNvX3rFi8M+Mddg8NLrdtpNm+h6lBOfs87P50SRlWPI4JJGBVvxF4d8T6b40l8UeGINPvhfWK2N3aXcpiI2MSjqwB/vEEUaActH4/8UeJdf8ACGpaDZwFtQ068ll0+e6McPyS7QxIByRjjjvW/qPxI8TSXerf2D4btb+00HCalJJdFGklCB3jgGPmKg9TjNZWlfDvxb4UsPCV7pY0u+1PR7a5tbq2llZI5FmfflGxn5T69atz+E/HWjz67DoI0d7bxE32iZ55GB0+d4wshUY+deMjoc09ALrfErWNe1vTtO8I6VY3cV1pcWrPPe3BjxE77doAB+YYNejDOBnrXB+Efh5N4U8TWt1DPHJp1roMOlJknzGkSQuWI6YOa7ypdugz4z+FurDUvG3jj4pamPMGj2093Fv5AmkJWMfgBUWs6jceCPgbAWkb+3vHty95eS5+f7MDnb685H5ms3w0DZfs5+KpoziTUdZtrJiOpQAHH6mui+KunpqPxc8BeFGH+i2llYW2zthiGY1v1JPcv2ffhvD4A8CWrywqNV1JFurt8cjIyqfRQfzzXe654j0nw1ai61a/hs4icBpG+8fQDvWiqhFCqMADAFeOfFHwHrnxB8Z+RYXEEdvYWcZ/fscBnZugHsKyS5pahJtLQ7qH4keHr63W4027+3x+b5TGIH5DjPOadH8TfCEmo/2cNctVug5jKOSuG9MnivHl+FnjvwakZ0/U7MLezpAyxtkBm4DHI7eoqOf9nnxU4kka/wBOkc5Yks2WPXrilCl78nKXu9DPnl2Pozhh2INfIvxf0ST4H/GPS/GOiIYdOvpTcGJOFBziaP6EHNfTfgCS5l8G6Qbty8624RyTkkrx1/CvL/2utJivfhjFfMoMljfRFT7PlT/OqhpKxpurnJW+iWVl8T/EvguLb/YHj3SjqFkv8KzFdwKj1zmpP2PNdntbjxJ4SumIa3cXKKT91gdjj8wKxo9RkNp8DPEW4m4SY6e7nqw3bf5VL8I0GjftSeI9Pi+WJ3vE2jp1DD+dU9mhn1bWXq3iPTtH+W4lzKekaDLUniTV/wCxtLedcGVvkjB/vHvXnmlaXd+ItQZFclj88sr84rxMdjpUpKjSV5MZ1J+Itrnixnx67hWvpPirTdWcRRSGOY/8s5Bgn6etZyfD3TRFtaa4aT+9ux+lZZ8A3cepxolx/oud3nDhlx2x6+9YRqZhTac0mgO9qtf30GnWr3Nw4SNBk+/sKLq7g0uzM1zKVijABZuSf/r1yWp+NtNu8RnTGuo1OR5hAGfXFd+JxcKMbSkkwJNK1/XtfvXFokMNqG5dkzsHpnua7FQQoBOT3PrXJ6N400wFLVrU2KE4XGCg+uK29c1n+xrIXfkNPGWAO04xnvWODrQVJ1JVObv5fIDSorG8PeJodfMypC0LxYO1jnIPetaeZLeF5nOFRSxPsK7qdaFSHPF6APormtJ8Zrq9/HaRWUi7sksWGAB3q/rXiWy0TCzM0kxGREnJ/H0rKOMoyg6ilogNaiuNT4jRFwHsJAvchwTXS6Xq9pq9v51rJuA4ZTwVPoRSo42jWfLTldgXaKx9Z8U2Oiv5UxeSbGfLQcge9Yn/AAsdN3/IPfb/ANdBmlVzDD0pcs5agdnRWXoviKy1xT9nYrIvLRvww/xrUrpp1I1I80HdAFFFFWB8NeHHN3+zp4ngT5pNP1u1vHHohAGf0NdL8Ur5NO+M/gTxK5xa3dpp9xv7bQQpqh8NdHOn+MfHXws1E7DrFtPawB+MzxkvGfxFReI7Ofxr8DdMvwrf214IuH0+/T+NYSflY/Qgfka6OpJ9qAgjI5Bry/xd8S7PwB46u4ryxuLlbuygZTBjKlS4wQfrVz4FfEOD4heA7K5aVTqNki2t7HnkOowG+jDn865H4vapd+DviRpniRbCC8hazMKLMMoWBOR7HkVnCPvWZNR2V0Lr37QGlahHZra6Re5hu4538xlHyqckD3rSk/aP0Ixvt0fUicHH3cZ968Cu7g3d3NclEjMsjSFEGFXJzge1bGj+KpdH8PatoiafazrqgUGaRcyRY/u10OjG2xgqsu59RfDy4+1+C9JuCu3zofMx6ZYnFeY/tdavFY/DGKwZgJL6+iCj2TLH+VeoeAbG40zwXo9ndJsmitUDr6E84/Wvmf406tJ8Y/jRpXgrSH82ysJfszuvK7yczP8ARQMVzxXvXOlbE8OnSi3+BvhraVuWmOoyKeqru3Z/KpPg2/8AwkH7TniTVohuhje8k3Dtlgg/kasSeILSX4heLvHEW3+xfBOl/wBlaaf4Xn27FC++cmtT9jrw5MLHXvFd0pL3sotY3P8AEFO5z/30ap6JsD174jM3lWK/w7mP44qf4eRoNOuJB99pcH6AcVb8b6Y9/pPmxKWkt23gDuO9cr4R8QJoty8VwT9mnxkjnY3r9K+YrSVDMOeps/8AIZ6VXO6p40ttKv5LOS1mdkx8ykYOa111fT5I/MW9gKdc7xXm3ie5hvNduJoJFkjJUBl6GuvMsY6VNSpSV2xnpd3bRapYPBKP3cyflkVzUeieFtOUR3VzFNKOGZ5O/wBB0qfxjfzWWgwJAxQzlUZlPIGM1g+FPDVrrkU01zM42NtCIcHp1NY4qsp140owUpW6gUfEdrpltdodLnWWGRclQc7DXYaFGdb8ILbzHcWRosn26f0rlfFWkWOjXUMFnI7sVJkDNnb6V1vgb/kX4v8Afb+dc2Bg/rc6ckldO6WwHH+GbxtI1+MS5UFjBKD/AJ9a6zx1qP2TSfsytiS5O3/gI61zfjfTjY6yZ0G1LkbwR2Ydao6vqk+vXVt8p3LGsSr6t3P4msFXlhqdTDdb6COg8EWos7C81eReilU+ijJ/WubtWTV9XEmoXIiSVi8kjHoPQfyr0iPSxb6D/Z0fUQFM+pI/xrzKwitlv0h1HekIYpIV4KnpmqxtF0YUafT8LgdhcWfhCa1MMc1vE+MLIrHcD61z/ha9fTdfiRZA0cr+S5HRh2NdEvg3w+0Xmi7cx4zu80YpdH8N6DczC5sbmaU28g/j4yP6VvLDVpVYSSjFrs9xmhr+gadqNxDd303krGCrfMF3jsM1l3GneEHhaJLiGJ8EK6yHINYXi29mvdcmhkciOJhGik8L710MfgXSY7USz3ExAXcz7wFqpT9vVmqVNO27YHK+HrhrLXrVkbP73YSP4lPFer15HpYQa1aiM5QXA2k9xnivXK0yNvkmvMAooor3QPl39qDwlf8AhTxZpPxM0MNG4ljS4dB9yZPuMfZh8pqlN4ks9C1e3+KWmWgu/CHiqP7F4j09Rn7LORh8jtzyP/r19O+I/D2n+KtEvNG1SAT2d3GY5EPoe49COorx34NfAnV/AWteIYNYv4b3w/ckRwWTKHS6AOVkdT91h0rWM9NRWD4WfCew+E+tax4ybxHt8O3EIayiyVDRMNwMgPVhnAFP1VNd+O91s0+NdM8OWbkxXM6ZaeTpkV6/q/hvSdft47fU7GG7hiOUjkHyg/SrlnZW9hbR2tpDHBBEu1I4xhVHsKlTtr1Jcb6PY8K/4Zr1D/oYLb/vwf8AGkb9nfWdPH2yz1y0luYCJYkaEgM45Az9a98oNV7afcn2MTzfw146k8b6VqHhm8nfQPE6wvAwwMq2MeYgPX6V4ivha5+AGhag9xLHqHjzxFI+n6WkB3mOFjhpfXLZ/OvqSTw3o82qrq0mm2rX6gBbkxjeMe9VPEvhe11ny9QhtLI63ZRyf2deXEe77PIy4B+lSpK+haT6nyF8QrSTw9ougfCDRD9q1WWdLvV2jOTLeSfdjJ/2Qcmvrf4e+ELfwL4O0vw/bgH7JCBIw/jkPLN+JJrx/wCAvwQ1nRPE+peMfHEe/V1nkS2V237mJ+afPv8Aw+1fQYom76IaAjPBFcjrfgRLmVrjTpFhZjkxN90n29K66iuTEYWnXjy1EM8yPgjWg+37PER6+YMVraZ8PmDq+ozrtHPlxd/qa7eiuKnk+Hg7u79QM/V9Gh1fT2s5CUHBRh/CR0rix4L120mP2aVMHjekm3Ir0Sit8Rl9Gu1KWjXYDgZvAN+YI3FxHLcsxMhZjgDtj1rqfDOmT6RpSWtwUMisxO05HJrVop0MvpUZ+0huBi+KdDfW7BY4donjfchbp7isTQPBd5Y6pFdXjQmOLLAKckt2rtaKKuApVKqrSWoCYNcz4h8GR6pK11aOILhvvAj5X9/Y109FbV8PTrR5KiugPNv+EF1nfs2RbfXzOK6jwx4WOhs8805knddpVOEA/rXQ0Vy0Mro0Z88b3A5XxN4OfVLk3lk6JMw+dH6N7/Wsm28Ga3PtgurnyrcHkeYW49hXoFFFTK6E5ubvruBwlt4Hv7XVI5kaHyI5gwy3zbQf513dFFb4bCU8OmqfUAooorqAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/9k=';
