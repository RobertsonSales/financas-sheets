/**
 * Backend Google Apps Script – CRUD da planilha de Receitas/Despesas com MÚLTIPLOS USUÁRIOS (e-mail + senha).
 * Cole este arquivo em Extensões > Apps Script da sua planilha e publique uma NOVA versão da implantação.
 *
 * Propriedades do script:
 *   API_TOKEN       (obrigatória) segredo entre a Vercel e este script.
 *   ALLOWED_EMAILS  (opcional) e-mails autorizados a criar conta, separados por vírgula. Vazio = qualquer e-mail
 *                   que confirme o código recebido pode se cadastrar.
 *   ADMIN_EMAIL     (opcional) dono dos lançamentos antigos, usado pela função atribuirLancamentosAntigos.
 *
 * Usuários ficam na aba "Usuarios" (senha guardada só como hash + salt).
 * Cada lançamento grava o e-mail do dono na coluna "Usuario"; cada usuário só vê/edita/exclui os seus.
 */
const SHEET_NAME = 'Lancamentos';
const USERS_SHEET = 'Usuarios';
const HEADERS = ['Data', 'Descrição', 'Valor', 'Tipo Despesa', 'Qt Parcelas', 'Forma Pagto', 'Tipo Receita', 'ID', 'Usuario'];
const USER_HEADERS = ['Email', 'Nome', 'Hash', 'Salt', 'Epoch', 'Criado em', 'Ativo'];
const NCOL = HEADERS.length;
const U_COL = 9; // coluna "Usuario"
// ---- Parâmetros de segurança ----
const SESSION_TTL = 21600;   // 6 h (máximo permitido pelo CacheService)
const CODE_TTL    = 900;     // código de e-mail vale 15 min
const HASH_ROUNDS = 2000;
const MAX_FAILS   = 5;       // tentativas de senha erradas...
const LOCK_SECS   = 900;     // ...bloqueiam o login por 15 min
const MIN_PW      = 8;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#0f3d3e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');        // Data como texto DD/MM/AAAA
    sh.getRange('C:C').setNumberFormat('#,##0.00');
  }
  if (sh.getRange(1, U_COL).getValue() !== 'Usuario') {
    sh.getRange(1, U_COL).setValue('Usuario').setFontWeight('bold').setBackground('#0f3d3e').setFontColor('#ffffff');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function err_(msg, code) { const o = { ok: false, error: msg }; if (code) o.code = code; return json_(o); }

function authorized_(token) {
  const t = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  return t && token === t;
}

function fmtDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(d || '');
}

function list_(email) {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, NCOL).getValues();
  return values.filter(r => r[7] !== '' && r[7] != null && normEmail_(r[8]) === email).map(r => ({
    data: fmtDate_(r[0]), descricao: String(r[1] || ''), valor: Number(r[2]) || 0,
    tipoDespesa: String(r[3] || ''), parcelas: r[4] === '' ? '' : Number(r[4]),
    formaPagto: String(r[5] || ''), tipoReceita: String(r[6] || ''), id: String(r[7])
  }));
}

// O dono (coluna Usuario) vem SEMPRE da sessão, nunca do navegador.
function toRow_(o, email) {
  return [String(o.data || ''), o.descricao || '', Number(o.valor) || 0, o.tipoDespesa || '',
    o.parcelas === '' || o.parcelas == null ? '' : Number(o.parcelas), o.formaPagto || '', o.tipoReceita || '',
    o.id || Utilities.getUuid(), email];
}

// Só encontra a linha se ela pertencer ao usuário logado.
function findRow_(sh, id, email) {
  if (sh.getLastRow() < 2) return -1;
  const vals = sh.getRange(2, 8, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++)
    if (String(vals[i][0]) === String(id) && normEmail_(vals[i][1]) === email) return i + 2;
  return -1;
}

/* ===================== USUÁRIOS ===================== */
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }
function sha_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }
function normEmail_(e) { return String(e || '').trim().toLowerCase(); }
function validEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e) && e.length <= 254; }

function hashPw_(pw, salt) {
  const saltB = Utilities.newBlob(salt).getBytes();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.newBlob(pw).getBytes().concat(saltB));
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(saltB));
  return Utilities.base64Encode(h);
}

function usersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS])
      .setFontWeight('bold').setBackground('#0f3d3e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:G').setNumberFormat('@');
  }
  return sh;
}

// Retorna { row, email, nome, hash, salt, epoch, ativo } ou null
function getUser_(email) {
  const sh = usersSheet_();
  if (sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, USER_HEADERS.length).getValues();
  for (let i = 0; i < v.length; i++) if (normEmail_(v[i][0]) === email) {
    return { row: i + 2, email, nome: String(v[i][1] || ''), hash: String(v[i][2]), salt: String(v[i][3]),
      epoch: String(v[i][4] || '0'), ativo: String(v[i][6]).toUpperCase() !== 'NAO' && String(v[i][6]).toUpperCase() !== 'NÃO' };
  }
  return null;
}

// Cria o usuário ou troca a senha (trocar a senha invalida todas as sessões daquele usuário).
function setPassword_(email, pw, nome) {
  const sh = usersSheet_(), u = getUser_(email), salt = Utilities.getUuid(), hash = hashPw_(pw, salt);
  if (u) {
    sh.getRange(u.row, 3, 1, 3).setValues([[hash, salt, String(Number(u.epoch) + 1)]]);
    if (nome && !u.nome) sh.getRange(u.row, 2).setValue(nome);
  } else {
    sh.appendRow([email, nome || '', hash, salt, '1', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), 'SIM']);
  }
  return getUser_(email);
}
function pwOk_(pw) { return typeof pw === 'string' && pw.length >= MIN_PW && pw.length <= 128; }
function checkPw_(u, pw) { return !!u && hashPw_(String(pw || ''), u.salt) === u.hash; }

function allowedToRegister_(email) {
  const raw = props_().getProperty('ALLOWED_EMAILS') || '';
  const list = raw.split(/[,;\s]+/).map(normEmail_).filter(Boolean);
  return !list.length || list.indexOf(email) >= 0;
}

/* ===================== SESSÕES ===================== */
function newSession_(u) {
  const t = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  cache_().put('s:' + sha_(t), u.email + '|' + u.epoch, SESSION_TTL);
  return t;
}
// Retorna o e-mail do usuário da sessão, ou '' se inválida
function sessionUser_(t) {
  if (!t) return '';
  const v = cache_().get('s:' + sha_(String(t)));
  if (!v) return '';
  const p = v.split('|'), u = getUser_(p[0]);
  return u && u.ativo && u.epoch === p[1] ? u.email : '';
}

/* ===================== LIMITES (por e-mail) ===================== */
function failCount_(email) { return Number(cache_().get('fails:' + email) || 0); }
function addFail_(email) { const n = failCount_(email) + 1; cache_().put('fails:' + email, String(n), LOCK_SECS); return n; }
function lockedMsg_() { return 'Muitas tentativas incorretas. Aguarde 15 minutos ou use "Esqueci minha senha".'; }
function maskEmail_(e) { const p = String(e).split('@'); return p[0].slice(0, 1) + '***@' + (p[1] || ''); }
function randCode_() { return String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000).padStart(6, '0'); }

function handleAuth_(action, body) {
  const c = cache_();
  const email = normEmail_(body.email);
  switch (action) {
    case 'auth_status':
      return json_({ ok: true, multiUser: true });

    case 'auth_login': {
      if (!validEmail_(email)) return err_('Informe um e-mail válido.');
      if (failCount_(email) >= MAX_FAILS) return err_(lockedMsg_());
      const u = getUser_(email);
      if (!u || !checkPw_(u, body.password)) {
        const n = addFail_(email);
        return err_(n >= MAX_FAILS ? lockedMsg_() : 'E-mail ou senha incorretos (' + n + '/' + MAX_FAILS + ').');
      }
      if (!u.ativo) return err_('Usuário desativado. Fale com o administrador.');
      c.remove('fails:' + email);
      return json_({ ok: true, session: newSession_(u), email: u.email, nome: u.nome });
    }

    // Usado para CRIAR CONTA e para RECUPERAR SENHA: o código vai para o próprio e-mail informado.
    case 'auth_code_request': {
      if (!validEmail_(email)) return err_('Informe um e-mail válido.');
      const exists = !!getUser_(email);
      if (!exists && !allowedToRegister_(email)) return err_('Este e-mail não está autorizado a criar conta. Fale com o administrador.');
      if (c.get('codeCool:' + email)) return err_('Aguarde 1 minuto para pedir um novo código.');
      const n = Number(c.get('codeCount:' + email) || 0);
      if (n >= 5) return err_('Limite de pedidos de código atingido. Tente novamente em 1 hora.');
      const code = randCode_(), salt = Utilities.getUuid();
      try {
        MailApp.sendEmail({
          to: email, name: 'FinançaPro',
          subject: 'FinançaPro – código de verificação: ' + code,
          body: 'Seu código de verificação é: ' + code + '\n\nEle vale por 15 minutos.\n' +
                (exists ? 'Se você não solicitou, ignore este e-mail: sua senha continua a mesma.'
                        : 'Use-o para concluir a criação da sua conta. Se não foi você, ignore este e-mail.')
        });
      } catch (e) {
        return err_('Não foi possível enviar o e-mail. No editor do Apps Script, execute a função autorizarEmail e publique uma nova versão da implantação.');
      }
      c.put('code:' + email, salt + ':' + sha_(salt + code), CODE_TTL);
      c.put('codeTries:' + email, '0', CODE_TTL);
      c.put('codeCool:' + email, '1', 60);
      c.put('codeCount:' + email, String(n + 1), 3600);
      return json_({ ok: true, email: maskEmail_(email), exists });
    }

    case 'auth_code_confirm': {
      if (!validEmail_(email)) return err_('Informe um e-mail válido.');
      const stored = c.get('code:' + email);
      if (!stored) return err_('Código expirado. Volte e peça um novo.');
      const tries = Number(c.get('codeTries:' + email) || 0) + 1;
      if (tries > 5) { c.remove('code:' + email); return err_('Muitas tentativas com código errado. Peça um novo código.'); }
      c.put('codeTries:' + email, String(tries), CODE_TTL);
      const parts = stored.split(':');
      if (sha_(parts[0] + String(body.code || '').trim()) !== parts[1]) return err_('Código incorreto.');
      if (!pwOk_(body.password)) return err_('A senha deve ter entre ' + MIN_PW + ' e 128 caracteres.');
      if (!getUser_(email) && !allowedToRegister_(email)) return err_('Este e-mail não está autorizado a criar conta.');
      const u = setPassword_(email, body.password, String(body.nome || '').trim().slice(0, 80));
      if (!u.ativo) return err_('Usuário desativado. Fale com o administrador.');
      c.remove('code:' + email); c.remove('fails:' + email);
      return json_({ ok: true, session: newSession_(u), email: u.email, nome: u.nome });
    }

    case 'auth_change': {
      const me = sessionUser_(body.session);
      if (!me) return err_('Sessão expirada. Saia e entre novamente.', 'AUTH_REQUIRED');
      if (failCount_(me) >= MAX_FAILS) return err_(lockedMsg_());
      if (!checkPw_(getUser_(me), body.oldPassword)) { addFail_(me); return err_('Senha atual incorreta.'); }
      if (!pwOk_(body.password)) return err_('A senha deve ter entre ' + MIN_PW + ' e 128 caracteres.');
      const u = setPassword_(me, body.password);
      c.remove('fails:' + me);
      return json_({ ok: true, session: newSession_(u), email: u.email, nome: u.nome });
    }

    case 'auth_logout':
      if (body.session) c.remove('s:' + sha_(String(body.session)));
      return json_({ ok: true });
  }
  return err_('Ação desconhecida');
}

/**
 * Execute UMA vez no editor, depois que o primeiro usuário criar a conta, para atribuir a ele
 * os lançamentos antigos (sem dono). Usa a propriedade ADMIN_EMAIL ou o e-mail do dono do script.
 */
function atribuirLancamentosAntigos() {
  const email = normEmail_(props_().getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail());
  if (!validEmail_(email)) throw new Error('Defina a propriedade ADMIN_EMAIL.');
  const sh = getSheet_(), last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, 8, last - 1, 2), v = rng.getValues();
  let n = 0;
  v.forEach(r => { if (r[0] !== '' && !String(r[1]).trim()) { r[1] = email; n++; } });
  rng.setValues(v);
  Logger.log(n + ' lançamento(s) atribuídos a ' + email);
}

/**
 * Execute UMA vez no editor (selecione a função e clique em Executar) para conceder a
 * permissão de envio de e-mail, necessária à criação/recuperação de senha.
 */
function autorizarEmail() {
  Logger.log('Permissão concedida. E-mails restantes hoje: ' + MailApp.getRemainingDailyQuota());
}

/* ===================== ENDPOINTS ===================== */
// GET não devolve mais dados: serve só para testar a implantação (exige o token do proxy).
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!authorized_(p.token)) return err_('Não autorizado');
  return json_({ ok: true, status: 'online', multiUsuario: true });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return err_('JSON inválido'); }
  if (!authorized_(body.token)) return err_('Não autorizado');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const action = String(body.action || '');
    if (action.indexOf('auth_') === 0) return handleAuth_(action, body);

    // Tudo abaixo exige sessão válida (login com senha)
    const me = sessionUser_(body.session);
    if (!me) return err_('Sessão expirada. Faça login novamente.', 'AUTH_REQUIRED');

    const sh = getSheet_();
    switch (action) {
      case 'list':
        return json_({ ok: true, data: list_(me) });
      case 'create': {
        const rows = (body.rows || []).map(o => toRow_(o, me));
        if (!rows.length) return err_('Nada a inserir');
        const start = sh.getLastRow() + 1;
        sh.getRange(start, 1, rows.length, NCOL).setNumberFormat('@').setValues(rows);
        sh.getRange(start, 3, rows.length, 1).setNumberFormat('#,##0.00');
        sh.getRange(start, 5, rows.length, 1).setNumberFormat('0');
        return json_({ ok: true, inserted: rows.length });
      }
      case 'update': {
        const r = findRow_(sh, body.row && body.row.id, me);
        if (r < 0) return err_('Registro não encontrado');
        sh.getRange(r, 1, 1, NCOL).setValues([toRow_(body.row, me)]);
        return json_({ ok: true });
      }
      case 'delete': {
        const ids = [].concat(body.ids || body.id || []);
        let n = 0;
        ids.forEach(id => { const r = findRow_(sh, id, me); if (r > 0) { sh.deleteRow(r); n++; } });
        return json_({ ok: true, deleted: n });
      }
      default:
        return err_('Ação desconhecida');
    }
  } finally {
    lock.releaseLock();
  }
}
