/**
 * Backend Google Apps Script – CRUD da planilha de Receitas/Despesas + autenticação por senha.
 * Cole este arquivo em Extensões > Apps Script da sua planilha.
 *
 * Propriedades do script: API_TOKEN (obrigatória, segredo entre a Vercel e este script) e
 * RECOVERY_EMAIL (e-mail de quem usa o app, que receberá os códigos de senha; opcional, padrão = dono do script).
 * Propriedades criadas automaticamente: AUTH_HASH, AUTH_SALT, AUTH_EPOCH (senha do app, guardada só como hash).
 */
const SHEET_NAME = 'Lancamentos';
const HEADERS = ['Data', 'Descrição', 'Valor', 'Tipo Despesa', 'Qt Parcelas', 'Forma Pagto', 'Tipo Receita', 'ID'];

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

function list_() {
  const values = getSheet_().getDataRange().getValues();
  values.shift();
  return values.filter(r => r[7] !== '' && r[7] != null).map(r => ({
    data: fmtDate_(r[0]), descricao: String(r[1] || ''), valor: Number(r[2]) || 0,
    tipoDespesa: String(r[3] || ''), parcelas: r[4] === '' ? '' : Number(r[4]),
    formaPagto: String(r[5] || ''), tipoReceita: String(r[6] || ''), id: String(r[7])
  }));
}

function toRow_(o) {
  return [String(o.data || ''), o.descricao || '', Number(o.valor) || 0, o.tipoDespesa || '',
    o.parcelas === '' || o.parcelas == null ? '' : Number(o.parcelas), o.formaPagto || '', o.tipoReceita || '',
    o.id || Utilities.getUuid()];
}

function findRow_(sh, id) {
  const ids = sh.getRange(2, 8, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/* ===================== AUTENTICAÇÃO ===================== */
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }
function sha_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }

function hashPw_(pw, salt) {
  const saltB = Utilities.newBlob(salt).getBytes();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.newBlob(pw).getBytes().concat(saltB));
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(saltB));
  return Utilities.base64Encode(h);
}

function isConfigured_() { return !!props_().getProperty('AUTH_HASH'); }
function epoch_() { return props_().getProperty('AUTH_EPOCH') || '0'; }
function sessKey_(t) { return 's:' + epoch_() + ':' + sha_(t); }

function newSession_() {
  const t = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  cache_().put(sessKey_(t), '1', SESSION_TTL);
  return t;
}
function validSession_(t) { return !!t && isConfigured_() && cache_().get(sessKey_(String(t))) === '1'; }

function setPassword_(pw) {   // trocar a senha invalida todas as sessões (AUTH_EPOCH muda)
  const salt = Utilities.getUuid();
  props_().setProperties({ AUTH_SALT: salt, AUTH_HASH: hashPw_(pw, salt), AUTH_EPOCH: String(Number(epoch_()) + 1) });
}
function pwOk_(pw) { return typeof pw === 'string' && pw.length >= MIN_PW && pw.length <= 128; }
function checkPw_(pw) { return hashPw_(String(pw || ''), props_().getProperty('AUTH_SALT')) === props_().getProperty('AUTH_HASH'); }

function failCount_() { return Number(cache_().get('fails') || 0); }
function addFail_() { const n = failCount_() + 1; cache_().put('fails', String(n), LOCK_SECS); return n; }
function lockedMsg_() { return 'Muitas tentativas incorretas. Aguarde 15 minutos ou use "Esqueci minha senha".'; }

// E-mails que recebem o código de criação/recuperação de senha.
// Definidos pelo dono do script na propriedade RECOVERY_EMAIL (um ou vários, separados por vírgula).
// Se a propriedade não existir, usa o e-mail do dono do script.
function recoveryEmails_() {
  const raw = props_().getProperty('RECOVERY_EMAIL') || '';
  const list = raw.split(/[,;\s]+/).filter(e => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e));
  if (list.length) return list;
  const owner = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
  return owner ? [owner] : [];
}
function maskEmail_(e) { const p = String(e).split('@'); return p[0].slice(0, 1) + '***@' + (p[1] || ''); }
function randCode_() { return String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000).padStart(6, '0'); }

function handleAuth_(action, body) {
  const c = cache_();
  switch (action) {
    case 'auth_status':
      return json_({ ok: true, configured: isConfigured_() });

    case 'auth_login': {
      if (!isConfigured_()) return err_('Senha ainda não criada.', 'AUTH_REQUIRED');
      if (failCount_() >= MAX_FAILS) return err_(lockedMsg_());
      if (!checkPw_(body.password)) {
        const n = addFail_();
        return err_(n >= MAX_FAILS ? lockedMsg_() : 'Senha incorreta (' + n + '/' + MAX_FAILS + ').');
      }
      c.remove('fails');
      return json_({ ok: true, session: newSession_() });
    }

    case 'auth_code_request': {   // usado para criar a 1ª senha e para recuperar senha esquecida
      if (c.get('codeCool')) return err_('Aguarde 1 minuto para pedir um novo código.');
      const n = Number(c.get('codeCount') || 0);
      if (n >= 5) return err_('Limite de pedidos de código atingido. Tente novamente em 1 hora.');
      const emails = recoveryEmails_();
      if (!emails.length) return err_('Nenhum e-mail de recuperação configurado. Defina a propriedade RECOVERY_EMAIL no script.');
      const code = randCode_(), salt = Utilities.getUuid();
      try {
        MailApp.sendEmail({
          to: emails.join(','), name: 'FinançaPro',
          subject: 'FinançaPro – código de verificação: ' + code,
          body: 'Seu código de verificação é: ' + code + '\n\nEle vale por 15 minutos.\n' +
                'Se você não solicitou, ignore este e-mail: sua senha continua a mesma.'
        });
      } catch (e) {
        return err_('Não foi possível enviar o e-mail. No editor do Apps Script, execute a função autorizarEmail e publique uma nova versão da implantação.');
      }
      c.put('code', salt + ':' + sha_(salt + code), CODE_TTL);
      c.put('codeTries', '0', CODE_TTL);
      c.put('codeCool', '1', 60);
      c.put('codeCount', String(n + 1), 3600);
      return json_({ ok: true, email: emails.map(maskEmail_).join(', ') });
    }

    case 'auth_code_confirm': {
      const stored = c.get('code');
      if (!stored) return err_('Código expirado. Volte e peça um novo.');
      const tries = Number(c.get('codeTries') || 0) + 1;
      if (tries > 5) { c.remove('code'); return err_('Muitas tentativas com código errado. Peça um novo código.'); }
      c.put('codeTries', String(tries), CODE_TTL);
      const parts = stored.split(':');
      if (sha_(parts[0] + String(body.code || '').trim()) !== parts[1]) return err_('Código incorreto.');
      if (!pwOk_(body.password)) return err_('A senha deve ter entre ' + MIN_PW + ' e 128 caracteres.');
      setPassword_(body.password);
      c.remove('code'); c.remove('fails');
      return json_({ ok: true, session: newSession_() });
    }

    case 'auth_change': {
      if (!validSession_(body.session)) return err_('Sessão expirada. Saia e entre novamente.', 'AUTH_REQUIRED');
      if (failCount_() >= MAX_FAILS) return err_(lockedMsg_());
      if (!checkPw_(body.oldPassword)) { addFail_(); return err_('Senha atual incorreta.'); }
      if (!pwOk_(body.password)) return err_('A senha deve ter entre ' + MIN_PW + ' e 128 caracteres.');
      setPassword_(body.password);
      c.remove('fails');
      return json_({ ok: true, session: newSession_() });
    }

    case 'auth_logout':
      if (body.session) c.remove(sessKey_(String(body.session)));
      return json_({ ok: true });
  }
  return err_('Ação desconhecida');
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
  return json_({ ok: true, status: 'online', senhaConfigurada: isConfigured_() });
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
    if (!validSession_(body.session)) return err_('Sessão expirada. Faça login novamente.', 'AUTH_REQUIRED');

    const sh = getSheet_();
    switch (action) {
      case 'list':
        return json_({ ok: true, data: list_() });
      case 'create': {
        const rows = (body.rows || []).map(toRow_);
        if (!rows.length) return err_('Nada a inserir');
        const start = sh.getLastRow() + 1;
        sh.getRange(start, 1, rows.length, HEADERS.length).setNumberFormat('@').setValues(rows);
        sh.getRange(start, 3, rows.length, 1).setNumberFormat('#,##0.00');
        sh.getRange(start, 5, rows.length, 1).setNumberFormat('0');
        return json_({ ok: true, inserted: rows.length });
      }
      case 'update': {
        const r = findRow_(sh, body.row.id);
        if (r < 0) return err_('Registro não encontrado');
        sh.getRange(r, 1, 1, HEADERS.length).setValues([toRow_(body.row)]);
        return json_({ ok: true });
      }
      case 'delete': {
        const ids = [].concat(body.ids || body.id || []);
        let n = 0;
        ids.forEach(id => { const r = findRow_(sh, id); if (r > 0) { sh.deleteRow(r); n++; } });
        return json_({ ok: true, deleted: n });
      }
      default:
        return err_('Ação desconhecida');
    }
  } finally {
    lock.releaseLock();
  }
}
