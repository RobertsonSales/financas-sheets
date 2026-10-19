/**
 * Backend Google Apps Script – CRUD da planilha de Receitas/Despesas
 * Cole este arquivo em Extensões > Apps Script da sua planilha.
 * Defina a propriedade do script API_TOKEN (Configurações do projeto > Propriedades do script).
 */
const SHEET_NAME = 'Lancamentos';
const HEADERS = ['Data', 'Descrição', 'Valor', 'Tipo Despesa', 'Qt Parcelas', 'Forma Pagto', 'Tipo Receita', 'ID'];

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

function doGet(e) {
  if (!authorized_(e.parameter.token)) return json_({ ok: false, error: 'Não autorizado' });
  return json_({ ok: true, data: list_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'JSON inválido' }); }
  if (!authorized_(body.token)) return json_({ ok: false, error: 'Não autorizado' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    switch (body.action) {
      case 'create': {
        const rows = (body.rows || []).map(toRow_);
        if (!rows.length) return json_({ ok: false, error: 'Nada a inserir' });
        const start = sh.getLastRow() + 1;
        sh.getRange(start, 1, rows.length, HEADERS.length).setNumberFormat('@').setValues(rows);
        sh.getRange(start, 3, rows.length, 1).setNumberFormat('#,##0.00');
        sh.getRange(start, 5, rows.length, 1).setNumberFormat('0');
        return json_({ ok: true, inserted: rows.length });
      }
      case 'update': {
        const r = findRow_(sh, body.row.id);
        if (r < 0) return json_({ ok: false, error: 'Registro não encontrado' });
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
        return json_({ ok: false, error: 'Ação desconhecida' });
    }
  } finally {
    lock.releaseLock();
  }
}
