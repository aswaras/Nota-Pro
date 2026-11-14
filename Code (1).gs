/**
 * NOTA PRO - Backend Google Apps Script
 * Database: Google Spreadsheet (script ini dipasang lewat Extensions > Apps Script di spreadsheet).
 *
 * Tab yang dibuat otomatis: Nota, Belanja, Arsip, Setting
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 * Setelah mengubah kode, deploy ulang lewat Manage deployments > Edit > New version.
 */

const SHEETS = {
  Nota:    ['ID','No Nota','Tanggal','Pelanggan','WA','Deadline','Qty','Sub Total','DP','Pelunasan','Sisa Tagihan','Status','Arsip','JSON'],
  Belanja: ['ID','Tanggal','Vendor','Pesanan','Total','Arsip'],
  Arsip:   ['ID','Tanggal','Label','Jml Nota','Jml Belanja','Qty','Omset','Belanja','Keuntungan','Tagihan'],
  Setting: ['Key','Value']
};
// kolom angka (1-based); kolom lain diformat teks supaya ID, tanggal, dan nomor WA tidak berubah
const NUM = { Nota: [7,8,9,10,11], Belanja: [5], Arsip: [4,5,6,7,8,9,10], Setting: [] };

/* ---------- Entry point ---------- */
function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'load';
  try {
    if (a === 'load') return out_(load_());
    return out_({ ok: false, error: 'aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const b = JSON.parse(e.postData.contents);
    return out_(handle_(b));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** Jalankan sekali dari editor untuk membuat semua tab dan memberi izin akses. */
function setup() {
  Object.keys(SHEETS).forEach(sh_);
}

function handle_(b) {
  switch (b.a) {
    case 'note':   return upNote_(b.note);
    case 'del':    return delRow_('Nota', b.id);
    case 'exp':    return upExp_(b.exp);
    case 'delexp': return delRow_('Belanja', b.id);
    case 'arc':    return arc_(b.arc, b.nids || [], b.eids || []);
    case 'unarc':  return unarc_(b.id);
    case 'set':    return upSet_(b.set);
    default:       return { ok: false, error: 'aksi tidak dikenal' };
  }
}

/* ---------- Helper sheet ---------- */
function sh_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const H = SHEETS[name];
  let s = ss.getSheetByName(name);

  if (s && s.getLastRow() > 0) {
    const w = Math.min(H.length, s.getMaxColumns());
    const cur = s.getRange(1, 1, 1, w).getValues()[0];
    if (cur.join('|') !== H.join('|')) {
      // format tab lama berbeda: amankan dengan mengganti nama, lalu buat tab baru
      s.setName(name + '_lama_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd_HHmmss'));
      s = null;
    }
  }
  if (!s) s = ss.insertSheet(name);

  if (s.getLastRow() === 0) {
    if (s.getMaxColumns() < H.length) s.insertColumnsAfter(s.getMaxColumns(), H.length - s.getMaxColumns());
    for (let c = 1; c <= H.length; c++) {
      if (NUM[name].indexOf(c) < 0) s.getRange(1, c, s.getMaxRows(), 1).setNumberFormat('@');
    }
    s.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }
  return s;
}

function findRow_(s, id) {
  const n = s.getLastRow();
  if (n < 2) return 0;
  const ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function rows_(name) {
  const s = sh_(name), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, SHEETS[name].length).getValues();
}

const d_ = v => (v instanceof Date ? v.toISOString() : String(v == null ? '' : v));

/* ---------- Load ---------- */
function load_() {
  const notes = rows_('Nota').map(r => {
    let o;
    try { o = JSON.parse(r[13]); } catch (x) { return null; }
    if (r[12]) o.arc = String(r[12]); else delete o.arc;
    return o;
  }).filter(Boolean);

  const exps = rows_('Belanja').map(r => ({
    id: String(r[0]), date: d_(r[1]), vendor: String(r[2]), name: String(r[3] || ''),
    amt: Number(r[4]) || 0, arc: r[5] ? String(r[5]) : ''
  }));

  const arcs = rows_('Arsip').map(r => ({
    id: String(r[0]), date: d_(r[1]), label: String(r[2]), nota: Number(r[3]) || 0, exp: Number(r[4]) || 0,
    qty: Number(r[5]) || 0, omset: Number(r[6]) || 0, bel: Number(r[7]) || 0, untung: Number(r[8]) || 0, tag: Number(r[9]) || 0
  }));

  const settings = {};
  rows_('Setting').forEach(r => { if (r[0]) settings[r[0]] = r[1]; });

  return { ok: true, notes: notes, exps: exps, arcs: arcs, settings: settings };
}

/* ---------- Nota ---------- */
function calc_(n) {
  const items = n.items || [];
  const sub = items.reduce((a, i) => a + (+i.q || 0) * (+i.p || 0) * (1 - (+i.dc || 0) / 100), 0);
  const paid = (n.pays || []).reduce((a, p) => a + (+p.a || 0), 0);
  const qty = items.reduce((a, i) => a + (+i.q || 0), 0);
  return { sub: sub, paid: paid, qty: qty, sisa: Math.max(0, sub - (+n.dp || 0) - paid) };
}

function upNote_(n) {
  const s = sh_('Nota'), c = calc_(n), r = findRow_(s, n.id);
  const copy = JSON.parse(JSON.stringify(n));
  delete copy.arc; // status arsip disimpan di kolom Arsip
  const row = [n.id, n.no, n.date, n.name, n.wa || '', n.deadline || '', c.qty, c.sub, +n.dp || 0, c.paid, c.sisa,
               c.sisa ? 'Belum Lunas' : 'Lunas', n.arc || '', JSON.stringify(copy)];
  if (r) {
    row[12] = s.getRange(r, 13).getValue(); // nota yang sudah ada: status arsip tidak diubah oleh edit biasa
    s.getRange(r, 1, 1, row.length).setValues([row]);
  } else {
    s.appendRow(row);
  }
  return { ok: true };
}

function delRow_(name, id) {
  const s = sh_(name), r = findRow_(s, id);
  if (r) s.deleteRow(r);
  return { ok: true };
}

/* ---------- Belanja ---------- */
function upExp_(x) {
  const s = sh_('Belanja'), r = findRow_(s, x.id);
  const row = [x.id, x.date, x.vendor, x.name || '', +x.amt || 0, x.arc || ''];
  if (r) {
    row[5] = s.getRange(r, 6).getValue();
    s.getRange(r, 1, 1, row.length).setValues([row]);
  } else {
    s.appendRow(row);
  }
  return { ok: true };
}

/* ---------- Arsip ---------- */
function arc_(a, nids, eids) {
  const s = sh_('Arsip'), r = findRow_(s, a.id);
  const row = [a.id, a.date, a.label, a.nota, a.exp, a.qty, a.omset, a.bel, a.untung, a.tag];
  if (r) s.getRange(r, 1, 1, row.length).setValues([row]); else s.appendRow(row);
  setArc_('Nota', 13, nids, a.id);
  setArc_('Belanja', 6, eids, a.id);
  return { ok: true };
}

function setArc_(name, col, ids, val) {
  if (!ids || !ids.length) return;
  const s = sh_(name), n = s.getLastRow();
  if (n < 2) return;
  const want = {};
  ids.forEach(i => { want[String(i)] = 1; });
  const idv = s.getRange(2, 1, n - 1, 1).getValues();
  const rg = s.getRange(2, col, n - 1, 1), cv = rg.getValues();
  let changed = false;
  idv.forEach((x, i) => { if (want[String(x[0])]) { cv[i][0] = val; changed = true; } });
  if (changed) rg.setValues(cv);
}

function unarc_(id) {
  [['Nota', 13], ['Belanja', 6]].forEach(p => {
    const s = sh_(p[0]), n = s.getLastRow();
    if (n < 2) return;
    const rg = s.getRange(2, p[1], n - 1, 1), cv = rg.getValues();
    let changed = false;
    cv.forEach(r => { if (String(r[0]) === String(id)) { r[0] = ''; changed = true; } });
    if (changed) rg.setValues(cv);
  });
  delRow_('Arsip', id);
  return { ok: true };
}

/* ---------- Setting ---------- */
function upSet_(o) {
  const s = sh_('Setting'), n = s.getLastRow(), m = {};
  if (n > 1) s.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0]) m[r[0]] = r[1]; });
  Object.keys(o || {}).forEach(k => { if (k !== 'api' && o[k] !== undefined && o[k] !== null) m[k] = o[k]; });
  const rows = Object.keys(m).map(k => [k, m[k]]);
  if (rows.length) s.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true };
}

/* ---------- Output ---------- */
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
