/* Gestor de Cartera · capa de datos sobre el libro de Excel (Office.js).
   Expone las funciones api_* que usa la interfaz (taskpane.html) por medio de
   window.google.script.run, el mismo contrato de la versión Sheets.
   Opción B: todo dato del sistema vive en pestañas «muy ocultas» del mismo
   libro. En la hoja de la cartera solo se agregan comentarios en la columna
   «Comentario Créditos»; nada más se escribe ni se formatea ahí. */
(() => {
'use strict';

/* ---------- estructura fija de la cartera ---------- */
const COLS = ['Organization', 'KAM ID', 'Ejecutivo', 'Customer ID', 'Legal Name', 'Tax Reg No', 'Cust. Address1', 'State Name', 'City Name', 'Cust. Telephone', 'Cust. Mail Address', 'Term ID', 'Report Code', 'Credit Limit', 'Creation Date', 'Last Purchase', 'Clasif.', 'Comentario Créditos'];
const KEYS = ['org', 'kam', 'exec', 'id', 'legal', 'tax', 'addr', 'state', 'city', 'phone', 'mail', 'term', 'report', 'credit', 'created', 'last', 'clasif', 'comment'];
const C = Object.fromEntries(KEYS.map((k, i) => [k, i]));

/* ---------- pestañas internas ---------- */
const SH = { cfg: '_Config', users: '_Usuarios', asg: '_Asignaciones', seg: '_Seguimiento', cm: '_ClienteMarcas', mar: '_Marcas', ord: '_Ordenes', ven: '_Ventas', his: '_Historial', tel: '_Telefonos' };
const HDR = {
  cfg: ['Clave', 'Valor'],
  users: ['Nombre', 'Rol', 'Activo', 'PinHash', 'Salt', 'CambiarPin', 'Intentos', 'BloqueadoHasta'],
  asg: ['ID', 'Usuario', 'Desde fila', 'Hasta fila', 'Creado por', 'Fecha', 'Estado'],
  seg: ['Customer ID', 'Revisado', 'Resultado', 'Motivo', 'Intentos', 'Próximo contacto', 'Tipo de negocio', 'Gestionado por', 'Primera gestión', 'Última gestión', 'Último contacto', 'Depurado', 'Teléfono corregido'],
  cm: ['Customer ID', 'Marca', 'Relación', 'Registrado por', 'Fecha'],
  mar: ['Marca', 'Línea', 'Activa'],
  ord: ['OV', 'Customer ID', 'Fecha OV', 'Total sin IGV', 'Estado', 'Ejecutivo', 'Fecha terminado', 'Terminado por', 'Nota'],
  ven: ['OV', 'Customer ID', 'Fecha OV', 'Marca', 'Monto sin IGV', 'Ejecutivo'],
  his: ['Customer ID', 'Fecha', 'Usuario', 'Acción', 'Icono'],
  tel: ['Customer ID', 'Teléfono original', 'Teléfono corregido', 'Corregido por', 'Fecha']
};
const S_ = Object.fromEntries(HDR.seg.map((h, i) => [h, i]));
const CATALOGO = {"Cómputo": ["ASUS", "Dell", "HP", "Lenovo", "MSI", "Huawei"], "Gaming": ["ROG", "HyperX", "Razer", "Primus Gaming"], "Componentes": ["AMD", "Intel", "Gigabyte"], "Almacenamiento y memoria": ["A-Data", "Kingston", "PNY", "SanDisk", "Seagate", "Western Digital", "QNAP"], "Servidores y centro de datos": ["HPE", "Dell EMC"], "Redes y conectividad": ["Ubiquiti", "TP-Link", "Meraki", "Ruckus", "Nexxt Solutions Home", "Nexxt Solutions Infrastructure", "Furukawa", "Starlink", "Riverbed"], "Energía": ["APC", "Salicru", "Forza"], "Seguridad electrónica": ["Hikvision", "EZVIZ", "Hanwha", "Axis", "VideoIQ", "Honeywell", "Firelite"], "Hogar inteligente": ["Ring", "Eufy", "Philips Hue", "Wiz"], "Punto de venta y código de barras": ["Zebra", "Newland", "Unitech", "Logic Controls", "Elo", "Custom"], "Impresión": ["Epson", "Xerox"], "Colaboración y videoconferencia": ["Jabra", "Poly", "Logitech", "Screenbeam", "Promethean"], "Audio profesional": ["AKG", "JBL Pro", "Crown", "Soundcraft"], "Audio de consumo": ["JBL", "Harman Kardon", "Soundcore"], "Pantallas y monitores": ["LG", "Viewsonic", "Philips", "Haier"], "Accesorios y periféricos": ["Anker", "Belkin", "Gear4", "Zagg", "Klip Xtreme", "Xtech", "StarTech.com", "Wacom"], "Móviles y wearables": ["Motorola", "Realme", "Oura"], "Software y ciberseguridad": ["Microsoft", "Bitdefender", "ESET", "McAfee", "BitTitan"], "Por definir": ["ELISE", "Folksafe", "Cyrus", "Nebula"]};
const CFG_DEF = { cc: '51', mobLen: 9, lockRow: true, scopeOwn: true, minInt: 2, poll: 20, goal: 40, salesGoal: 0, waMode: 'web', waMsg: '', mailMode: 'outlook', mailSubj: '' };
const NUM_CFG = ['mobLen', 'minInt', 'poll', 'goal', 'salesGoal'], BOOL_CFG = ['lockRow', 'scopeOwn'];
const ESTADOS = ['Despachando', 'Incidencia', 'Terminado', 'Anulada'];

/* ---------- utilidades ---------- */
const p2 = n => String(n).padStart(2, '0');
const dayStr = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
const TODAY = () => dayStr(new Date());
const nowStr = () => { const d = new Date(); return `${dayStr(d)} ${p2(d.getHours())}:${p2(d.getMinutes())}`; };
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const str = x => x == null ? '' : String(x);
const num = x => { const n = +String(x == null ? '' : x).replace(/[^\d.-]/g, ''); return isFinite(n) ? n : 0; };
const err = m => { throw new Error(m); };
const ICON = { 'Contestó': 'check', 'No contesta': 'x', 'Volver a llamar': 'clock', 'Número errado': 'hash', 'No interesado': 'ban' };

/* fechas de la cartera: número de serie de Excel, dd/mm/aaaa o aaaa-mm-dd → aaaa-mm-dd */
function isoDate(value, text) {
  if (typeof value === 'number' && value > 20000 && value < 80000) return dayStr(new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 864e5 + 12 * 36e5));
  const t = str(text).trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${p2(m[2])}-${p2(m[3])}`;
  m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/); if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${p2(m[2])}-${p2(m[1])}`;
  return t;
}
async function sha(text) {
  const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join('');
}
const randId = () => [...crypto.getRandomValues(new Uint8Array(12))].map(x => x.toString(16).padStart(2, '0')).join('');

/* ---------- sesión (por dispositivo) ---------- */
const SESS_KEY = 'gcx:sesion';
const SESS = {};
try { const s = JSON.parse(localStorage.getItem(SESS_KEY) || 'null'); if (s && s.exp > Date.now()) SESS[s.token] = s.name; } catch (e) {}
const saveSess = (token, name) => { SESS[token] = name; try { localStorage.setItem(SESS_KEY, JSON.stringify({ token, name, exp: Date.now() + 12 * 36e5 })); } catch (e) {} };
const dropSess = token => { delete SESS[token]; try { localStorage.removeItem(SESS_KEY); } catch (e) {} };

/* ---------- acceso a Excel ---------- */
const run = fn => Excel.run(fn);
async function sheet(ctx, name) { const ws = ctx.workbook.worksheets.getItemOrNullObject(name); await ctx.sync(); return ws.isNullObject ? null : ws; }
async function readAll(ctx, ws, prop) {
  const u = ws.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount,columnIndex,columnCount'); await ctx.sync();
  if (u.isNullObject) return [];
  const rg = ws.getRangeByIndexes(0, 0, u.rowIndex + u.rowCount, u.columnIndex + u.columnCount); rg.load(prop || 'values'); await ctx.sync();
  return rg[prop || 'values'];
}
async function readAux(ctx, key) { const ws = await sheet(ctx, SH[key]); if (!ws) return { ws: null, rows: [] }; const v = await readAll(ctx, ws); return { ws, rows: v.slice(1) }; }
/* escribe una fila como texto (las pestañas internas guardan todo como texto) */
function writeRow(ws, excelRow, arr) {
  const rg = ws.getRangeByIndexes(excelRow, 0, 1, arr.length);
  rg.numberFormat = [arr.map(() => '@')];
  rg.values = [arr.map(str)];
}
async function appendRows(ctx, key, rows) {
  if (!rows.length) return;
  const ws = await sheet(ctx, SH[key]);
  /* solo se mide hasta dónde hay datos; no se relee la pestaña completa (el historial crece mucho) */
  const u = ws.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount'); await ctx.sync();
  const at = u.isNullObject ? 1 : Math.max(u.rowIndex + u.rowCount, 1);
  const rg = ws.getRangeByIndexes(at, 0, rows.length, rows[0].length);
  rg.numberFormat = rows.map(r => r.map(() => '@'));
  rg.values = rows.map(r => r.map(str));
  await ctx.sync();
}
/* actualiza la fila que cumpla `match` leyendo la pestaña en ese momento (otro usuario pudo agregar filas) */
async function upsert(ctx, key, match, build) {
  const ws = await sheet(ctx, SH[key]);
  const v = await readAll(ctx, ws);
  let i = v.findIndex((r, j) => j > 0 && match(r));
  const cur = i > 0 ? v[i].map(str) : null;
  if (i < 1) i = Math.max(v.length, 1);
  const arr = build(cur);
  writeRow(ws, i, arr); await ctx.sync();
  return arr;
}
async function bumpSeq(ctx) {
  await upsert(ctx, 'cfg', r => str(r[0]) === 'seq', cur => ['seq', String((cur ? num(cur[1]) : 0) + 1)]);
}

/* ---------- la cartera ---------- */
function headerProblems(row) {
  const got = row.map(x => str(x).trim()); while (got.length && !got[got.length - 1]) got.pop();
  const g = got.map(norm), out = [];
  COLS.forEach((c, i) => { const j = g.indexOf(norm(c)); if (j < 0) out.push(`Falta la columna «${c}».`); else if (j !== i) out.push(`«${c}» está en la columna ${j + 1}; debe estar en la ${i + 1}.`); });
  got.forEach(h => { if (h && !COLS.some(c => norm(c) === norm(h))) out.push(`Columna no prevista: «${h}».`); });
  return out;
}
async function findCartera(ctx) {
  const wss = ctx.workbook.worksheets; wss.load('items/name,items/visibility'); await ctx.sync();
  const vis = wss.items.filter(w => w.visibility === 'Visible' && w.name[0] !== '_');
  const order = [...vis.filter(w => w.name === 'Hoja1'), ...vis.filter(w => w.name !== 'Hoja1')];
  let first = null;
  for (const w of order) {
    const h = w.getRangeByIndexes(0, 0, 1, 30); h.load('values'); await ctx.sync();
    const probs = headerProblems(h.values[0]);
    if (!probs.length) return { ws: w, name: w.name, problems: [] };
    if (!first && h.values[0].some(x => norm(x) === 'customer id')) first = { ws: w, name: w.name, problems: probs };
  }
  return first || { ws: null, name: '', problems: ['No se encontró la hoja de la cartera: la fila 1 debe tener las 18 columnas de la plantilla.'] };
}
function clientKeys(ids) {
  const cnt = {};
  return ids.map((id, i) => { id = str(id).trim(); if (!id) return 'fila:' + (i + 2); cnt[id] = (cnt[id] || 0) + 1; return cnt[id] > 1 ? id + '#' + cnt[id] : id; });
}
async function readCartera(ctx) {
  const f = await findCartera(ctx);
  if (f.problems.length) { const e = new Error('ESTRUCTURA: ' + f.problems.join('\n')); e.problems = f.problems; throw e; }
  const u = f.ws.getUsedRangeOrNullObject(true); u.load('rowIndex,rowCount'); await ctx.sync();
  const n = u.isNullObject ? 1 : u.rowIndex + u.rowCount;
  const rg = f.ws.getRangeByIndexes(0, 0, Math.max(1, n), COLS.length); rg.load('values,text'); await ctx.sync();
  const ks = clientKeys(rg.values.slice(1).map(r => r[C.id]));
  const rows = [];
  rg.values.slice(1).forEach((vals, i) => {
    const tx = rg.text[i + 1];
    if (!vals.some(x => str(x).trim() !== '')) return;
    const v = tx.map(str);
    v[C.created] = isoDate(vals[C.created], tx[C.created]);
    v[C.last] = isoDate(vals[C.last], tx[C.last]);
    if (typeof vals[C.credit] === 'number') v[C.credit] = String(vals[C.credit]);
    v[C.id] = str(vals[C.id]).trim() || v[C.id];
    rows.push({ k: ks[i], r: i + 2, v });
  });
  return { name: f.name, ws: f.ws, rows };
}

/* ---------- lectura completa ---------- */
let DB = null;
async function loadDB(ctx) {
  const cart = await readCartera(ctx);
  const R = {}; for (const k in SH) R[k] = (await readAux(ctx, k)).rows;
  const cfg = Object.assign({}, CFG_DEF);
  let seq = 0;
  R.cfg.forEach(r => { const k = str(r[0]), v = str(r[1]); if (k === 'seq') seq = num(v); else if (k in CFG_DEF) cfg[k] = NUM_CFG.includes(k) ? num(v) : BOOL_CFG.includes(k) ? v === 'SI' : v; });
  const users = R.users.filter(r => str(r[0]).trim()).map((r, i) => ({ row: i + 1, name: str(r[0]).trim(), role: /admin/i.test(r[1]) ? 'admin' : 'ejecutivo', active: str(r[2]) !== 'NO', hash: str(r[3]), salt: str(r[4]), mustChange: str(r[5]) === 'SI', tries: num(r[6]), lockUntil: num(r[7]) }));
  const assigns = R.asg.filter(r => str(r[0]) && str(r[6]) !== 'Liberada').map(r => ({ id: str(r[0]), user: str(r[1]), from: num(r[2]), to: num(r[3]), by: str(r[4]), date: str(r[5]) }));
  const brands = R.mar.filter(r => str(r[0]).trim() && str(r[2]) !== 'NO').map(r => ({ m: str(r[0]).trim(), c: str(r[1]) || 'Por definir' }));
  const track = {};
  const T = k => track[k] || (track[k] = { rev: false, res: '', int: 0, buy: false, next: '', own: '', first: '', last: '', lastc: '', dep: '', motivo: '', monto: '', fix: '', tipo: '', brands: [], bmeta: {}, hist: [] });
  R.seg.forEach(a => { const k = str(a[0]); if (!k) return; Object.assign(T(k), { rev: str(a[S_['Revisado']]) === 'SI', res: str(a[S_['Resultado']]), motivo: str(a[S_['Motivo']]), int: num(a[S_['Intentos']]), next: str(a[S_['Próximo contacto']]), tipo: str(a[S_['Tipo de negocio']]), own: str(a[S_['Gestionado por']]), first: str(a[S_['Primera gestión']]), last: str(a[S_['Última gestión']]), lastc: str(a[S_['Último contacto']]), dep: str(a[S_['Depurado']]), fix: str(a[S_['Teléfono corregido']]) }); });
  R.cm.forEach(a => { const k = str(a[0]), b = str(a[1]); if (!k || !b || str(a[2]) === 'quitada') return; const t = T(k); if (!t.brands.includes(b)) t.brands.push(b); t.bmeta[b] = { rel: str(a[2]) || 'habitual', by: str(a[3]), d: str(a[4]) }; });
  R.his.forEach(a => { const k = str(a[0]); if (k) T(k).hist.push({ d: str(a[1]), who: str(a[2]), txt: str(a[3]), k: str(a[4]) || 'clock', cm: str(a[4]) === 'msg' }); });
  Object.values(track).forEach(t => { t.hist.sort((x, y) => y.d.localeCompare(x.d)); t.hist = t.hist.slice(0, 40); });
  const orders = R.ord.filter(a => str(a[0])).map(a => ({ ov: str(a[0]), k: str(a[1]), d: str(a[2]), tot: num(a[3]), est: ESTADOS.includes(str(a[4])) ? str(a[4]) : 'Despachando', by: str(a[5]), fin: str(a[6]), finBy: str(a[7]), nota: str(a[8]) }));
  const sales = R.ven.filter(a => str(a[0])).map(a => ({ ov: str(a[0]), k: str(a[1]), d: str(a[2]), b: str(a[3]), m: num(a[4]), by: str(a[5]) }));
  orders.forEach(o => { if (o.est === 'Anulada') return; const t = T(o.k); t.buy = true; t.monto = String(num(t.monto) + o.tot); });
  DB = { cart, cfg, seq, users, assigns, brands, track, orders, sales, setup: R.users.length > 0 };
  return DB;
}
const ensureDB = async () => DB || run(ctx => loadDB(ctx));

/* ---------- usuarios y permisos ---------- */
function who(tk) {
  const n = SESS[tk]; if (!n) err('SESION: Tu sesión venció, vuelve a ingresar.');
  const u = DB && DB.users.find(x => x.name === n && x.active); if (!u) err('SESION: Tu usuario no está activo.');
  return u;
}
const isAdm = u => u.role === 'admin';
const resvFor = n => DB.assigns.find(a => n >= a.from && n <= a.to);
const inScope = (u, row) => isAdm(u) || !DB.cfg.scopeOwn || str(row.v[C.exec]).trim() === u.name || (DB.track[row.k] || {}).own === u.name || (resvFor(row.r) || {}).user === u.name;
const stateOf = row => ({ k: row.k, r: row.r, v: row.v.slice(), t: JSON.parse(JSON.stringify(DB.track[row.k] || { brands: [], bmeta: {}, hist: [] })) });
const rowOf = k => DB.cart.rows.find(r => r.k === k);
function segArr(k, t) {
  const a = new Array(HDR.seg.length).fill('');
  a[0] = k; a[S_['Revisado']] = t.rev ? 'SI' : ''; a[S_['Resultado']] = t.res; a[S_['Motivo']] = t.motivo; a[S_['Intentos']] = t.int || 0; a[S_['Próximo contacto']] = t.next;
  a[S_['Tipo de negocio']] = t.tipo; a[S_['Gestionado por']] = t.own; a[S_['Primera gestión']] = t.first; a[S_['Última gestión']] = t.last; a[S_['Último contacto']] = t.lastc; a[S_['Depurado']] = t.dep; a[S_['Teléfono corregido']] = t.fix;
  return a;
}
const segObj = a => ({ rev: str(a[S_['Revisado']]) === 'SI', res: str(a[S_['Resultado']]), motivo: str(a[S_['Motivo']]), int: num(a[S_['Intentos']]), next: str(a[S_['Próximo contacto']]), tipo: str(a[S_['Tipo de negocio']]), own: str(a[S_['Gestionado por']]), first: str(a[S_['Primera gestión']]), last: str(a[S_['Última gestión']]), lastc: str(a[S_['Último contacto']]), dep: str(a[S_['Depurado']]), fix: str(a[S_['Teléfono corregido']]) });
async function addHist(ctx, k, user, txt, icon) {
  const d = nowStr(); await appendRows(ctx, 'his', [[k, d, user, txt, icon || 'clock']]);
  const t = DB.track[k]; if (t) t.hist.unshift({ d, who: user, txt, k: icon || 'clock', cm: icon === 'msg' });
}

/* ---------- reportes: libro nuevo, la cartera no recibe hojas ---------- */
function loadXlsx() {
  if (typeof XLSX !== 'undefined') return Promise.resolve();
  return new Promise((ok, ko) => { const s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = ok; s.onerror = () => ko(new Error('No se pudo cargar el generador de Excel. Revisa tu conexión.')); document.head.appendChild(s); });
}
const CREATED = [];

/* ======================================================================
   API: mismas funciones que la versión Sheets
   ====================================================================== */
const API = {
  async api_status() {
    return run(async ctx => {
      try { await loadDB(ctx); } catch (e) { if (e.problems) return { ok: false, problems: e.problems }; throw e; }
      return { ok: true, needSetup: !DB.users.length, sheet: DB.cart.name, rows: DB.cart.rows.length };
    });
  },
  async api_setup(name, pin) {
    name = str(name).trim(); pin = str(pin).trim();
    if (!name) err('Escribe tu nombre.');
    if (!/^\d{4,8}$/.test(pin)) err('El PIN debe tener de 4 a 8 números.');
    return run(async ctx => {
      await loadDB(ctx);
      if (DB.users.length) err('Este libro ya está configurado. Ingresa con tu usuario.');
      for (const k in SH) {
        let ws = await sheet(ctx, SH[k]);
        if (!ws) { ws = ctx.workbook.worksheets.add(SH[k]); writeRow(ws, 0, HDR[k]); }
        ws.visibility = 'VeryHidden';
      }
      await ctx.sync();
      DB.cart.ws.activate();
      await appendRows(ctx, 'cfg', [...Object.entries(CFG_DEF).map(([k, v]) => [k, typeof v === 'boolean' ? (v ? 'SI' : 'NO') : String(v)]), ['seq', '1']]);
      await appendRows(ctx, 'mar', Object.entries(CATALOGO).flatMap(([c, l]) => l.map(m => [m, c, 'SI'])));
      const salt = randId();
      await appendRows(ctx, 'users', [[name, 'admin', 'SI', await sha(salt + ':' + pin), salt, 'NO', '0', '0']]);
      await loadDB(ctx);
      const token = randId(); saveSess(token, name);
      return { token, name };
    });
  },
  async api_login(name, pin) {
    return run(async ctx => {
      await loadDB(ctx);
      const u = DB.users.find(x => norm(x.name) === norm(name) && x.active);
      if (!u) err('Nombre o PIN incorrecto.');
      if (u.lockUntil > Date.now()) err(`Usuario bloqueado por intentos fallidos. Prueba en ${Math.ceil((u.lockUntil - Date.now()) / 60000)} min.`);
      const ok = u.hash && (await sha(u.salt + ':' + str(pin).trim())) === u.hash;
      const tries = ok ? 0 : u.tries + 1, lock = !ok && tries >= 5 ? Date.now() + 5 * 60000 : 0;
      await upsert(ctx, 'users', r => norm(r[0]) === norm(u.name), cur => { const a = cur || []; a[6] = String(ok ? 0 : tries % 5); a[7] = String(lock); return HDR.users.map((_, i) => a[i] == null ? '' : a[i]); });
      if (!ok) err(lock ? 'Demasiados intentos. Usuario bloqueado 5 minutos.' : 'Nombre o PIN incorrecto.');
      const token = randId(); saveSess(token, u.name);
      return { token, name: u.name, mustChange: u.mustChange };
    });
  },
  async api_resume() { const t = Object.keys(SESS)[0]; if (!t) return null; await ensureDB(); const u = DB.users.find(x => x.name === SESS[t] && x.active); return u && !u.mustChange ? { token: t, name: u.name } : null; },
  async api_changePin(tk, pin) {
    const u = who(tk); pin = str(pin).trim();
    if (!/^\d{4,8}$/.test(pin)) err('El PIN debe tener de 4 a 8 números.');
    return run(async ctx => {
      const salt = randId(), h = await sha(salt + ':' + pin);
      await upsert(ctx, 'users', r => norm(r[0]) === norm(u.name), cur => { const a = cur.slice(); a[3] = h; a[4] = salt; a[5] = 'NO'; a[6] = '0'; a[7] = '0'; return a; });
      u.hash = h; u.salt = salt; u.mustChange = false; return true;
    });
  },
  api_logout(tk) { dropSess(tk); return true; },

  async api_load(tk) {
    return run(async ctx => {
      await loadDB(ctx); const u = who(tk);
      const vis = DB.cart.rows.filter(r => inScope(u, r)), ks = new Set(vis.map(r => r.k)), track = {};
      vis.forEach(r => { if (DB.track[r.k]) track[r.k] = DB.track[r.k]; });
      return { me: { name: u.name, role: u.role }, cfg: DB.cfg, scoped: vis.length < DB.cart.rows.length, total: DB.cart.rows.length, header: COLS, C, assigns: DB.assigns,
        users: DB.users.filter(x => x.active).map(x => x.name), seq: DB.seq, lastRow: DB.cart.rows.length + 1, today: TODAY(), sheet: DB.cart.name,
        rows: vis.map(r => [r.k, r.r, r.v]), track, brands: DB.brands, orders: DB.orders.filter(o => ks.has(o.k)), sales: DB.sales.filter(x => ks.has(x.k)), warnings: [] };
    });
  },
  async api_changes(tk, seq) {
    who(tk);
    return run(async ctx => {
      const c = await readAux(ctx, 'cfg'), r = c.rows.find(x => str(x[0]) === 'seq'), cur = r ? num(r[1]) : 0;
      return { assigns: DB.assigns, today: TODAY(), reload: cur !== seq, changes: [], seq: cur };
    });
  },

  async api_action(tk, p) {
    const u = who(tk);
    return run(async ctx => {
      const row = rowOf(p.k); if (!row) err('Cliente no encontrado. Pulsa recargar.');
      if (!inScope(u, row)) err('Este cliente no está en tu cartera.');
      const seg = await readAux(ctx, 'seg'), cur = seg.rows.find(a => str(a[0]) === row.k);
      const t = Object.assign(DB.track[row.k] || (DB.track[row.k] = { buy: false, monto: '', brands: [], bmeta: {}, hist: [] }), cur ? segObj(cur) : {});
      if (!isAdm(u) && p.op !== 'unlock') {
        const a = resvFor(row.r);
        if (a && a.user !== u.name) return { ok: false, error: 'Reservada para ' + a.user, state: stateOf(row) };
        if (DB.cfg.lockRow && t.own && t.own !== u.name) return { ok: false, error: 'Gestionado por ' + t.own, state: stateOf(row) };
      }
      const now = nowStr(), H = [];
      switch (p.op) {
        case 'comment': {
          const txt = str(p.val).trim(); if (!txt) err('El comentario está vacío.');
          /* vuelve a ubicar la fila por Customer ID y relee la celda justo antes de agregar */
          const ids = DB.cart.ws.getRangeByIndexes(1, C.id, Math.max(1, DB.cart.rows.length + 50), 1); ids.load('values'); await ctx.sync();
          const ks = clientKeys(ids.values.map(x => x[0])), i = ks.indexOf(row.k);
          if (i < 0) err('No se encontró al cliente en la hoja. Pulsa recargar.');
          const cell = DB.cart.ws.getRangeByIndexes(i + 1, C.comment, 1, 1); cell.load('values'); await ctx.sync();
          const prev = str(cell.values[0][0]).trim(), line = `${now.slice(8, 10)}/${now.slice(5, 7)} ${u.name}: ${txt}`;
          cell.values = [[prev ? line + ' | ' + prev : line]]; await ctx.sync();
          row.r = i + 2; row.v[C.comment] = prev ? line + ' | ' + prev : line;
          H.push([txt, 'msg']); break;
        }
        case 'rev': t.rev = !!p.val; H.push([p.val ? 'Marcó revisado' : 'Quitó revisado', 'eye']); break;
        case 'result':
          t.res = str(p.val); if (p.val) t.rev = true; t.next = p.val === 'Volver a llamar' ? (p.date || t.next || TODAY()) : '';
          if (p.val !== 'No interesado') t.motivo = '';
          if (p.val === 'No contesta' && (!t.lastc || Date.now() - new Date(t.lastc.replace(' ', 'T')).getTime() > 15 * 60000)) { t.int = (t.int || 0) + 1; t.lastc = now; }
          H.push([p.val ? 'Resultado: ' + p.val : 'Quitó el resultado', ICON[p.val] || 'x']); break;
        case 'next': t.next = str(p.val); H.push(['Reagendó para ' + (p.val || '—'), 'cal']); break;
        case 'motivo': t.motivo = str(p.val); H.push(['Motivo: ' + (p.val || '—'), 'ban']); break;
        case 'tipo': t.tipo = str(p.val); H.push(['Tipo de negocio: ' + (p.val || 'sin definir'), 'store']); break;
        case 'contact': t.int = (t.int || 0) + 1; t.lastc = now; t.rev = true; H.push([p.val ? 'Contactó por ' + p.val : 'Contactó', /WhatsApp/.test(p.val) ? 'wa' : /Correo/.test(p.val) ? 'mail' : 'phone']); break;
        case 'fixphone':
          await appendRows(ctx, 'tel', [[row.k, t.fix || row.v[C.phone], p.val, u.name, now]]);
          H.push(['Corrigió el número: ' + (t.fix || row.v[C.phone]) + ' → ' + p.val, 'pencil']); t.fix = str(p.val); if (t.res === 'Número errado') t.res = ''; break;
        case 'brand': {
          const b = p.val && p.val.b; if (!DB.brands.some(x => x.m === b)) err('La marca no está en el catálogo.');
          const rel = p.val.on ? (p.val.rel || 'habitual') : 'quitada', had = t.brands.includes(b);
          await upsert(ctx, 'cm', r => str(r[0]) === row.k && str(r[1]) === b, () => [row.k, b, rel, u.name, TODAY()]);
          t.brands = t.brands.filter(x => x !== b); t.bmeta = t.bmeta || {};
          if (p.val.on) { t.brands.push(b); t.bmeta[b] = { rel, by: u.name, d: TODAY() }; } else delete t.bmeta[b];
          const RL = { habitual: 'compra habitualmente', ocasional: 'compra a veces', interes: 'le interesa' };
          H.push([p.val.on ? (had ? `Actualizó ${b}: ${RL[rel]}` : `Agregó la marca ${b} (${RL[rel]})`) : 'Quitó la marca ' + b, 'tag']); break;
        }
        case 'unlock': if (!isAdm(u)) err('Solo el administrador puede liberar.'); t.own = ''; H.push(['Liberó el cliente', 'lock']); break;
        default: err('Acción no válida.');
      }
      if (p.op !== 'unlock' && p.op !== 'comment') { if (!t.own) t.own = u.name; if (!t.first) t.first = now; t.last = now; }
      if (p.op === 'comment') { if (!t.own) t.own = u.name; if (!t.first) t.first = now; t.last = now; }
      if (p.op !== 'brand') await upsert(ctx, 'seg', r => str(r[0]) === row.k, () => segArr(row.k, t));
      for (const [txt, ic] of H) await addHist(ctx, row.k, u.name, txt, ic);
      await bumpSeq(ctx);
      return { ok: true, state: stateOf(row) };
    });
  },

  async api_addOrder(tk, p) {
    const u = who(tk);
    return run(async ctx => {
      const row = rowOf(p.k); if (!row) err('Cliente no encontrado.');
      if (!inScope(u, row)) err('Este cliente no está en tu cartera.');
      const ov = str(p.ov).trim().toUpperCase();
      if (!ov) err('El número de OV es obligatorio.');
      if (!/^[A-Z0-9][A-Z0-9\-\/]{2,24}$/.test(ov)) err('La OV solo puede tener letras, números, guiones o barras (3 a 25 caracteres).');
      const ord = await readAux(ctx, 'ord'), dup = ord.rows.find(a => str(a[0]).toUpperCase() === ov);
      if (dup) { const c = rowOf(str(dup[1])); err(`La OV ${ov} ya está registrada para ${c ? c.v[C.legal] : 'otro cliente'}.`); }
      const lines = (p.lines || []).map(l => ({ b: l.b, m: Math.round(num(l.m) * 100) / 100 })).filter(l => l.m > 0);
      if (!lines.length) err('Agrega al menos una marca con su monto sin IGV.');
      lines.forEach(l => { if (!DB.brands.some(x => x.m === l.b)) err('La marca ' + l.b + ' no está en el catálogo.'); });
      const d = /^\d{4}-\d{2}-\d{2}$/.test(p.d || '') ? p.d : TODAY(), tot = lines.reduce((a, l) => a + l.m, 0), now = nowStr();
      const o = { ov, k: row.k, d, tot, est: 'Despachando', by: u.name, fin: '', finBy: '', nota: str(p.nota) };
      await appendRows(ctx, 'ord', [[ov, row.k, d, String(tot), o.est, u.name, '', '', o.nota]]);
      await appendRows(ctx, 'ven', lines.map(l => [ov, row.k, d, l.b, String(l.m), u.name]));
      const ls = lines.map(l => ({ ov, k: row.k, d, b: l.b, m: l.m, by: u.name }));
      DB.orders.push(o); DB.sales.push(...ls);
      const t = DB.track[row.k] || (DB.track[row.k] = { brands: [], bmeta: {}, hist: [] });
      t.buy = true; t.rev = true; t.monto = String(num(t.monto) + tot); if (!t.own) t.own = u.name; if (!t.first) t.first = now; t.last = now;
      await upsert(ctx, 'seg', r => str(r[0]) === row.k, () => segArr(row.k, t));
      for (const l of lines) {
        await upsert(ctx, 'cm', r => str(r[0]) === row.k && str(r[1]) === l.b, () => [row.k, l.b, 'habitual', u.name, TODAY()]);
        if (!t.brands.includes(l.b)) t.brands.push(l.b); t.bmeta = t.bmeta || {}; t.bmeta[l.b] = { rel: 'habitual', by: u.name, d: TODAY() };
      }
      await addHist(ctx, row.k, u.name, `Registró ${ov} · ${lines.map(l => l.b + ' S/ ' + l.m.toLocaleString('es-PE')).join(', ')} · total S/ ${tot.toLocaleString('es-PE')} sin IGV · Despachando`, 'cart');
      await bumpSeq(ctx);
      return { order: o, lines: ls, state: stateOf(row) };
    });
  },
  async api_setOrder(tk, ov, est, nota) {
    const u = who(tk);
    if (!ESTADOS.includes(est)) err('Estado no válido.');
    return run(async ctx => {
      const o = DB.orders.find(x => x.ov === ov); if (!o) err('OV no encontrada.');
      const row = rowOf(o.k); if (!row || !inScope(u, row)) err('Esta OV no es de tu cartera.');
      const ord = await readAux(ctx, 'ord'), a = ord.rows.find(x => str(x[0]) === ov), curEst = a ? str(a[4]) : o.est;
      if ((curEst === 'Terminado' || curEst === 'Anulada') && !isAdm(u)) err('Solo el administrador puede reabrir una OV cerrada.');
      if (est === 'Anulada' && !isAdm(u)) err('Solo el administrador puede anular una OV.');
      o.est = est; if (nota != null) o.nota = str(nota);
      if (est === 'Terminado') { o.fin = TODAY(); o.finBy = u.name; } else { o.fin = ''; o.finBy = ''; }
      await upsert(ctx, 'ord', r => str(r[0]) === ov, cur => { const x = cur || [ov, o.k, o.d, String(o.tot)]; return [x[0], x[1], x[2], x[3], est, x[5] || o.by, o.fin, o.finBy, o.nota]; });
      const t = DB.track[o.k]; if (t) { const live = DB.orders.filter(x => x.k === o.k && x.est !== 'Anulada'); t.monto = String(live.reduce((s, x) => s + x.tot, 0)); t.buy = live.length > 0; }
      await addHist(ctx, o.k, u.name, `${ov}: ${curEst} → ${est}${o.nota ? ' · ' + o.nota : ''}`, est === 'Terminado' ? 'check' : est === 'Incidencia' ? 'alert' : 'cart');
      await bumpSeq(ctx);
      return { order: o, state: stateOf(row) };
    });
  },
  async api_activity(tk) {
    who(tk);
    const days = Array.from({ length: 14 }, (_, i) => { const d = new Date(); d.setDate(d.getDate() + i - 13); return dayStr(d); }), m = {};
    Object.entries(DB.track).forEach(([, t]) => { const seen = {}; t.hist.forEach(h => { const d = h.d.slice(0, 10); if (!days.includes(d) || seen[h.who + d]) return; seen[h.who + d] = 1; (m[h.who] = m[h.who] || {})[d] = ((m[h.who] || {})[d] || 0) + 1; }); });
    return { days, m };
  },
  async api_saveReport(tk, name, aoa) {
    who(tk);
    await loadXlsx();
    let nm = str(name).replace(/[\\\/\?\*\[\]:]/g, '-').slice(0, 28), base = nm, i = 2;
    while (CREATED.some(s => s.name === nm)) nm = base + ' (' + i++ + ')';
    const wb = XLSX.utils.book_new(), ws = XLSX.utils.aoa_to_sheet(aoa.map(r => r.map(str)));
    ws['!cols'] = aoa[0].map((h, j) => ({ wch: Math.min(45, Math.max(String(h).length + 2, ...aoa.slice(1, 200).map(r => str(r[j]).length + 1))) }));
    XLSX.utils.book_append_sheet(wb, ws, nm.slice(0, 31));
    const b64 = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
    if (Excel.createWorkbook) await Excel.createWorkbook(b64);
    else XLSX.writeFile(wb, nm + '.xlsx');
    CREATED.push({ name: nm, rows: aoa.length - 1 });
    return { sheet: nm, sheets: CREATED };
  },
  async api_purge(tk, keys, undo) {
    const u = who(tk); if (!isAdm(u)) err('Solo el administrador puede depurar.');
    return run(async ctx => {
      const seg = await readAux(ctx, 'seg'), ws = await sheet(ctx, SH.seg);
      let next = Math.max(seg.rows.length + 1, 1);
      for (const k of keys) {
        const t = DB.track[k] || (DB.track[k] = { brands: [], bmeta: {}, hist: [] }); t.dep = undo ? '' : TODAY();
        let i = seg.rows.findIndex(a => str(a[0]) === k); const at = i >= 0 ? i + 1 : next++;
        writeRow(ws, at, segArr(k, Object.assign({}, segObj(i >= 0 ? seg.rows[i] : []), { dep: t.dep })));
      }
      await ctx.sync();
      await appendRows(ctx, 'his', keys.map(k => [k, nowStr(), u.name, undo ? 'Restauró de depurados' : 'Marcó como depurado', 'trash']));
      await bumpSeq(ctx);
      return { count: keys.length };
    });
  },
  async api_reserve(tk, p) {
    const u = who(tk), from = Math.max(2, parseInt(p.from, 10)), to = parseInt(p.to, 10), user = isAdm(u) && p.user ? p.user : u.name;
    if (!from || !to || to < from) err('Rango de filas no válido.');
    return run(async ctx => {
      await loadDB(ctx);
      const clash = DB.assigns.find(a => a.user !== user && from <= a.to && to >= a.from);
      if (clash) err(`Se cruza con las filas ${clash.from}–${clash.to} de ${clash.user}.`);
      await appendRows(ctx, 'asg', [[randId().slice(0, 10), user, String(from), String(to), u.name, TODAY(), 'Activa']]);
      await bumpSeq(ctx); await loadDB(ctx);
      return { assigns: DB.assigns, updated: 0, reload: true };
    });
  },
  async api_release(tk, id) {
    const u = who(tk);
    return run(async ctx => {
      const a = DB.assigns.find(x => x.id === id); if (!a) err('Reserva no encontrada.');
      if (!isAdm(u) && a.user !== u.name) err('Solo puedes liberar tus propias filas.');
      await upsert(ctx, 'asg', r => str(r[0]) === id, cur => { const x = cur.slice(); x[6] = 'Liberada'; return x; });
      await bumpSeq(ctx); await loadDB(ctx);
      return DB.assigns;
    });
  },
  async api_users(tk) { const u = who(tk); if (!isAdm(u)) err('Solo el administrador.'); return DB.users.map(x => ({ row: x.row, name: x.name, role: x.role, active: x.active, hasPin: !!x.hash, mustChange: x.mustChange })); },
  async api_saveUser(tk, p) {
    const u = who(tk); if (!isAdm(u)) err('Solo el administrador.');
    const name = str(p.name).trim(); if (!name) err('Escribe un nombre.');
    if (p.pin && !/^\d{4,8}$/.test(p.pin)) err('El PIN temporal debe tener de 4 a 8 números.');
    const ex = p.row ? DB.users.find(x => x.row === p.row) : null;
    if (!ex && !p.pin) err('Asigna un PIN temporal: la persona lo cambiará en su primer ingreso.');
    if (!ex && DB.users.some(x => norm(x.name) === norm(name))) err('Ya existe un usuario con ese nombre.');
    if (ex && ex.name === u.name && (p.role !== 'admin' || !p.active)) err('No puedes quitarte el rol de administrador.');
    return run(async ctx => {
      let hash = ex ? ex.hash : '', salt = ex ? ex.salt : '', must = ex ? (ex.mustChange ? 'SI' : 'NO') : 'SI';
      if (p.pin) { salt = randId(); hash = await sha(salt + ':' + p.pin); must = 'SI'; }
      await upsert(ctx, 'users', r => norm(r[0]) === norm(ex ? ex.name : name), () => [name, p.role === 'admin' ? 'admin' : 'ejecutivo', p.active ? 'SI' : 'NO', hash, salt, must, '0', '0']);
      await bumpSeq(ctx); await loadDB(ctx);
      return API.api_users(tk);
    });
  },
  async api_saveCfg(tk, c) {
    const u = who(tk); if (!isAdm(u)) err('Solo el administrador.');
    const v = Object.assign({}, DB.cfg, c, { mobLen: num(c.mobLen) || 9, minInt: Math.max(1, num(c.minInt) || 1), poll: Math.max(10, num(c.poll) || 20), goal: Math.max(1, num(c.goal) || 40), salesGoal: Math.max(0, num(c.salesGoal)), cc: str(c.cc).replace(/\D/g, '') });
    return run(async ctx => {
      for (const k of Object.keys(CFG_DEF)) await upsert(ctx, 'cfg', r => str(r[0]) === k, () => [k, typeof CFG_DEF[k] === 'boolean' ? (v[k] ? 'SI' : 'NO') : String(v[k])]);
      await bumpSeq(ctx); DB.cfg = v; return v;
    });
  },
  async api_addBrand(tk, b) {
    const u = who(tk); if (!isAdm(u)) err('Solo el administrador.');
    const m = str(b.m).trim(); if (!m) err('Escribe el nombre de la marca.');
    if (DB.brands.some(x => norm(x.m) === norm(m))) err('Esa marca ya está en el catálogo.');
    return run(async ctx => { await appendRows(ctx, 'mar', [[m, b.c || 'Por definir', 'SI']]); await bumpSeq(ctx); DB.brands.push({ m, c: b.c || 'Por definir' }); return DB.brands; });
  },
  /* selección en la hoja: devuelve el Customer ID de la fila marcada (si es la cartera) */
  async api_selection() {
    return run(async ctx => {
      const rg = ctx.workbook.getSelectedRange(); rg.load('rowIndex'); rg.worksheet.load('name'); await ctx.sync();
      if (!DB || rg.worksheet.name !== DB.cart.name || rg.rowIndex < 1) return null;
      const row = DB.cart.rows.find(r => r.r === rg.rowIndex + 1); return row ? row.k : null;
    });
  },
  async api_select(k) {
    const row = DB && rowOf(k); if (!row) return false;
    return run(async ctx => { DB.cart.ws.getRangeByIndexes(row.r - 1, C.id, 1, 1).select(); await ctx.sync(); return true; });
  }
};

/* mismo contrato que google.script.run (la interfaz no cambia).
   Las llamadas se atienden en fila, una a la vez: así una recarga (loadDB) nunca
   se cruza con una escritura a medias del mismo usuario. */
let queue = Promise.resolve();
const runner = (ok, ko) => new Proxy({}, { get(_, fn) {
  if (fn === 'withSuccessHandler') return f => runner(f, ko);
  if (fn === 'withFailureHandler') return f => runner(ok, f);
  return (...args) => {
    if (!API[fn]) { ko && ko(new Error('Función no disponible: ' + fn)); return; }
    const job = queue.then(() => API[fn](...args));
    queue = job.catch(() => {});
    job.then(r => ok && ok(r === undefined ? null : JSON.parse(JSON.stringify(r))), e => ko && ko(e));
  };
} });
window.google = { script: { get run() { return runner(); } } };
window.GC_API = API;
})();
