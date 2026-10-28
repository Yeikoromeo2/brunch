/* Simulador mínimo de Office.js / Excel para probar el complemento en un navegador.
   Implementa solo lo que usa datos.js: hojas, rangos por índice, rango usado,
   visibilidad, selección y Excel.createWorkbook. No es para producción. */
(() => {
'use strict';
const book = { sheets: [], selection: { sheet: null, row: 0, col: 0 }, created: [], handlers: [] };
const S = v => v == null ? '' : v;

function makeSheet(name) {
  const ws = { name, visibility: 'Visible', cells: [], formats: [] };
  book.sheets.push(ws); return ws;
}
const get = (ws, r, c) => (ws.cells[r] && ws.cells[r][c] != null) ? ws.cells[r][c] : '';
const set = (ws, r, c, v) => { (ws.cells[r] = ws.cells[r] || [])[c] = v; };

/* Cada Excel.run tiene su propio contexto, como en Excel real:
   - una propiedad solo se puede leer después de load() y de ctx.sync() de ESE contexto;
   - un objeto de un Excel.run que ya terminó no se puede volver a usar. */
const NOT_LOADED = p => new Error(`La propiedad "${p}" no está disponible. Antes de leer el valor de la propiedad, llame al método de carga en el objeto contenedor y llame a "context.sync()" en el contexto de solicitud asociado.`);
const alive = ctx => { if (!ctx.alive) throw new Error('InvalidObjectPath: el objeto pertenece a otro Excel.run que ya terminó.'); };
function loadable(ctx, obj, props) {
  obj._loaded = new Set();
  obj.load = function (list) { alive(ctx); String(list || props.join(',')).split(',').map(x => x.trim()).filter(Boolean).forEach(x => ctx.pending.push([obj, x])); return obj; };
  const need = p => { if (!obj._loaded.has(p)) throw NOT_LOADED(p); };
  return need;
}
function range(ws, r0, c0, nr, nc, ctx) {
  const o = { isNullObject: false, worksheet: api(ws, ctx) };
  const need = loadable(ctx, o, ['values']);
  const vals = () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(ws, r0 + i, c0 + j)));
  Object.defineProperties(o, {
    rowIndex: { get() { need('rowIndex'); return r0; } },
    columnIndex: { get() { need('columnIndex'); return c0; } },
    rowCount: { get() { need('rowCount'); return nr; } },
    columnCount: { get() { need('columnCount'); return nc; } },
    values: { get() { need('values'); return vals(); }, set(v) { alive(ctx); v.forEach((row, i) => row.forEach((x, j) => set(ws, r0 + i, c0 + j, x))); } },
    text: { get() { need('text'); return vals().map(row => row.map(x => typeof x === 'number' ? String(x) : String(S(x)))); } },
    numberFormat: { set(f) { alive(ctx); f.forEach((row, i) => row.forEach((x, j) => { (ws.formats[r0 + i] = ws.formats[r0 + i] || [])[c0 + j] = x; })); } }
  });
  o.select = () => { alive(ctx); book.selection = { sheet: ws.name, row: r0, col: c0 }; };
  return o;
}
function used(ws, ctx) {
  let maxR = -1, maxC = -1;
  ws.cells.forEach((row, r) => row && row.forEach((x, c) => { if (x !== '' && x != null) { if (r > maxR) maxR = r; if (c > maxC) maxC = c; } }));
  if (maxR < 0) return { isNullObject: true, load() { return this; } };
  return range(ws, 0, 0, maxR + 1, maxC + 1, ctx);
}
function api(ws, ctx) {
  return {
    get name() { return ws.name; },
    get visibility() { return ws.visibility; },
    set visibility(v) { alive(ctx); ws.visibility = v; },
    isNullObject: false,
    load() { return this; },
    getUsedRangeOrNullObject() { alive(ctx); return used(ws, ctx); },
    getRangeByIndexes(r, c, nr, nc) { alive(ctx); return range(ws, r, c, nr, nc, ctx); },
    activate() { alive(ctx); book.active = ws.name; }
  };
}
function makeWorkbook(ctx) {
  return {
    worksheets: {
      load() { return this; },
      get items() { return book.sheets.map(w => api(w, ctx)); },
      getItem(n) { alive(ctx); const ws = book.sheets.find(x => x.name === n); if (!ws) throw new Error('ItemNotFound: ' + n); return api(ws, ctx); },
      getItemOrNullObject(n) { alive(ctx); const ws = book.sheets.find(x => x.name === n); return ws ? api(ws, ctx) : { isNullObject: true }; },
      add(n) { alive(ctx); if (book.sheets.some(x => x.name === n)) throw new Error('Ya existe la hoja ' + n); return api(makeSheet(n), ctx); }
    },
    getSelectedRange() {
      alive(ctx);
      const ws = book.sheets.find(x => x.name === book.selection.sheet) || book.sheets[0];
      return range(ws, book.selection.row, book.selection.col, 1, 1, ctx);
    }
  };
}
window.Excel = {
  run: async fn => {
    book.runs = (book.runs || 0) + 1;
    const ctx = { alive: true, pending: [] };
    ctx.sync = async () => { alive(ctx); ctx.pending.splice(0).forEach(([o, p]) => o._loaded.add(p)); };
    ctx.workbook = makeWorkbook(ctx);
    try { return await fn(ctx); } finally { ctx.alive = false; book.runs--; }
  },
  createWorkbook: async b64 => { book.created.push(b64); }
};
/* ventana de diálogo: el panel la abre con window.open y se hablan por mensajes de texto */
const EV = { DocumentSelectionChanged: 'documentSelectionChanged', DialogMessageReceived: 'dialogMessageReceived', DialogEventReceived: 'dialogEventReceived', DialogParentMessageReceived: 'dialogParentMessageReceived' };
const ui = {
  displayDialogAsync(url, opts, cb) {
    const w = window.open(url, '_blank');
    if (!w) return cb({ status: 'failed', error: { message: 'Ventana bloqueada' } });
    const on = { [EV.DialogMessageReceived]: [], [EV.DialogEventReceived]: [] }, queue = [];
    window.__fromChild = m => on[EV.DialogMessageReceived].forEach(h => h({ message: m }));
    const timer = setInterval(() => { if (w.closed) { clearInterval(timer); on[EV.DialogEventReceived].forEach(h => h({ error: 12006 })); } }, 200);
    book.dialog = w;
    cb({ status: 'succeeded', value: {
      addEventHandler(t, h) { on[t].push(h); },
      messageChild(m) { if (w.__toChild) w.__toChild(m); else queue.push(m); },
      close() { clearInterval(timer); w.close(); }
    } });
    window.__dlgReady = () => { while (queue.length) w.__toChild(queue.shift()); };
  }
};
const isChild = !!(window.opener && window.opener.__fromChild);
if (isChild) {
  /* en la ventana no hay acceso al libro */
  delete window.Excel;
  ui.messageParent = m => window.opener.__fromChild(m);
  ui.addHandlerAsync = (t, h, cb) => { window.__toChild = m => h({ message: m }); if (window.opener.__dlgReady) window.opener.__dlgReady(); cb && cb({ status: 'succeeded' }); };
}
window.Office = {
  HostType: { Excel: 'Excel' },
  EventType: EV,
  AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
  onReady(cb) { setTimeout(() => cb({ host: 'Excel' }), 0); return Promise.resolve({ host: 'Excel' }); },
  context: { document: { addHandlerAsync(t, h) { book.handlers.push(h); } }, ui, requirements: { isSetSupported: () => true } }
};
/* libro inicial que la prueba deja en window.__seed = [{ name, rows }] antes de cargar la página */
(window.__seed || []).forEach(s => { const ws = makeSheet(s.name); s.rows.forEach((r, i) => r.forEach((x, j) => set(ws, i, j, x))); });
/* ayudas para las pruebas */
window.__book = {
  raw: book,
  addSheet(name, rows) { const ws = makeSheet(name); rows.forEach((r, i) => r.forEach((x, j) => set(ws, i, j, x))); return ws; },
  sheet(name) { return book.sheets.find(s => s.name === name); },
  dump(name) { const ws = book.sheets.find(s => s.name === name); return ws ? ws.cells.map(r => (r || []).map(S)) : null; },
  selectRow(name, r) { book.selection = { sheet: name, row: r, col: 0 }; book.handlers.forEach(h => h({})); }
};
})();
