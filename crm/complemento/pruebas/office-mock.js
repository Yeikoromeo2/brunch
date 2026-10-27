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

function range(ws, r0, c0, nr, nc) {
  return {
    rowIndex: r0, columnIndex: c0, rowCount: nr, columnCount: nc, isNullObject: false,
    worksheet: api(ws),
    load() { return this; },
    get values() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(ws, r0 + i, c0 + j))); },
    set values(v) { v.forEach((row, i) => row.forEach((x, j) => set(ws, r0 + i, c0 + j, x))); },
    get text() { return this.values.map(row => row.map(x => typeof x === 'number' ? String(x) : String(S(x)))); },
    set numberFormat(f) { f.forEach((row, i) => row.forEach((x, j) => { (ws.formats[r0 + i] = ws.formats[r0 + i] || [])[c0 + j] = x; })); },
    select() { book.selection = { sheet: ws.name, row: r0, col: c0 }; }
  };
}
function used(ws) {
  let maxR = -1, maxC = -1;
  ws.cells.forEach((row, r) => row && row.forEach((x, c) => { if (x !== '' && x != null) { if (r > maxR) maxR = r; if (c > maxC) maxC = c; } }));
  if (maxR < 0) return { isNullObject: true, load() { return this; } };
  return Object.assign(range(ws, 0, 0, maxR + 1, maxC + 1), {});
}
function api(ws) {
  return {
    get name() { return ws.name; },
    get visibility() { return ws.visibility; },
    set visibility(v) { ws.visibility = v; },
    isNullObject: false,
    load() { return this; },
    getUsedRangeOrNullObject() { return used(ws); },
    getRangeByIndexes(r, c, nr, nc) { return range(ws, r, c, nr, nc); },
    activate() { book.active = ws.name; }
  };
}
const workbook = {
  worksheets: {
    load() { return this; },
    get items() { return book.sheets.map(api); },
    getItemOrNullObject(n) { const ws = book.sheets.find(s => s.name === n); return ws ? api(ws) : { isNullObject: true }; },
    add(n) { if (book.sheets.some(s => s.name === n)) throw new Error('Ya existe la hoja ' + n); return api(makeSheet(n)); }
  },
  getSelectedRange() {
    const ws = book.sheets.find(s => s.name === book.selection.sheet) || book.sheets[0];
    return range(ws, book.selection.row, book.selection.col, 1, 1);
  }
};
window.Excel = {
  run: async fn => { book.runs = (book.runs || 0) + 1; try { return await fn({ workbook, sync: async () => {} }); } finally { book.runs--; } },
  createWorkbook: async b64 => { book.created.push(b64); }
};
window.Office = {
  HostType: { Excel: 'Excel' },
  EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
  onReady(cb) { setTimeout(() => cb({ host: 'Excel' }), 0); return Promise.resolve({ host: 'Excel' }); },
  context: { document: { addHandlerAsync(t, h) { book.handlers.push(h); } }, ui: {} }
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
