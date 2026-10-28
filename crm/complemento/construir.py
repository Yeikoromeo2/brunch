"""Genera complemento/taskpane.html y complemento/ventana.html a partir del prototipo aprobado.

Uso:  python3 crm/complemento/construir.py

Toma la interfaz de crm/prototipo-pro/index.html, quita el servidor simulado,
la conecta a datos.js (Office.js) y agrega las pantallas reales: configuración
inicial, cambio de PIN y error de estructura. ventana.html es la misma interfaz
para la ventana aparte: no toca el libro, le pide todo al panel por puente.js. Cada reemplazo debe encontrarse
exactamente una vez; si el prototipo cambia y algo no calza, el script falla.
"""
import re
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
SRC = AQUI.parent / "prototipo-pro" / "index.html"
OUT = AQUI / "taskpane.html"
OUT_WIN = AQUI / "ventana.html"

html = SRC.read_text(encoding="utf-8")


def rep(old, new, count=1):
    global html
    n = html.count(old)
    if n != count:
        sys.exit(f"Reemplazo esperado {count} vez/veces, encontrado {n}: {old[:80]!r}")
    html = html.replace(old, new)


def rep_re(pat, new, flags=re.S):
    global html
    html, n = re.subn(pat, new, html, count=1, flags=flags)
    if n != 1:
        sys.exit(f"Patrón no encontrado: {pat[:80]!r}")


# ---------- documento completo ----------
rep("<title>Gestor de Cartera Pro</title>",
    '<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
    '<title>Gestor de Cartera</title>\n'
    '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>')
rep('<svg style="display:none" xmlns="http://www.w3.org/2000/svg">',
    '</head>\n<body>\n<svg style="display:none" xmlns="http://www.w3.org/2000/svg">')

# ---------- estilos propios del complemento ----------
CSS = """
/* complemento: pantallas de inicio, panel angosto de Excel */
.cover .steps{margin:0;padding-left:18px;color:var(--ink2);font-size:13px;display:flex;flex-direction:column;gap:4px}
.cover .probs{margin:0;padding:10px 12px 10px 28px;background:color-mix(in srgb,var(--crit) 8%,var(--surface));border:1px solid color-mix(in srgb,var(--crit) 30%,var(--line));border-radius:9px;font-size:13px;display:flex;flex-direction:column;gap:4px}
.cover .ok-note{font-size:12px;color:var(--muted)}
.p-nav{display:none;align-items:center;gap:6px;padding:8px 12px;border-bottom:1px solid var(--line);background:var(--surface);position:sticky;top:0;z-index:2}
.p-nav .spacer{flex:1}
.p-nav .btn{height:30px}
#bFilt{display:none}
@media (max-width:1000px){.p-nav{display:flex}}
@media (max-width:760px){
  header.top{padding:6px 8px;gap:4px}
  header.top .btn{padding:0 7px}
  .tabs button{padding:0 9px}
  #bFilt{display:inline-flex}
  .toolbar .field:not(.search),.toolbar #bClear,.toolbar .ck{display:none}
  .toolbar.open .field,.toolbar.open #bClear,.toolbar.open .ck{display:flex}
  .toolbar .field.search{flex:1 1 100%}
  .chips{flex-wrap:nowrap;overflow-x:auto;scrollbar-width:none;max-width:100%}
  .chips .chip{flex:none}
  .strip{gap:10px;padding:6px 10px}
  #panel{width:100%}
}
/* lista angosta: fila, gestión y razón social primero; acciones y comentario en el panel */
#grid.narrow .s1,#grid.narrow .sr{display:none}
</style>"""
rep("@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}\n</style>",
    "@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}" + CSS)

# ---------- pantallas de acceso ----------
rep_re(r'    <div class="demo"><b>Prototipo · usuarios de prueba</b>.*?</div>\n', "")
rep('<p>Ingresa con tu nombre y tu PIN.</p>',
    '<p>Ingresa con tu nombre y tu PIN.</p><p class="ok-note" id="lgSheet"></p>')
rep('<div class="cover" id="login">', '<div class="cover" id="login" hidden>')
COVERS = """
<!-- PRIMER USO: configura el libro y crea al administrador -->
<div class="cover" id="setup" hidden>
  <form class="login" id="setupForm" autocomplete="off">
    <div class="logo"><svg><use href="#i-list"/></svg></div>
    <div><h1>Configurar este libro</h1><p>Es la primera vez que se abre el Gestor en este archivo.</p></div>
    <ol class="steps">
      <li>Se validan las 18 columnas de la cartera (<b id="suSheet"></b>, <b id="suRows"></b> clientes).</li>
      <li>Se crean las pestañas internas «muy ocultas» del sistema.</li>
      <li>Tú quedas como administrador y luego creas a tu equipo.</li>
    </ol>
    <label class="fl"><span>Tu nombre</span><input type="text" id="suName" required></label>
    <div class="two">
      <label class="fl"><span>PIN (4 a 8 números)</span><input type="password" id="suPin" inputmode="numeric" required></label>
      <label class="fl"><span>Repite el PIN</span><input type="password" id="suPin2" inputmode="numeric" required></label>
    </div>
    <div class="err" id="suErr"></div>
    <button class="btn primary" id="suBtn" type="submit" style="height:40px;justify-content:center">Configurar y entrar</button>
    <p class="ok-note">La hoja de la cartera no se modifica: solo se agregarán comentarios en «Comentario Créditos».</p>
  </form>
</div>

<!-- CAMBIO DE PIN (primer ingreso con PIN temporal) -->
<div class="cover" id="pinCover" hidden>
  <form class="login" id="pinForm" autocomplete="off">
    <div class="logo"><svg><use href="#i-lock"/></svg></div>
    <div><h1>Crea tu PIN</h1><p>Entraste con un PIN temporal. Elige uno nuevo que solo tú conozcas.</p></div>
    <div class="two">
      <label class="fl"><span>PIN nuevo</span><input type="password" id="pnPin" inputmode="numeric" required></label>
      <label class="fl"><span>Repítelo</span><input type="password" id="pnPin2" inputmode="numeric" required></label>
    </div>
    <div class="err" id="pnErr"></div>
    <button class="btn primary" id="pnBtn" type="submit" style="height:40px;justify-content:center">Guardar y entrar</button>
  </form>
</div>

<!-- LA ESTRUCTURA NO COINCIDE -->
<div class="cover" id="probCover" hidden>
  <div class="login">
    <div class="logo" style="background:var(--crit)"><svg><use href="#i-alert"/></svg></div>
    <div><h1>No se puede usar este archivo</h1><p>La cartera debe tener las 18 columnas de la plantilla, con los mismos nombres y en el mismo orden. No se modificó nada.</p></div>
    <ul class="probs" id="probList"></ul>
    <button class="btn primary" id="probRetry" type="button" style="height:40px;justify-content:center">Volver a revisar</button>
  </div>
</div>

<!-- EL GESTOR ESTÁ ABIERTO EN LA VENTANA APARTE (solo en el panel de Excel) -->
<div class="cover" id="winCover" hidden>
  <div class="login">
    <div class="logo"><svg><use href="#i-ext"/></svg></div>
    <div><h1>Abierto en otra ventana</h1><p>El Gestor está trabajando en la ventana aparte. Deja este panel abierto: es el que guarda los cambios en el libro.</p></div>
    <button class="btn primary" id="winClose" type="button" style="height:40px;justify-content:center">Cerrar la ventana y volver aquí</button>
    <p class="ok-note">Si cierras este panel, la ventana deja de guardar.</p>
  </div>
</div>

<!-- LA VENTANA PERDIÓ AL PANEL DE EXCEL (solo en la ventana aparte) -->
<div class="cover" id="lostCover" hidden>
  <div class="login">
    <div class="logo" style="background:var(--crit)"><svg><use href="#i-alert"/></svg></div>
    <div><h1>Sin conexión con Excel</h1><p>Excel volvió a cargar el panel del Gestor (lo hace por su cuenta a veces) o el panel se cerró. Esta ventana ya no puede guardar en el libro; lo que ya guardaste no se pierde.</p></div>
    <ol class="steps"><li>Cierra esta ventana.</li><li>En el panel del Gestor en Excel pulsa <b>Volver a abrir la ventana grande</b> (o <b>Ampliar</b>).</li></ol>
    <button class="btn primary" id="lostClose" type="button" style="height:40px;justify-content:center">Cerrar esta ventana</button>
  </div>
</div>

<!-- EXCEL RECARGÓ EL PANEL MIENTRAS LA VENTANA ESTABA ABIERTA (solo en el panel de Excel) -->
<div class="cover" id="winLost" hidden>
  <div class="login">
    <div class="logo"><svg><use href="#i-ext"/></svg></div>
    <div><h1>Excel recargó el panel</h1><p>Estabas trabajando en la ventana grande. Excel volvió a cargar este panel por su cuenta, así que la ventana se cerró. Todo lo que guardaste está en el libro.</p></div>
    <button class="btn primary" id="winReopen" type="button" style="height:40px;justify-content:center">Volver a abrir la ventana grande</button>
    <button class="btn ghost" id="winStay" type="button" style="justify-content:center">Seguir en este panel</button>
  </div>
</div>
<div class="cover" id="loading">"""
rep('<div class="cover" id="loading" hidden>', COVERS)
rep('<div id="loadingTx">Cargando cartera…</div>', '<div id="loadingTx">Abriendo el libro…</div>')

# ---------- barra superior ----------
rep('Gestor de Cartera<small>Datos de ejemplo</small>', 'Gestor de Cartera<small id="brandSheet">Excel · OneDrive</small>')
rep('<button class="btn ghost" id="bReload" title="Recargar datos">',
    '<button class="btn" id="bWin" title="Abrir el Gestor en una ventana aparte, más grande"><svg><use href="#i-ext"/></svg>Ampliar</button>\n  <button class="btn ghost" id="bReload" title="Recargar datos">')
rep('<svg style="display:none" xmlns="http://www.w3.org/2000/svg">',
    '<svg style="display:none" xmlns="http://www.w3.org/2000/svg">\n  <symbol id="i-ext" viewBox="0 0 24 24"><path fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" d="M14 3h7v7M21 3l-9 9M19 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h5"/></symbol>')
rep('<div class="field search"><svg><use href="#i-search"/></svg>',
    '<button class="btn ghost" id="bFilt" type="button" title="Mostrar filtros"><svg><use href="#i-flag"/></svg>Filtros</button>\n  <div class="field search"><svg><use href="#i-search"/></svg>')

# ---------- usuarios: PIN temporal ----------
rep('<label class="fl"><span>PIN (4 a 8 números)</span><input type="text" id="uPin"',
    '<label class="fl"><span>PIN temporal (4 a 8 números)</span><input type="text" id="uPin"')
rep('<div class="mf"><button class="btn ghost" id="uNew">Nuevo</button>',
    '<div class="hint" style="margin:0 20px 10px">La persona entra con el PIN temporal y el sistema le pide crear el suyo. Si olvida su PIN, escribe uno temporal nuevo aquí.</div>\n  <div class="mf"><button class="btn ghost" id="uNew">Nuevo</button>')

# ---------- quitar el servidor simulado ----------
rep_re(r"<script>\n/\* =+\n   SERVIDOR SIMULADO.*?</script>\n", '<!--CAPA-DATOS-->\n')

# ---------- aplicación ----------
rep("const DEMO = true;", "const DEMO = false;")
rep("$$('[data-demo]').forEach(b => b.onclick = () => { const [n, p] = b.dataset.demo.split('|'); $('#lgName').value = n; $('#lgPin').value = p; $('#loginForm').requestSubmit(); });\n", "")
rep("const money = n => 'S/ ' + Math.round(Number(n) || 0).toLocaleString('es-PE');",
    "const money = n => 'S/ ' + Math.round(Number(String(n == null ? '' : n).replace(/[^\\d.-]/g, '')) || 0).toLocaleString('es-PE');")
rep("const d = daysSince(val(r, 'last')), cr = +val(r, 'credit') || 0,",
    "const d = daysSince(val(r, 'last')), cr = +val(r, 'credit').replace(/[^\\d.]/g, '') || 0,")

# ventana aparte: si se pierde el panel de Excel, se bloquea con un aviso claro
rep("""      if (/SESION/.test(msg)) sessionLost(msg.replace(/^.*SESION:\\s*/, ''));""",
    """      if (window.GC_BRIDGE && /conexión con el panel/.test(msg)) { stopPoll(); showCover('#lostCover'); }
      if (/SESION/.test(msg)) sessionLost(msg.replace(/^.*SESION:\\s*/, ''));""")

# arranque rápido: copia de la última carga en el navegador, recarga en segundo plano
rep("""async function loadAll(first) {
  $('#loading').hidden = false; $('#loadingTx').textContent = 'Cargando cartera…';""",
    """async function loadAll(first, quiet, pre) {
  if (!quiet) { $('#loading').hidden = false; $('#loadingTx').textContent = 'Cargando cartera…'; }""")
rep("""    const d = await call('api_load', S.token);
    S.me = d.me;""",
    """    const d = pre || await call('api_load', S.token);
    if (!pre) CACHE.set({ name: d.me.name, token: S.token, d });
    S.me = d.me;""")
rep("""    setTab(first ? 'day' : S.tab); updBell();""",
    """    setTab(first ? (SS.get('tab') || 'day') : S.tab); updBell();""")
rep("""    setSync('', 'Sincronizado');
    startPoll();
    if (first) toast(""",
    """    setSync(pre ? 'saving' : '', pre ? 'Actualizando…' : 'Sincronizado');
    startPoll();
    if (first && !pre) toast(""")
rep("""function setTab(t) {
  S.tab = t;""", """function setTab(t) {
  S.tab = t; SS.set('tab', t);""")
rep("""$('#bOut').onclick = () => { call('api_logout', S.token).catch(() => {}); sessionLost(' '); };""",
    """$('#bOut').onclick = () => { CACHE.clear(); call('api_logout', S.token).catch(() => {}); sessionLost(' '); };""")

# login: PIN temporal → pantalla de cambio
rep("""    S.token = res.token; LS.set('lastName', res.name);
    $('#lgPin').value = '';
    await loadAll(true);""",
    """    S.token = res.token; LS.set('lastName', res.name);
    $('#lgPin').value = '';
    if (res.mustChange) showCover('#pinCover'); else await loadAll(true);""")
rep("""function showLogin(msg) {
  $('#login').hidden = false; $('#loading').hidden = true;""",
    """function showCover(id) {
  ['#login', '#setup', '#pinCover', '#probCover', '#winCover', '#lostCover', '#winLost'].forEach(s => $(s).hidden = s !== id);
  $('#loading').hidden = true;
  const f = $(id).querySelector('input'); if (f) setTimeout(() => { const a = document.activeElement; if (!a || a === document.body) f.focus(); }, 50);
}
function showLogin(msg) {
  showCover('#login');""")
rep("    $('#login').hidden = true;\n    $('#meName')", "    ['#login', '#setup', '#pinCover', '#probCover', '#winCover'].forEach(s => $(s).hidden = true);\n    $('#brandSheet').textContent = 'Excel · ' + d.sheet;\n    $('#meName')")

# abrir enlaces: ventana real, o el navegador del sistema en Excel de escritorio
rep("""  if (DEMO) return toast('En la versión real se abre ' + label + '. Se registró el intento.');
  const w = window.open(u, '_blank'); if (!w) toast('Tu navegador bloqueó la ventana. Permite ventanas emergentes para este sitio.', true);""",
    """  if (DEMO) return toast('En la versión real se abre ' + label + '. Se registró el intento.');
  let w = null; try { w = window.open(u, '_blank'); } catch (e) {}
  if (w) return;
  try { if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) return Office.context.ui.openBrowserWindow(u); } catch (e) {}
  toast('Tu navegador bloqueó la ventana. Permite ventanas emergentes para este sitio.', true);""")

# reportes: libro nuevo real, sin botón de descarga simulada
rep("""        <button class="btn" id="rXlsx" ${rows.length ? '' : 'disabled'}>${icon('down')}Descargar Excel</button>\n""", "")
rep("""  if (id === 'rXlsx') { if (DEMO) toast(`En la versión real se descarga «Reporte_${name.replace(/\\s+/g, '_')}_${S.today}.xlsx» con ${fmtN(rows.length)} filas.`); }
  else if (id === 'rCopy')""", "  if (id === 'rCopy')")
rep("S.sheets = res.sheets; toast('En la versión real se abre el libro nuevo «' + res.sheet + '.xlsx» en tu OneDrive. La cartera original no se toca.'); }",
    "S.sheets = res.sheets; toast('Libro nuevo «' + res.sheet + '» creado: Excel lo abre en otra pestaña. La cartera no se toca.'); }")
rep("S.sheets = res.sheets; toast(`En la versión real se abre el libro nuevo «${res.sheet}.xlsx» con ${fmtN(rows.length)} clientes, las 18 columnas y sus marcas.`); }",
    "S.sheets = res.sheets; toast(`Libro nuevo «${res.sheet}» con ${fmtN(rows.length)} clientes, las 18 columnas y sus marcas. Excel lo abre en otra pestaña.`); }")
rep("Archivos creados en tu OneDrive", "Libros creados en esta sesión")
rep("Las de «Por definir» se pueden reasignar en la versión real.", "Para cambiar la línea de una marca, pídeselo al administrador del libro.")

# panel angosto: volver a la lista y anterior / siguiente
rep("""  p.innerHTML = `<div class="p-head" style="--tc:var(--t${tr})">""",
    """  p.innerHTML = `<div class="p-nav"><button class="btn" data-act="navback">${icon('list')}Lista</button><span class="spacer"></span><span style="font-size:12px;color:var(--muted)">${r._vi != null ? fmtN(r._vi + 1) + ' de ' + fmtN(S.view.length) : ''}</span><button class="btn" data-act="navprev" title="Anterior">‹</button><button class="btn" data-act="navnext" title="Siguiente">›</button></div><div class="p-head" style="--tc:var(--t${tr})">""")
rep("""$('#panel').addEventListener('click', e => { const b = e.target.closest('[data-act]'); if (!b || !S.active || b.disabled) return; e.preventDefault(); act(S.active, b.dataset.act, actArg(b)); });""",
    """$('#panel').addEventListener('click', e => {
  const b = e.target.closest('[data-act]'); if (!b || !S.active || b.disabled) return; e.preventDefault();
  const a = b.dataset.act;
  if (a === 'navback') { UI.panelM = false; saveUI(); return applyPanel(); }
  if (a === 'navprev' || a === 'navnext') return setActive(S.active._vi + (a === 'navprev' ? -1 : 1), true);
  act(S.active, a, actArg(b));
});""")
rep("""  if (b) { e.preventDefault(); if (!b.disabled) act(r, b.dataset.act, actArg(b)); }
  else if (!e.target.closest('.cmt')) $('#grid').focus({ preventScroll: true });""",
    """  if (b) { e.preventDefault(); if (!b.disabled) act(r, b.dataset.act, actArg(b)); }
  else if (!e.target.closest('.cmt')) { $('#grid').focus({ preventScroll: true }); if (innerWidth <= 1000 && !UI.panelM) { UI.panelM = true; saveUI(); applyPanel(); } }""")
rep("$('#bReload').onclick = () => { flushAll(); loadAll(false); };",
    "$('#bReload').onclick = () => { flushAll(); loadAll(false); };\n$('#bFilt').onclick = () => $('#toolbar').classList.toggle('open');")

# lista angosta (panel de Excel): columnas compactas
rep("""  const FIRST = ['id', 'legal', 'tax', 'phone', 'mail', 'state', 'city', 'exec'];""",
    """  const narrow = innerWidth <= 760; NARROW = narrow;
  const FIRST = narrow ? ['legal', 'id', 'phone', 'tax', 'mail', 'state', 'city', 'exec'] : ['id', 'legal', 'tax', 'phone', 'mail', 'state', 'city', 'exec'];""")
rep("""  const w = [78, 176, 200, ...COLS.map(c => WIDTH[c.k] || 130), 300];
  const g = $('#grid');""",
    """  const w = narrow ? [70, 124, ...COLS.map(c => c.k === 'legal' ? 210 : WIDTH[c.k] || 130)] : [78, 176, 200, ...COLS.map(c => WIDTH[c.k] || 130), 300];
  const g = $('#grid'); g.classList.toggle('narrow', narrow);""")
rep("window.addEventListener('resize', () => { render(); applyPanel(); moveInd(); });",
    "let NARROW = null;\nwindow.addEventListener('resize', () => { if (S.me && (innerWidth <= 760) !== NARROW) buildCols(); render(); applyPanel(); moveInd(); });")
rep("  if (prev && prev !== S.active) { if (S.cmtT[prev.k]) flushComment(prev);",
    "  if (prev && prev !== S.active) { $('#panel').scrollTop = 0; if (S.cmtT[prev.k]) flushComment(prev);")

# la hoja sigue al cliente elegido en el panel
rep("  if (prev !== S.active) peekSoon(S.active.k);\n}",
    "  if (prev !== S.active) peekSoon(S.active.k);\n  followSheet(S.active.k);\n}")

# ---------- inicio real ----------
BOOT = r"""/* ---------- vínculo con la hoja: seleccionar una fila abre el cliente y viceversa ---------- */
let selT = null, selQuiet = 0, selLast = '';
function followSheet(k) {
  /* desde la ventana grande no se mueve la selección de la hoja: Excel traería su página al frente
     y taparía la ventana (con dos pantallas se nota mucho) */
  if (window.GC_BRIDGE || !UI.follow || !S.me || k === selLast) return;
  clearTimeout(selT);
  selT = setTimeout(() => {
    const a = document.activeElement; if (a && /INPUT|TEXTAREA|SELECT/.test(a.tagName)) return;
    selLast = k; selQuiet = Date.now();
    call('api_select', k).catch(() => {});
  }, 600);
}
function onSheetSelection() {
  if (window.GC_HOST && GC_HOST.open) return GC_HOST.selection();
  if (!S.me) return;
  call('api_selection').then(k => {
    /* se ignora el eco de nuestra propia selección, nunca una fila distinta elegida por la persona */
    if (!k || (S.active && S.active.k === k) || (k === selLast && Date.now() - selQuiet < 1500)) return;
    const r = S.byKey.get(k);
    if (!r) return toast('Esa fila no está en tu cartera.', true);
    selLast = k; openClient(k);
    if (innerWidth <= 1000 && !UI.panelM) { UI.panelM = true; saveUI(); applyPanel(); }
  }).catch(() => {});
}
if (UI.follow == null) { UI.follow = true; saveUI(); }

/* ---------- pantallas de inicio ---------- */
$('#setupForm').addEventListener('submit', async e => {
  e.preventDefault();
  const pin = $('#suPin').value.trim();
  if (pin !== $('#suPin2').value.trim()) { $('#suErr').textContent = 'Los PIN no coinciden.'; return; }
  $('#suErr').textContent = ''; $('#suBtn').disabled = true; $('#suBtn').textContent = 'Configurando…';
  try {
    const res = await call('api_setup', $('#suName').value, pin);
    S.token = res.token; LS.set('lastName', res.name);
    await loadAll(true);
    toast('Libro configurado. Crea a tu equipo desde el botón de usuarios.');
  } catch (er) { $('#suErr').textContent = er.message; }
  $('#suBtn').disabled = false; $('#suBtn').textContent = 'Configurar y entrar';
});
$('#pinForm').addEventListener('submit', async e => {
  e.preventDefault();
  const pin = $('#pnPin').value.trim();
  if (pin !== $('#pnPin2').value.trim()) { $('#pnErr').textContent = 'Los PIN no coinciden.'; return; }
  $('#pnErr').textContent = ''; $('#pnBtn').disabled = true;
  try { await call('api_changePin', S.token, pin); $('#pnPin').value = $('#pnPin2').value = ''; toast('PIN actualizado'); await loadAll(true); }
  catch (er) { $('#pnErr').textContent = er.message; }
  $('#pnBtn').disabled = false;
});
$('#probRetry').onclick = () => boot();

/* ---------- ventana aparte: el panel queda de puente y guarda en el libro ---------- */
if (window.GC_BRIDGE) $('#bWin').hidden = true;
$('#bWin').onclick = async () => {
  if (!window.GC_HOST || !GC_HOST.supported()) return toast('Tu versión de Excel no permite abrir ventanas aparte.', true);
  flushAll();
  try {
    await GC_HOST.openWindow(location.href.replace(/[?#].*$/, '').replace(/[^/]*$/, '') + 'ventana.html', () => { SS.set('win', ''); $('#winCover').hidden = true; boot(); });
    SS.set('win', '1'); stopPoll(); showCover('#winCover');
  } catch (e) { toast('No se pudo abrir la ventana: ' + e.message + '. Si Excel lo pregunta, permite las ventanas emergentes.', true); }
};
$('#winClose').onclick = () => GC_HOST.close();
/* si Excel recarga el panel con la ventana abierta: se cierra la ventana antes de irse (quedaría sin
   conexión) y, al volver, se ofrece abrirla de nuevo con un clic */
window.addEventListener('pagehide', () => { if (window.GC_HOST && GC_HOST.open) GC_HOST.drop(); });
$('#winReopen').onclick = () => { $('#winLost').hidden = true; $('#bWin').click(); };
$('#winStay').onclick = () => { SS.set('win', ''); $('#winLost').hidden = true; };
$('#lostClose').onclick = () => window.close();

/* copia local de la última carga (IndexedDB de este navegador). Si Excel recarga el panel,
   la pantalla aparece al instante y se pone al día en segundo plano. Se borra al salir. */
const CACHE = (() => {
  const open = () => new Promise((ok, ko) => { const r = indexedDB.open('gestor-cartera', 1); r.onupgradeneeded = () => r.result.createObjectStore('c'); r.onsuccess = () => ok(r.result); r.onerror = () => ko(r.error); });
  const tx = (mode, fn) => open().then(db => new Promise((ok, ko) => { const t = db.transaction('c', mode), q = fn(t.objectStore('c')); t.oncomplete = () => ok(q && q.result); t.onerror = () => ko(t.error); }));
  return {
    get: () => tx('readonly', st => st.get('carga')).catch(() => null),
    set: v => tx('readwrite', st => st.put(v, 'carga')).catch(() => {}),
    clear: () => tx('readwrite', st => st.delete('carga')).catch(() => {})
  };
})();
const SS = { get(k) { try { return sessionStorage.getItem('gc:' + k); } catch (e) { return null; } }, set(k, v) { try { sessionStorage.setItem('gc:' + k, v); } catch (e) {} } };
/* la copia sirve solo si la sesión guardada sigue vigente y es de la misma persona */
async function cachedStart() {
  try {
    const ses = JSON.parse(localStorage.getItem('gcx:sesion') || 'null'), c = await CACHE.get();
    if (!ses || ses.exp < Date.now() || !c || !c.d || c.name !== ses.name || c.token !== ses.token) return false;
    S.token = c.token; await loadAll(true, false, c.d); return true;
  } catch (e) { return false; }
}

async function boot() {
  const fast = await cachedStart();
  if (!fast) { $('#loading').hidden = false; $('#loadingTx').textContent = 'Revisando la estructura de la cartera…'; }
  try {
    const st = await call('api_status');
    if (!st.ok) {
      $('#probList').innerHTML = st.problems.map(p => `<li>${esc(p)}</li>`).join('');
      return showCover('#probCover');
    }
    $('#lgSheet').textContent = `Libro: hoja «${st.sheet}» · ${fmtN(st.rows)} clientes`;
    if (st.needSetup) { $('#suSheet').textContent = '«' + st.sheet + '»'; $('#suRows').textContent = fmtN(st.rows); return showCover('#setup'); }
    const ses = await call('api_resume');
    if (ses) { S.token = ses.token; return fast ? loadAll(false, true) : loadAll(true); }
    CACHE.clear(); showLogin();
  } catch (e) { showLogin('No se pudo leer el libro: ' + e.message); }
}

/* ---------- inicio ---------- */
applyTheme(); applyPanel();
const start = () => {
  boot().then(() => { if (!window.GC_BRIDGE && SS.get('win') === '1' && S.me && $('#login').hidden) showCover('#winLost'); });
  if (window.GC_BRIDGE) return GC_BRIDGE.onSelection(onSheetSelection);
  try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSheetSelection); } catch (e) {}
};
if (window.Office && Office.onReady) Office.onReady(info => {
  if (!window.GC_BRIDGE && info && info.host && info.host !== Office.HostType.Excel) { $('#loading').hidden = true; return showLogin('Este complemento funciona solo dentro de Excel.'); }
  start();
});
else showLogin('Abre este archivo desde Excel: Inicio → Complementos → Gestor de Cartera.');
})();
</script>"""
rep("/* ---------- inicio ---------- */\napplyTheme(); applyPanel();\nshowLogin();\n})();\n</script>", BOOT)

html = html.rstrip() + "\n</body>\n</html>\n"
panel = html.replace("<!--CAPA-DATOS-->", '<script src="datos.js"></script>\n<script src="puente.js"></script>')
ventana = html.replace("<!--CAPA-DATOS-->", '<script src="puente.js"></script>').replace("<title>Gestor de Cartera</title>", "<title>Gestor de Cartera · ventana</title>")
for out, txt in ((OUT, panel), (OUT_WIN, ventana)):
    out.write_text(txt, encoding="utf-8")
    print(f"OK {out.relative_to(AQUI.parent.parent)} · {len(txt.splitlines())} líneas")
