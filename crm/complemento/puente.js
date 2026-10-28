/* Gestor de Cartera · puente entre el panel de Excel y la ventana aparte.
   Solo el panel de Excel puede leer y escribir el libro (Office.js no da acceso
   al libro desde una ventana de diálogo). Por eso:
   - En el panel (taskpane.html, junto a datos.js) este archivo define GC_HOST:
     abre la ventana y atiende sus llamadas con las mismas funciones api_*.
   - En la ventana (ventana.html, sin datos.js) define window.google.script.run,
     que envía cada llamada al panel y espera su respuesta. La interfaz no cambia.
   Los mensajes viajan como texto y se parten en trozos, porque Office limita su tamaño. */
(() => {
'use strict';
const CHUNK = () => window.GC_CHUNK || 30000;

/* parte un objeto en trozos de texto y los reensambla del otro lado */
function sender(post) {
  let n = 0;
  return obj => {
    const s = JSON.stringify(obj), id = 'm' + (++n) + '.' + Date.now(), size = CHUNK(), parts = Math.max(1, Math.ceil(s.length / size));
    for (let i = 0; i < parts; i++) post(JSON.stringify({ gc: 1, id, i, n: parts, d: s.slice(i * size, (i + 1) * size) }));
  };
}
function receiver(onObj) {
  const buf = {};
  return text => {
    let p; try { p = JSON.parse(text); } catch (e) { return; }
    if (!p || p.gc !== 1) return;
    const b = buf[p.id] || (buf[p.id] = []); b[p.i] = p.d;
    if (b.filter(x => x != null).length < p.n) return;
    delete buf[p.id];
    let obj; try { obj = JSON.parse(b.join('')); } catch (e) { return; }
    onObj(obj);
  };
}

if (window.GC_API) {
  /* ---------- lado del panel de Excel ---------- */
  let dlg = null, send = null, closed = null;
  const supported = () => { try { return !!(Office.context.ui && Office.context.ui.displayDialogAsync && Office.context.requirements.isSetSupported('DialogApi', '1.2')); } catch (e) { return false; } };
  function handle(m) {
    if (m.t !== 'call') return;
    const reply = (ok, r) => send && send({ t: 'res', id: m.id, ok, r });
    google.script.run.withSuccessHandler(r => reply(true, r)).withFailureHandler(e => reply(false, (e && e.message) || String(e)))[m.fn](...(m.args || []));
  }
  window.GC_HOST = {
    supported,
    get open() { return !!dlg; },
    /* abre la ventana; onClose se llama cuando la persona la cierra */
    openWindow(url, onClose) {
      return new Promise((ok, ko) => {
        Office.context.ui.displayDialogAsync(url, { width: 92, height: 92, displayInIframe: false, promptBeforeOpen: false }, res => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) return ko(new Error(res.error && res.error.message || 'No se pudo abrir la ventana.'));
          dlg = res.value; closed = onClose;
          send = sender(t => { try { dlg.messageChild(t); } catch (e) {} });
          const recv = receiver(handle);
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, ev => recv(ev.message));
          dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { const cb = closed; dlg = null; send = null; closed = null; if (cb) cb(); });
          ok(true);
        });
      });
    },
    close() { const d = dlg, cb = closed; dlg = null; send = null; closed = null; try { d && d.close(); } catch (e) {} if (cb) cb(); },
    /* cierra la ventana sin avisar (cuando Excel está por recargar este panel) */
    drop() { const d = dlg; dlg = null; send = null; closed = null; try { d && d.close(); } catch (e) {} },
    /* avisa a la ventana que cambió la selección en la hoja */
    selection() { if (send) send({ t: 'sel' }); }
  };
} else {
  /* ---------- lado de la ventana aparte ---------- */
  const pend = {}; let seq = 0, selCb = null, send = null;
  const LOST = 'Se perdió la conexión con el panel de Excel. Lo que hagas aquí no se guarda: cierra esta ventana y vuelve a abrir el Gestor desde Excel.';
  const ready = new Promise(ok => {
    Office.onReady(() => {
      send = sender(t => Office.context.ui.messageParent(t));
      const recv = receiver(m => {
        if (m.t === 'sel') { if (selCb) selCb(); return; }
        const p = pend[m.id]; if (!p) return; delete pend[m.id];
        if (m.ok) p.ok(m.r === undefined ? null : m.r); else p.ko(new Error(m.r));
      });
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, ev => recv(ev.message), () => ok());
    });
  });
  const runner = (ok, ko) => new Proxy({}, { get(_, fn) {
    if (fn === 'withSuccessHandler') return f => runner(f, ko);
    if (fn === 'withFailureHandler') return f => runner(ok, f);
    return (...args) => ready.then(() => {
      const id = ++seq;
      /* si el panel de Excel se cerró, la llamada nunca tendría respuesta: se avisa en vez de esperar */
      const fail = () => { const p = pend[id]; if (!p) return; delete pend[id]; p.ko(new Error(LOST)); };
      const timer = setTimeout(fail, 60000);
      pend[id] = { ok: r => { clearTimeout(timer); ok && ok(r); }, ko: e => { clearTimeout(timer); ko && ko(e); } };
      try { send({ t: 'call', id, fn, args }); } catch (e) { fail(); }
    });
  } });
  window.google = { script: { get run() { return runner(); } } };
  window.GC_BRIDGE = { onSelection(cb) { selCb = cb; } };
}
})();
