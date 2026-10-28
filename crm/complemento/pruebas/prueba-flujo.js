/* Prueba de punta a punta del complemento con el simulador de Excel.
   Uso:  node crm/complemento/pruebas/prueba-flujo.js [ruta/xlsx.full.min.js] [carpeta-capturas]
   Sirve los archivos del complemento desde disco, cambia office.js por el
   simulador y verifica que la hoja de la cartera solo cambie en «Comentario Créditos». */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const DIR = path.resolve(__dirname, '..');
const XLSX_JS = process.argv[2] && fs.existsSync(process.argv[2]) ? fs.readFileSync(process.argv[2], 'utf8') : null;
const SHOTS = process.argv[3] || null;

const COLS = ['Organization', 'KAM ID', 'Ejecutivo', 'Customer ID', 'Legal Name', 'Tax Reg No', 'Cust. Address1', 'State Name', 'City Name', 'Cust. Telephone', 'Cust. Mail Address', 'Term ID', 'Report Code', 'Credit Limit', 'Creation Date', 'Last Purchase', 'Clasif.', 'Comentario Créditos'];
const NAMES = ['Ferretería San Martín SAC', 'Distribuidora El Sol EIRL', 'Inversiones Kallpa SRL', 'Comercial Rímac SAC', 'Grupo Pachamama SAC', 'Bodega Don Lucho EIRL', 'Constructora Huascarán SAC', 'Textiles Majes SRL', 'Servicios Altiplano SAC', 'Importaciones Mistura SAC', 'Maderera Piura Norte EIRL', 'Botica Titicaca SAC'];
/* como en una cartera real: mayúsculas, espacios de más y el nombre completo en vez del usuario */
const execs = ['Harold', ' HAROLD ', 'harold', 'Luis', 'Luis', '', '', '', 'HAROLD QUISPE', '', 'Luis', ''];
const cartera = [COLS, ...NAMES.map((n, i) => ['PE01', 'K21' + (i % 4), execs[i], String(100101 + i), n, '20600' + (100101 + i), 'Av. Ejemplo ' + (100 + i), 'Lima', 'Lima', 987650000 + i * 111, `compras${i}@ejemplo.pe`, '30D', 'R1' + (i % 5), 5000 + i * 7000, 43500 + i * 30, 45900 + i * 20, ['A', 'B', 'C'][i % 3], i === 2 ? 'Tiene tienda en galería Wilson' : ''])];
const SEED = [{ name: 'Hoja1', rows: cartera }, { name: 'Resumen', rows: [['Notas'], ['hoja ajena al sistema']] }];

const fails = [];
let PAGE = null; const LOGS = [];
const ok = (cond, msg) => { console.log((cond ? 'OK    ' : 'FALLA ') + msg); if (!cond) fails.push(msg); };

(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage(); PAGE = page;
  const errors = [];
  page.on('pageerror', e => errors.push(e.message)); page.on('console', m => LOGS.push(m.text()));
  page.on('dialog', d => d.accept());
  /* trozos chicos para que el puente de la ventana pruebe el armado de mensajes grandes */
  await context.addInitScript(seed => { window.__seed = seed; window.GC_CHUNK = 2000; }, SEED);
  await context.route('**/*', route => {
    const u = new URL(route.request().url());
    if (u.hostname === 'appsforoffice.microsoft.com') return route.fulfill({ contentType: 'text/javascript', body: fs.readFileSync(path.join(__dirname, 'office-mock.js'), 'utf8') });
    if (u.hostname === 'cdnjs.cloudflare.com' && XLSX_JS) return route.fulfill({ contentType: 'text/javascript', body: XLSX_JS });
    if (u.hostname === 'gcx.local') { const f = path.join(DIR, u.pathname); return fs.existsSync(f) ? route.fulfill({ path: f }) : route.fulfill({ status: 404, body: '' }); }
    return route.abort();
  });
  const toast = async () => (await page.textContent('#toast')) || '';
  const shot = async n => { if (SHOTS) await page.screenshot({ path: path.join(SHOTS, n + '.png') }); };
  const aux = n => page.evaluate(x => window.__book.dump(x), n);

  await page.goto('https://gcx.local/taskpane.html');

  /* 1. primer uso */
  await page.waitForSelector('#setup:not([hidden])'); await page.waitForTimeout(120);
  ok(/12/.test(await page.textContent('#suRows')), 'Pantalla de configuración inicial con 12 clientes');
  await shot('01-configurar');
  await page.fill('#suName', 'Ana'); await page.fill('#suPin', '1234'); await page.fill('#suPin2', '1234');
  await page.click('#suBtn');
  await page.waitForFunction(() => document.querySelector('#meName').textContent === 'Ana');
  const vis = await page.evaluate(() => window.__book.raw.sheets.map(s => s.name + ':' + s.visibility));
  ok(vis.filter(x => x.startsWith('_')).length === 10 && vis.filter(x => x.startsWith('_')).every(x => x.endsWith(':VeryHidden')), '10 pestañas internas creadas como VeryHidden');
  ok(vis.includes('Hoja1:Visible') && vis.includes('Resumen:Visible'), 'Hoja1 y hojas ajenas siguen visibles');
  ok((await aux('_Marcas')).length === 90, 'Catálogo de 89 marcas en _Marcas');
  const u0 = await aux('_Usuarios');
  ok(u0[1][0] === 'Ana' && u0[1][1] === 'admin' && u0[1][3].length === 64 && u0[1][3] !== '1234', 'Administrador creado con PIN cifrado (no en texto)');

  /* 2. usuario con PIN temporal */
  await page.click('#bUsers');
  await page.waitForFunction(() => document.querySelectorAll('#dlExecs option').length === 3);
  ok(true, 'Usuarios ofrece los nombres de la columna Ejecutivo para vincular');
  await page.fill('#uName', 'Harold'); await page.fill('#uPin', '1111'); await page.fill('#uAlias', 'HAROLD QUISPE');
  await page.click('#uSave'); await page.waitForTimeout(150);
  ok(/4 clientes en la columna Ejecutivo/.test(await page.textContent('#uList')), 'El usuario muestra cuántos clientes le corresponden (mayúsculas, espacios y nombre en cartera)');
  const u1 = await aux('_Usuarios');
  ok(u1.some(r => r[0] === 'Harold' && r[5] === 'SI' && r[8] === 'HAROLD QUISPE'), 'Harold creado con PIN temporal y su nombre en cartera');
  await page.click('#mUsers [data-close]');

  /* 3. lista y panel */
  await page.click('#mainTabs [data-tab="list"]');
  await page.waitForSelector('#gbody .row');
  ok((await page.$$('#gbody .row')).length === 12, 'El administrador ve los 12 clientes');
  await page.click('#gbody .row[data-i="2"]');
  await page.waitForFunction(() => document.querySelector('#panel h2').textContent.includes('Kallpa'));
  ok((await page.textContent('#panel')).includes('Sugerido: Tienda'), 'Tipo de negocio sugerido por palabras clave');
  await shot('02-lista-panel');

  /* 4. comentario: se agrega sin borrar el anterior */
  await page.fill('#pCmt', 'Pidió precios de HP');
  await page.click('[data-act="cmtadd"]'); await page.waitForTimeout(200);
  const h1 = await aux('Hoja1');
  ok(/^\d\d\/\d\d Ana: Pidió precios de HP \| Tiene tienda en galería Wilson$/.test(h1[3][17]), 'Comentario agregado con fecha y usuario, conserva el anterior');

  /* 4b. sincronizar después de un cambio propio no recarga la cartera */
  await page.evaluate(() => { window.__loads = 0; const f = window.GC_API.api_load; window.GC_API.api_load = (...a) => { window.__loads++; return f(...a); }; });
  const sync = async () => { await page.evaluate(() => document.dispatchEvent(new Event('visibilitychange'))); await page.waitForTimeout(300); };
  await sync();
  ok((await page.evaluate(() => window.__loads)) === 0, 'Tras un cambio propio la sincronización no vuelve a cargar todo');
  /* 4c. un cambio de otra persona llega solo para ese cliente */
  await page.evaluate(() => {
    const seg = window.__book.sheet('_Seguimiento'), cfg = window.__book.sheet('_Config');
    const n = seg.cells.length; seg.cells[n] = ['100104', 'SI', 'Volver a llamar', '', '1', '2030-01-01', '', 'Luis', '', '', '', '', ''];
    const r = cfg.cells.find(x => x && x[0] === 'seq'); r[1] = String(+r[1] + 1);
  });
  await sync();
  ok(/Cambios del equipo/.test(await page.textContent('#syncTx')) && !(await page.textContent('#gbody .row[data-i="3"]')).includes('Volver a llamar'), 'En modo manual solo avisa «Cambios del equipo · Actualizar», sin leer el libro');
  await page.click('#sync'); await page.waitForTimeout(300);
  ok((await page.evaluate(() => window.__loads)) === 0 && await page.isHidden('#loading'), 'Al pulsar Actualizar trae el cambio de otro usuario sin recargar la pantalla');
  ok((await page.textContent('#gbody .row[data-i="3"]')).includes('Volver a llamar') && /Sincronizado/.test(await page.textContent('#syncTx')), 'La fila del cliente cambiado se actualiza');
  /* al seleccionar un cliente se lee al día solo ese cliente */
  await page.evaluate(() => { const seg = window.__book.sheet('_Seguimiento'); const r = seg.cells.find(x => x && x[0] === '100104'); r[2] = 'No interesado'; r[3] = 'Precio'; });
  await page.click('#gbody .row[data-i="3"]', { position: { x: 20, y: 20 } }); await page.waitForTimeout(800);
  ok((await page.textContent('#gbody .row[data-i="3"]')).includes('No interesado') && (await page.evaluate(() => window.__loads)) === 0, 'Al seleccionar un cliente se lee al día solo ese cliente');
  await page.click('#gbody .row[data-i="2"]', { position: { x: 20, y: 20 } }); await page.waitForTimeout(600);

  /* 5. resultado, tipo y marca */
  await page.click('[data-act="res"][data-n="Contestó"]'); await page.waitForTimeout(150);
  await page.click('[data-act="tipo"][data-n="Tienda"]'); await page.waitForTimeout(150);
  await page.fill('#pBrand', 'HP'); await page.press('#pBrand', 'Enter'); await page.waitForTimeout(200);
  const seg = await aux('_Seguimiento'), cm = await aux('_ClienteMarcas');
  ok(seg.some(r => r[0] === '100103' && r[2] === 'Contestó' && r[6] === 'Tienda' && r[7] === 'Ana'), '_Seguimiento guarda resultado, tipo y gestor por Customer ID');
  ok(cm.some(r => r[0] === '100103' && r[1] === 'HP' && r[2] === 'habitual'), '_ClienteMarcas guarda la marca con su relación');

  /* 5b. WhatsApp: la primera vez pregunta web o aplicación y luego recuerda la elección */
  await page.evaluate(() => { window.__opened = []; const o = window.open; window.__openedArgs = []; window.open = (u, ...a) => { window.__opened.push(u); window.__openedArgs.push(a); return /^https:\/\/gcx\.local/.test(u) ? o(u, ...a) : { close() { window.__closed = (window.__closed || 0) + 1; }, focus() {} }; }; });
  await page.click('#panel .p-actions [data-act="wa"]');
  await page.waitForSelector('#mWa.on'); await page.waitForTimeout(250); await shot('11-whatsapp');
  ok((await page.evaluate(() => window.__opened.length)) === 0, 'La primera vez WhatsApp pregunta antes de abrir');
  await page.click('#mWa [data-wa="app"]'); await page.waitForTimeout(200);
  await page.click('#panel .p-actions [data-act="wa"]'); await page.waitForTimeout(200);
  const opened = await page.evaluate(() => window.__opened);
  ok(opened.length === 2 && opened.every(u => /^whatsapp:\/\/send\?phone=51987650222/.test(u)) && !(await page.$('#mWa.on')), 'Recuerda «aplicación» y la segunda vez abre directo');
  ok((await aux('_Historial')).filter(r => r[0] === '100103' && /WhatsApp/.test(r[3])).length === 2, 'Cada apertura de WhatsApp queda como intento de contacto');
  await page.click('#bCfg');
  ok((await page.inputValue('#cWaMe')) === 'app', 'Ajustes muestra la elección guardada en este equipo');
  await page.click('#mCfg [data-close]');
  await page.waitForTimeout(2700);
  ok((await page.evaluate(() => window.__closed)) === 2, 'En modo aplicación la ventana en blanco se cierra sola');
  await page.evaluate(() => { const u = JSON.parse(localStorage.getItem('gcp:ui')); u.wa = 'web'; localStorage.setItem('gcp:ui', JSON.stringify(u)); });
  await page.reload(); await page.waitForFunction(() => document.querySelector('#meName').textContent === 'Ana');
  await page.evaluate(() => { window.__opened = []; window.__openedArgs = []; const o = window.open; window.open = (u, ...a) => { window.__opened.push(u); window.__openedArgs.push(a); return /^https:\/\/gcx\.local/.test(u) ? o(u, ...a) : { close() {}, focus() {} }; }; });
  await page.click('#mainTabs [data-tab="list"]'); await page.click('#gbody .row[data-i="2"]', { position: { x: 20, y: 20 } }); await page.waitForTimeout(200);
  await page.click('#panel .p-actions [data-act="wa"]'); await page.waitForTimeout(200);
  const wa = await page.evaluate(() => [window.__opened[0], window.__openedArgs[0]]);
  ok(/^https:\/\/web\.whatsapp\.com\/send\?phone=51987650222/.test(wa[0]) && wa[1][0] === 'gcWhatsApp' && /popup=yes,width=\d+/.test(wa[1][1]), 'WhatsApp Web se abre en una ventana al costado, siempre la misma');
  await page.evaluate(() => { const u = JSON.parse(localStorage.getItem('gcp:ui')); u.wa = 'app'; localStorage.setItem('gcp:ui', JSON.stringify(u)); });

  /* 6. venta con OV */
  await page.click('[data-act="ovopen"]');
  await page.fill('#oNum', 'ov-5001');
  await page.fill('#oQuick', 'HP 1500, Lenovo 2300'); await page.press('#oQuick', 'Enter');
  await page.click('[data-act="ovsave"]'); await page.waitForTimeout(250);
  const ord = await aux('_Ordenes'), ven = await aux('_Ventas');
  ok(ord.some(r => r[0] === 'OV-5001' && r[1] === '100103' && r[3] === '3800' && r[4] === 'Despachando'), 'OV registrada en Despachando con total sin IGV');
  ok(ven.filter(r => r[0] === 'OV-5001').length === 2, 'Una línea por marca en _Ventas');
  ok(!(await page.$('#bellCt[hidden]')), 'La campana avisa de la OV abierta');
  await page.click('[data-act="ovopen"]');
  await page.fill('#oNum', 'OV-5001'); await page.fill('#oQuick', 'HP 100'); await page.press('#oQuick', 'Enter');
  await page.click('[data-act="ovsave"]'); await page.waitForTimeout(200);
  ok(/ya está registrada/.test(await toast()), 'OV duplicada rechazada');
  await page.click('[data-act="ovopen"]');
  await page.click('[data-act="ovst"][data-n="OV-5001|Terminado"]'); await page.waitForTimeout(200);
  ok((await aux('_Ordenes')).some(r => r[0] === 'OV-5001' && r[4] === 'Terminado' && r[7] === 'Ana'), 'OV marcada como Terminado por quien la cerró');

  /* 7. número corregido */
  await page.click('#gbody .row[data-i="5"]');
  await page.click('[data-act="fixopen"]');
  await page.fill('#pNewPhone', '912345678');
  await page.click('[data-act="phsave"]'); await page.waitForTimeout(200);
  ok((await aux('_Telefonos')).some(r => r[0] === '100106' && r[2] === '912 345 678'), 'Teléfono corregido en _Telefonos, el original se conserva');

  /* 8. reserva de filas para Harold */
  await page.click('#bResv'); await page.waitForTimeout(120);
  await page.fill('#rvFrom', '11'); await page.fill('#rvTo', '13'); await page.selectOption('#rvUser', 'Harold');
  await page.click('#bDoResv'); await page.waitForTimeout(300);
  ok((await aux('_Asignaciones')).some(r => r[1] === 'Harold' && r[2] === '11' && r[3] === '13'), 'Reserva guardada en _Asignaciones');
  if (await page.$('#mResv.on')) await page.click('#mResv [data-close]');

  /* 9. dashboard y reporte en libro nuevo */
  await page.click('#mainTabs [data-tab="dash"]'); await page.waitForTimeout(300);
  await shot('03-dashboard');
  await page.click('#mainTabs [data-tab="rep"]');
  await page.click('[data-rt="all"]');
  await page.waitForTimeout(150);
  const before = await page.evaluate(() => window.__book.raw.sheets.length);
  await page.click('#rSheet'); await page.waitForTimeout(XLSX_JS ? 600 : 300);
  const created = await page.evaluate(() => window.__book.raw.created.length);
  ok(XLSX_JS ? created === 1 : true, 'Reporte creado como libro nuevo (Excel.createWorkbook)' + (XLSX_JS ? '' : ' · omitido, falta SheetJS'));
  ok((await page.evaluate(() => window.__book.raw.sheets.length)) === before, 'La cartera no recibe hojas nuevas');
  await shot('04-reportes');

  /* 10. selección en la hoja → abre el cliente */
  await page.click('#mainTabs [data-tab="list"]');
  await page.evaluate(() => window.__book.selectRow('Hoja1', 8));
  await page.waitForTimeout(300);
  ok((await page.textContent('#panel h2')).includes('Textiles Majes'), 'Seleccionar la fila 9 de la hoja abre ese cliente');

  /* 10b. ventana aparte: el panel queda de puente y guarda en el libro */
  let [win] = await Promise.all([page.waitForEvent('popup'), page.click('#bWin')]);
  win.on('pageerror', e => errors.push('ventana: ' + e.message)); win.on('dialog', d => d.accept());
  await win.setViewportSize({ width: 1500, height: 900 });
  await win.waitForFunction(() => document.querySelector('#meName').textContent === 'Ana');
  ok(await page.isVisible('#winCover'), 'El panel de Excel muestra «Abierto en otra ventana»');
  ok(await win.isHidden('#bWin'), 'La ventana no ofrece abrir otra ventana');
  await win.click('#mainTabs [data-tab="list"]');
  await win.waitForSelector('#gbody .row');
  ok((await win.$$('#gbody .row')).length === 12, 'La ventana carga la cartera completa a través del panel');
  await win.click('#gbody .row[data-i="7"]', { position: { x: 20, y: 20 } });
  /* desde la ventana: elegir clientes no mueve la selección de la hoja (no trae Excel al frente) y la
     aplicación de WhatsApp se lanza sin abrir ventanas */
  const selBefore = await page.evaluate(() => JSON.stringify(window.__book.raw.selection));
  await win.waitForTimeout(900);
  ok((await page.evaluate(() => JSON.stringify(window.__book.raw.selection))) === selBefore, 'Desde la ventana grande, elegir un cliente no mueve la selección de Excel');
  await win.evaluate(() => { window.__hrefs = []; window.__wopen = 0; HTMLAnchorElement.prototype.click = function () { window.__hrefs.push(this.href); }; const o = window.open; window.open = (...a) => { window.__wopen++; return o(...a); }; });
  await win.click('#panel .p-actions [data-act="wa"]'); await win.waitForTimeout(200);
  if (await win.$('#mWa.on')) { await win.click('#mWa [data-wa="app"]'); await win.waitForTimeout(200); }
  const wl = await win.evaluate(() => [window.__hrefs, window.__wopen]);
  ok(wl[0].length === 1 && /^whatsapp:\/\/send\?phone=51/.test(wl[0][0]) && wl[1] === 0, 'Desde la ventana grande la aplicación de WhatsApp se lanza sin abrir ventanas');
  await win.fill('#pCmt', 'Desde la ventana');
  await win.click('[data-act="cmtadd"]'); await win.waitForTimeout(300);
  ok(/^\d\d\/\d\d Ana: Desde la ventana$/.test((await aux('Hoja1'))[8][17]), 'Comentario hecho en la ventana queda guardado en el libro');
  await page.evaluate(() => window.__book.selectRow('Hoja1', 4));
  await win.waitForTimeout(400);
  ok((await win.textContent('#panel h2')).includes('Comercial Rímac'), 'Seleccionar una fila en la hoja abre el cliente en la ventana');
  if (SHOTS) await win.screenshot({ path: path.join(SHOTS, '09-ventana.png') });
  /* si el panel de Excel deja de responder, la ventana lo avisa en vez de quedarse guardando */
  await page.evaluate(() => { window.__fromChildOk = window.__fromChild; window.__fromChild = () => { throw new Error('panel cerrado'); }; });
  await win.click('#bReload');
  await win.waitForSelector('#lostCover:not([hidden])', { timeout: 5000 }).catch(() => {});
  ok(await win.isVisible('#lostCover'), 'Sin el panel de Excel la ventana muestra «Sin conexión con Excel»');
  if (SHOTS) await win.screenshot({ path: path.join(SHOTS, '10-ventana-sin-panel.png') });
  await page.evaluate(() => { window.__fromChild = window.__fromChildOk; });
  /* si Excel recarga el panel con la ventana abierta, la ventana se cierra y el panel ofrece reabrirla */
  await page.reload();
  await page.waitForSelector('#winLost:not([hidden])', { timeout: 10000 }).catch(() => {});
  ok(win.isClosed(), 'Si Excel recarga el panel, la ventana grande desconectada se cierra sola');
  ok(await page.isVisible('#winLost'), 'Al volver, el panel ofrece «Volver a abrir la ventana grande»');
  [win] = await Promise.all([page.waitForEvent('popup'), page.click('#winReopen')]);
  await win.waitForFunction(() => document.querySelector('#meName').textContent === 'Ana');
  ok(await page.isVisible('#winCover'), 'La ventana grande se vuelve a abrir con un clic y funciona');
  await page.click('#winClose');
  await page.waitForFunction(() => document.querySelector('#winCover').hidden && document.querySelector('#loading').hidden);
  ok(win.isClosed() && (await page.textContent('#meName')) === 'Ana', 'Al cerrar la ventana el panel vuelve a funcionar');

  /* 11. Harold entra con PIN temporal y ve solo su cartera */
  await page.click('#bOut');
  await page.waitForSelector('#login:not([hidden])'); await page.waitForTimeout(120);
  await page.fill('#lgName', 'harold'); await page.fill('#lgPin', '0000'); await page.click('#lgBtn'); await page.waitForTimeout(150);
  ok(/incorrecto/.test(await page.textContent('#lgErr')), 'PIN incorrecto rechazado');
  await page.fill('#lgPin', '1111'); await page.click('#lgBtn');
  await page.waitForSelector('#pinCover:not([hidden])'); await page.waitForTimeout(120);
  await shot('05-cambiar-pin');
  await page.fill('#pnPin', '2468'); await page.fill('#pnPin2', '2468'); await page.click('#pnBtn');
  await page.waitForFunction(() => document.querySelector('#meName').textContent === 'Harold');
  await page.click('#mainTabs [data-tab="list"]');
  const n = (await page.$$('#gbody .row')).length;
  ok(n === 7, `Harold ve solo su cartera: Ejecutivo + filas reservadas (${n} de 12)`);
  ok((await aux('_Usuarios')).some(r => r[0] === 'Harold' && r[5] === 'NO'), 'El PIN temporal quedó reemplazado');

  /* 11b. si Excel recarga el panel, la pantalla vuelve al instante desde la copia local */
  await page.click('#mainTabs [data-tab="brands"]');
  await page.evaluate(() => sessionStorage.setItem('__mockDelay', '400'));
  await page.reload();
  const fromCache = await page.waitForFunction(() => document.querySelector('#meName').textContent === 'Harold' && /Actualizando|Sincronizado/.test(document.querySelector('#syncTx').textContent) ? document.querySelector('#syncTx').textContent : false, null, { timeout: 5000 }).then(h => h.jsonValue());
  ok(fromCache === 'Actualizando…', 'Tras recargar el panel la cartera aparece desde la copia local mientras se actualiza');
  ok(await page.isVisible('#vBrands.on'), 'Tras recargar vuelve a la misma pestaña');
  await page.waitForFunction(() => /Sincronizado/.test(document.querySelector('#syncTx').textContent), null, { timeout: 8000 }).catch(() => {});
  ok(/Sincronizado/.test(await page.textContent('#syncTx')), 'Y se pone al día con el libro en segundo plano');
  await page.evaluate(() => sessionStorage.removeItem('__mockDelay'));
  await page.click('#mainTabs [data-tab="list"]');
  ok((await page.$$('#gbody .row')).length === 7, 'Sigue mostrando solo la cartera de Harold');

  /* 12. la hoja de la cartera solo cambió en «Comentario Créditos» */
  const fin = await aux('Hoja1');
  let diff = [];
  cartera.forEach((row, i) => row.forEach((x, j) => { if (j !== 17 && String(fin[i][j]) !== String(x)) diff.push(`${i + 1},${j + 1}`); }));
  ok(diff.length === 0 && fin.length === cartera.length && fin.every(r => r.length <= 18), 'Hoja1 intacta salvo la columna «Comentario Créditos»');
  const fmt = await page.evaluate(() => window.__book.sheet('Hoja1').formats.length);
  ok(fmt === 0, 'Sin cambios de formato en Hoja1');

  /* 13. panel angosto de Excel */
  await page.setViewportSize({ width: 420, height: 820 }); await page.waitForTimeout(300);
  await shot('06-taskpane-lista');
  await page.click('#gbody .row[data-i="0"]', { position: { x: 20, y: 20 } }); await page.waitForTimeout(200);
  ok(await page.isVisible('#panel .p-nav'), 'En el panel angosto el detalle trae «Lista» y anterior/siguiente');
  await shot('07-taskpane-detalle');
  const over = await page.evaluate(() => document.documentElement.scrollWidth - innerWidth);
  ok(over <= 0, 'Sin desplazamiento horizontal a 420 px');

  ok(errors.length === 0, 'Sin errores de JavaScript' + (errors.length ? ': ' + errors.join(' | ') : ''));
  await browser.close();
  console.log(fails.length ? `\n${fails.length} prueba(s) fallaron` : '\nTodas las pruebas pasaron');
  process.exit(fails.length ? 1 : 0);
})().catch(async e => {
  console.error(e.message.split('\n').slice(0, 4).join(' | '));
  try { console.error('Estado:', await PAGE.evaluate(() => ({ toast: document.querySelector('#toast').textContent, covers: [...document.querySelectorAll('.cover')].filter(c => !c.hidden).map(c => c.id), pnErr: document.querySelector('#pnErr').textContent, lgErr: document.querySelector('#lgErr').textContent, me: document.querySelector('#meName').textContent, runs: window.__book.raw.runs, pnBtn: document.querySelector('#pnBtn').disabled, pin: document.querySelector('#pnPin').value, users: window.__book.dump('_Usuarios') }))); } catch (er) {}
  console.error(LOGS.slice(-8).join('\n'));
  process.exit(1);
});
