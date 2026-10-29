# Gestor de Cartera: guía de instalación

Complemento de Excel para la web que trabaja directamente sobre el archivo de la cartera en OneDrive. No necesita TI, servidores ni programas instalados.

## Qué hay en esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `taskpane.html` | El panel del Gestor (Mi día, Lista, Marcas, Dashboard, Reportes). Se genera con `construir.py` a partir del prototipo aprobado. |
| `ventana.html` | La misma interfaz en una ventana aparte, grande (botón **Ampliar**). También la genera `construir.py`. |
| `puente.js` | Conecta la ventana aparte con el panel: la ventana no puede tocar el libro, así que le pide todo al panel. |
| `datos.js` | Lee y escribe el libro con Office.js. Aquí están las reglas: la cartera no se modifica y todo lo demás va a pestañas muy ocultas. |
| `commands.html` | Archivo técnico que pide Excel para el botón de la cinta. |
| `instalar.html` | Página que genera el `manifest.xml` con la dirección correcta y explica cómo cargarlo. |
| `assets/` | Íconos del botón. |
| `construir.py` | Vuelve a generar `taskpane.html` si cambia el prototipo. |
| `pruebas/` | Simulador de Excel y prueba de punta a punta. |

## 1. Publicar los archivos (una sola vez, la hace el administrador)

Excel solo carga complementos desde una dirección `https://`. La forma gratuita es GitHub Pages:

1. En GitHub, abre el repositorio → **Settings** → **Pages**.
2. En **Source** elige **Deploy from a branch**, la rama con estos archivos y la carpeta `/ (root)`. Guarda.
3. A los pocos minutos queda publicado en
   `https://<usuario>.github.io/<repositorio>/crm/complemento/instalar.html`.

Para este repositorio, la dirección es `https://yeikoromeo2.github.io/brunch/crm/complemento/instalar.html`.

> En la cuenta gratuita de GitHub, Pages publica repositorios públicos. Los archivos publicados solo tienen el código del complemento, **ningún dato de clientes**: la cartera sigue en tu OneDrive.

## 2. Cargar el complemento (cada persona, una vez)

1. Abre `instalar.html` en el navegador y pulsa **Descargar manifest.xml**.
2. Abre el archivo de la cartera en **Excel para la web**.
3. **Inicio → Complementos → Mis complementos → Cargar mi complemento**, elige el `manifest.xml` y pulsa **Cargar**.
4. Aparece el botón **Gestor de Cartera** en la pestaña Inicio.

Si «Cargar mi complemento» no aparece, la organización lo tiene bloqueado y solo un administrador de Microsoft 365 puede habilitarlo.

## 3. Primer uso del archivo

1. Antes de usar un Excel real, valida su estructura:
   `python3 crm/estructura/validar_cartera.py cartera.xlsx`
   El complemento hace la misma revisión al abrirse y no continúa si falta, sobra o se movió una columna.
2. La primera persona que abre el panel ve **Configurar este libro**: escribe su nombre y un PIN y queda como **administrador**. En ese momento se crean las pestañas internas (`_Config`, `_Usuarios`, `_Seguimiento`, `_ClienteMarcas`, `_Marcas`, `_Ordenes`, `_Ventas`, `_Historial`, `_Telefonos`, `_Asignaciones`) como **muy ocultas**.
3. El administrador abre **Usuarios** y crea a cada ejecutivo con un **PIN temporal**. Al entrar por primera vez, el sistema le pide crear su propio PIN.
4. Cinco PIN incorrectos seguidos bloquean al usuario durante 5 minutos. Si alguien olvida su PIN, el administrador le asigna uno temporal nuevo.
5. **Si un administrador olvida su PIN:** otro administrador le asigna uno temporal. Si no hay otro, en la pantalla de ingreso pulsa **¿Olvidaste tu PIN? (administrador)** y usa el **código de recuperación** que te entregaron (sirve una sola vez por libro; cinco intentos fallidos bloquean 15 minutos). En el código publicado solo está la versión cifrada del código, nunca el código. Conviene tener siempre **dos administradores**.

## Mi día y borradores

- **Mi día** muestra solo tus pendientes: los clientes cuya última gestión hiciste tú y que quedaron en **Volver a llamar** (primero los vencidos y los de hoy) o **Revisados sin comentario**. Lo de tus compañeros no se mezcla. En la lista, el chip **Mis pendientes** filtra lo mismo.
- Lo que escribes en un comentario se guarda como **borrador** en tu equipo: no se pierde al tocar una marca u otra opción, al cambiar de cliente ni si Excel recarga el panel. La fila lo marca con un lápiz. Se escribe en Excel solo al pulsar «Agregar comentario».
- Cada fila de la lista y la barrita de arriba usan el color de su estado: compra (azul), atendido (verde), volver a llamar (violeta), no contesta (rojo), número errado (naranja), no interesado (gris), revisado (amarillo) y pendiente (sin color).

## Trabajar en una ventana grande

El panel de Excel tiene un ancho limitado. Con el botón **Ampliar** (arriba, en el panel) el Gestor se abre en una **ventana aparte** al 92 % de la pantalla, con la lista y el detalle lado a lado. La ventana se puede mover, maximizar o llevar a otro monitor.

- El panel de Excel muestra «Abierto en otra ventana» y **debe quedar abierto**: es el que lee y guarda en el libro. Si se cierra el panel, la ventana deja de guardar.
- Seleccionar una fila en la hoja abre ese cliente en la ventana.
- Desde la ventana, elegir un cliente **no** mueve la selección de la hoja (eso traería Excel al frente y taparía la ventana, sobre todo con dos pantallas). En el panel lateral sí la mueve.
- La ventana revisa cada 4 segundos que el panel de Excel siga respondiendo. Si deja de responder (Excel lo recargó o se cerró), **la ventana se cierra sola** en unos segundos; el panel recargado abre la nueva y los borradores se conservan. Lo que ya se guardó no se pierde.
- Para volver, cierra la ventana o pulsa **Cerrar la ventana y volver aquí** en el panel.
- Excel web a veces recarga el panel por su cuenta (abajo aparece «Cargando complementos»). Si pasa con la ventana abierta, la ventana vieja se cierra y **se reabre sola** al instante. Para que el navegador lo permita, autoriza una vez las ventanas emergentes de Excel:
  - **Edge:** Configuración → Cookies y permisos del sitio → **Elementos emergentes y redireccionamientos** → Permitir → Agregar `https://[*.]officeapps.live.com`, `https://[*.]office.com` y `https://[*.]sharepoint.com`.
  - **Chrome:** Configuración → Privacidad y seguridad → Configuración de sitios → **Ventanas emergentes y redirecciones** → Permitido → agregar los mismos sitios.
  - Si el navegador la bloquea igual, el panel muestra **Volver a abrir la ventana grande**: un clic y sigues donde estabas.
- Si Excel pregunta si permite abrir una ventana, responde **Permitir**.

## Cambios del equipo: manual o automático

Para que Excel trabaje liviano con varias personas a la vez, el Gestor **no relee el libro cada vez que otro cambia algo**:

- Cada 20 segundos solo revisa si hubo cambios (lee una celda). Si los hay, arriba aparece **«Cambios del equipo · Actualizar»**; al pulsarlo se traen solo los clientes que cambiaron, sin recargar la pantalla.
- **Al abrir un cliente** se lee al día solo ese cliente: estado, quién lo gestiona, comentario, marcas, OV e historial. Así siempre trabajas sobre lo último.
- Quien prefiera que se actualice solo lo cambia en **Ajustes → Cambios del equipo → Automáticamente**.
- Cada acción (comentario, resultado, OV) revisa el dato más reciente del libro justo antes de guardar, así que no se pisa el trabajo de otro.

## Si el panel «se duerme» y vuelve a cargar

Excel para la web recarga el panel del complemento cuando el navegador suspende la pestaña por inactividad o cuando Excel se reconecta. El Gestor guarda una copia de la última carga en el navegador: al recargarse, la cartera aparece al instante en la misma pestaña y el mismo cliente, y se pone al día sola en segundo plano («Actualizando…»). La copia se borra al pulsar **Salir**.

Para que el navegador no suspenda Excel:

- **Edge:** Configuración → Sistema y rendimiento → **Pestañas en suspensión** → «Nunca poner en suspensión estos sitios» → Agregar: `office.com`, `officeapps.live.com`, `sharepoint.com` y `onedrive.live.com`.
- **Chrome:** Configuración → Rendimiento → **Ahorro de memoria** → «Sitios que siempre deben estar activos» → Agregar los mismos sitios.

## Qué se escribe en el archivo

- **Hoja de la cartera:** solo la columna «Comentario Créditos», agregando `dd/mm Usuario: texto | comentario anterior`. Nunca se escribe en `Ejecutivo`, no se pintan filas ni se cambian formatos.
- **Pestañas muy ocultas:** seguimiento, marcas, OV, ventas, historial, teléfonos corregidos, reservas, usuarios y ajustes, enlazados por `Customer ID`.
- **Reportes:** cada reporte se crea como **libro nuevo** (Excel lo abre en otra pestaña) con las 18 columnas en el mismo orden y las de seguimiento a la derecha. La cartera no recibe hojas nuevas.

## Límites que conviene conocer

- Quien tiene permiso de edición del archivo puede abrir la hoja y ver todas las filas. El complemento solo filtra lo que muestra; la privacidad real entre ejecutivos requiere un archivo por ejecutivo.
- Las pestañas muy ocultas no aparecen en «Mostrar hojas», pero alguien con conocimientos técnicos podría leerlas con un script. Los PIN se guardan cifrados (SHA-256 con sal), nunca en texto.
- Los cambios de otros usuarios se ven al sincronizar (cada 20 segundos por defecto, se ajusta en **Ajustes**).
- La sesión dura 12 horas en el mismo navegador.

## Probar sin un Excel real

```
node crm/complemento/pruebas/prueba-flujo.js [ruta/a/xlsx.full.min.js] [carpeta-de-capturas]
```

Usa el Chromium y Playwright ya instalados y un simulador de Excel (`pruebas/office-mock.js`). Recorre la configuración inicial, usuarios con PIN temporal, comentarios, resultado, tipo de negocio, marcas, OV (duplicada y Terminado), teléfono corregido, reservas, reporte en libro nuevo, la selección de filas y la vista angosta. Al final comprueba que la hoja de la cartera solo cambió en «Comentario Créditos». Sin la ruta de SheetJS se omite solo la parte del reporte.

## Si se cambia el diseño

El diseño vive en `crm/prototipo-pro/index.html`. Después de cambiarlo:

```
python3 crm/complemento/construir.py
node crm/complemento/pruebas/prueba-flujo.js
```

`construir.py` falla con un mensaje claro si algún bloque del prototipo que reemplaza ya no existe.
