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

## Trabajar en una ventana grande

El panel de Excel tiene un ancho limitado. Con el botón **Ampliar** (arriba, en el panel) el Gestor se abre en una **ventana aparte** al 92 % de la pantalla, con la lista y el detalle lado a lado. La ventana se puede mover, maximizar o llevar a otro monitor.

- El panel de Excel muestra «Abierto en otra ventana» y **debe quedar abierto**: es el que lee y guarda en el libro. Si se cierra el panel, la ventana deja de guardar.
- Seleccionar una fila en la hoja abre ese cliente en la ventana.
- Para volver, cierra la ventana o pulsa **Cerrar la ventana y volver aquí** en el panel.
- Si Excel pregunta si permite abrir una ventana, responde **Permitir**.

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
