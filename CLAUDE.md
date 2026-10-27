# CLAUDE.md

Este repositorio contiene el firmware brunch y, en `crm/`, el sistema de gestión de cartera (CRM) del usuario. Responde siempre en español.

## Estructura de la cartera: regla fija

La cartera de clientes es la base de todo el CRM. Su estructura está definida en `crm/estructura/columnas-cartera.json`, con la plantilla original en `crm/estructura/plantilla-cartera.xlsx`.

- Las 18 columnas de `Hoja1` (de `Organization` a `Comentario Créditos`) nunca se agregan, quitan, renombran ni reordenan.
- **La hoja de la cartera no se modifica.** Lo único que el sistema escribe en `Hoja1` es la columna `Comentario Créditos`, siempre agregando (`dd/mm Usuario: texto`) sin borrar lo anterior. No escribe en `Ejecutivo`, no pinta filas ni cambia formatos.
- **Opción B (decidida):** como el usuario no tiene acceso a Entra ID para registrar una app de Microsoft Graph, todo dato adicional (seguimiento, historial, asignaciones, marcas, ventas, órdenes, teléfonos) se guarda en **pestañas internas «muy ocultas»** (`VeryHidden`) del mismo archivo, enlazadas por `Customer ID`. No aparecen en «Mostrar hojas». Si más adelante TI registra la app, esos datos se migran a un archivo paralelo sin cambiar su estructura.
- Cada fila es un cliente independiente. Se identifica por `Customer ID`, nunca por el número de fila, porque la hoja puede ordenarse o filtrarse.
- `Customer ID`, `Tax Reg No` y los teléfonos se tratan como texto.
- Antes de procesar un Excel real, validarlo con `python3 crm/estructura/validar_cartera.py <archivo.xlsx>`. Si no pasa, no procesarlo.
- La versión final trabaja sobre el Excel en OneDrive (complemento de Excel web). Los reportes y listas se abren como **libros nuevos** (`Excel.createWorkbook`), con las 18 columnas en el mismo orden; nunca se crean hojas visibles en el archivo de la cartera ni se borran datos.

## Ventas y órdenes (OV)

- Toda venta se registra con su **número de OV** (obligatorio y único) y una línea por marca con su monto.
- Los montos siempre son **sin IGV**.
- Al registrar una OV su estado pasa automáticamente a **Despachando**. La alerta sigue activa hasta que alguien la marque como **Terminado**; no se puede desactivar.
- Estados: `Despachando` → `Terminado`, con `Incidencia` para problemas de entrega. Solo el administrador puede **reabrir** una OV cerrada o **anularla**; las anuladas no cuentan en las ventas.
- El monto comprado de cada cliente es la suma de sus OV no anuladas; no se escribe a mano.
- Dashboard de ventas: período, marca, línea, tipo y ejecutivo como filtros; share de venta y de clientes por marca y por línea, venta mensual frente a los 12 meses anteriores, avance y proyección de la meta mensual, escala y estatus de compradores.

## Tipo de negocio y marcas

- Tipo de negocio de cada cliente: Tienda, Oficina o Híbrido. Se sugiere por palabras clave del comentario y se confirma en el panel; se guarda en `_Seguimiento`.
- El catálogo de marcas está en `crm/estructura/marcas.json` (pestaña `_Marcas`). Las marcas que distribuye cada cliente van en `_ClienteMarcas`, nunca como columnas de la cartera.
- Cada marca de un cliente lleva su relación (`habitual`, `ocasional`, `interes`), quién la registró y la fecha. Pasados 90 días se marca «por confirmar».

## Visibilidad

- El ejecutivo solo ve su cartera: clientes con su nombre en `Ejecutivo`, en sus filas reservadas o gestionados por él. El administrador ve todo, incluido el mapeo de marcas por ejecutivo.
- El modelo completo de pestañas auxiliares está en `crm/estructura/modelo-datos.md`.
