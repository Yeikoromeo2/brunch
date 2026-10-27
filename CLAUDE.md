# CLAUDE.md

Este repositorio contiene el firmware brunch y, en `crm/`, el sistema de gestión de cartera (CRM) del usuario. Responde siempre en español.

## Estructura de la cartera: regla fija

La cartera de clientes es la base de todo el CRM. Su estructura está definida en `crm/estructura/columnas-cartera.json`, con la plantilla original en `crm/estructura/plantilla-cartera.xlsx`.

- Las 18 columnas de `Hoja1` (de `Organization` a `Comentario Créditos`) nunca se agregan, quitan, renombran ni reordenan.
- El sistema solo escribe en dos columnas: `Ejecutivo` y `Comentario Créditos`. Todo dato adicional (seguimiento, historial, prioridad, motivos, montos) va en hojas o tablas aparte, nunca como columnas nuevas de la cartera.
- Cada fila es un cliente independiente. Se identifica por `Customer ID`, nunca por el número de fila, porque la hoja puede ordenarse o filtrarse.
- `Customer ID`, `Tax Reg No` y los teléfonos se tratan como texto.
- Antes de procesar un Excel real, validarlo con `python3 crm/estructura/validar_cartera.py <archivo.xlsx>`. Si no pasa, no procesarlo.
- La versión final trabaja directo sobre el Excel en OneDrive (complemento de Excel web). Los reportes y listas se crean como hojas nuevas con las 18 columnas en el mismo orden; nunca se borran datos de la cartera ni hojas anteriores.

## Tipo de negocio y marcas

- Tipo de negocio de cada cliente: Tienda, Oficina o Híbrido. Se sugiere por palabras clave del comentario y se confirma en el panel; se guarda en `_Seguimiento`.
- El catálogo de marcas está en `crm/estructura/marcas.json` (pestaña `_Marcas`). Las marcas que distribuye cada cliente van en `_ClienteMarcas`, nunca como columnas de la cartera.
