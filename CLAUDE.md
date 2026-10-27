# CLAUDE.md

Este repositorio contiene el firmware brunch y, en `crm/`, el sistema de gestión de cartera (CRM) del usuario. Responde siempre en español.

## Estructura de la cartera: regla fija

La cartera de clientes es la base de todo el CRM. Su estructura está definida en `crm/estructura/columnas-cartera.json`, con la plantilla original en `crm/estructura/plantilla-cartera.xlsx`.

- Las 18 columnas de `Hoja1` (de `Organization` a `Comentario Créditos`) nunca se agregan, quitan, renombran ni reordenan.
- El sistema solo escribe en dos columnas: `Ejecutivo` y `Comentario Créditos`. Todo dato adicional (seguimiento, historial, prioridad, motivos, montos) va en hojas o tablas aparte, nunca como columnas nuevas de la cartera.
- Cada fila es un cliente independiente. Se identifica por `Customer ID`, nunca por el número de fila, porque la hoja puede ordenarse o filtrarse.
- `Customer ID`, `Tax Reg No` y los teléfonos se tratan como texto.
- Antes de procesar un Excel real, validarlo con `python3 crm/estructura/validar_cartera.py <archivo.xlsx>`. Si no pasa, no procesarlo.
- En la base de datos del CRM (Supabase), la tabla de clientes conserva estas 18 columnas con el mismo significado, de modo que el Excel se pueda importar y exportar sin pérdidas.
