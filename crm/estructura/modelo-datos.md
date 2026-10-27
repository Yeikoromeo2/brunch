# Modelo de datos del Gestor de Cartera

Todo vive en el mismo Excel de OneDrive. La cartera (`Hoja1`) conserva sus 18 columnas sin cambios (ver `columnas-cartera.json`). Lo demás se guarda en pestañas auxiliares, cuyo nombre empieza con `_`. Cada registro se enlaza con la cartera por **Customer ID**, nunca por número de fila.

## Pestañas auxiliares

### `_Seguimiento`: una fila por cliente
| Columna | Contenido |
|---|---|
| Customer ID | Clave del cliente |
| Revisado | SI / vacío |
| Resultado | Contestó · No contesta · Volver a llamar · Número errado · No interesado |
| Motivo | Motivo de «No interesado» |
| Intentos | Número de intentos de contacto |
| Compra / Monto | SI y monto en soles |
| Próximo contacto | Fecha para «Volver a llamar» |
| Tipo de negocio | Tienda · Oficina · Híbrido (vacío = sin definir) |
| Gestionado por | Ejecutivo que atiende al cliente |
| Primera gestión / Última gestión / Último contacto | Fecha y hora |
| Depurado | Fecha en que se marcó como depurado |

### `_ClienteMarcas`: una fila por cliente y marca
| Columna | Contenido |
|---|---|
| Customer ID | Clave del cliente |
| Marca | Nombre exacto del catálogo `_Marcas` |
| Relación | `habitual` (compra habitualmente) · `ocasional` (compra a veces) · `interes` (le interesa) |
| Registrado por | Ejecutivo que la registró o confirmó |
| Fecha | Última confirmación. Pasados 90 días se muestra «por confirmar» |

La línea de la marca no se repite aquí: se obtiene de `_Marcas`, así un cambio de línea se aplica a todos los clientes.

### `_Marcas`: catálogo
| Columna | Contenido |
|---|---|
| Marca | Nombre único |
| Línea | Una de las líneas de `marcas.json` («Por definir» si no se sabe) |
| Activa | SI / NO (las inactivas dejan de ofrecerse, pero se conservan en el historial) |

### `_Historial`: un movimiento por fila (solo se agrega, nunca se edita)
Customer ID · Fecha y hora · Usuario · Acción (resultado, comentario, marca, tipo, número corregido…) · Detalle.

### `_Telefonos`: números corregidos
Customer ID · Teléfono original · Teléfono corregido · Corregido por · Fecha. La columna `Cust. Telephone` de la cartera conserva el original.

### `_Usuarios`, `_Asignaciones`, `_Config`
- `_Usuarios`: nombre, rol (admin / ejecutivo) y activo.
- `_Asignaciones`: rangos de filas reservados por ejecutivo.
- `_Config`: ajustes del equipo (meta diaria, intentos para purga, visibilidad, WhatsApp…).

## Quién ve qué

| Rol | Ve en el panel, la lista, Marcas, Dashboard y Reportes |
|---|---|
| Ejecutivo | Solo **su cartera**: clientes con su nombre en `Ejecutivo`, en sus filas reservadas o gestionados por él |
| Administrador | Toda la cartera, más el mapeo de marcas por ejecutivo, usuarios y ajustes |

La opción se activa en `_Config` → «Cada ejecutivo ve solo su cartera».

**Límite importante:** en Excel para la web, quien tiene permiso de edición puede abrir la hoja y ver todas las filas. El complemento filtra lo que muestra, pero no reemplaza los permisos del archivo. Si se necesita privacidad real entre ejecutivos, la alternativa es un archivo por ejecutivo, consolidado por el administrador.

## Hojas que genera el sistema

Reportes, promociones y mapeos se crean como **hojas nuevas** con un nombre fechado. Si el nombre ya existe, se crea uno con «(2)», «(3)», etc. Cada hoja trae las 18 columnas de la cartera en el mismo orden y, a la derecha:

Fila · Prioridad · Tipo de negocio · Marcas habituales · Marcas a veces · Marcas de interés · Líneas · Marcas actualizadas · Gestionado por · Resultado · Intentos · Último contacto · Próximo contacto · Compra · Monto · Motivo · Teléfono corregido · Depurado
