# Auditorías 5'S

Aplicación web para diligenciar el formato **AUDITORÍA 5'S** desde el celular o el computador y descargarlo en Excel con el mismo formato original (logo, colores, fórmulas de total, calificación y próxima auditoría).

## Qué hace

- Crea todas las auditorías que necesites; cada una queda guardada en el dispositivo.
- **Nueva iteración**: repite un proceso copiando proceso y equipo auditor; en el criterio 32 muestra las contramedidas de la auditoría anterior.
- Cumple / No cumple con botones grandes; al marcar *No cumple* pide anomalía, contramedida, responsable y plazo.
- Fotos de evidencia con la cámara o la galería (hasta 6 por criterio) → van a la hoja **EVIDENCIAS** del Excel.
- Resumen con % de cumplimiento, pendientes, no conformidades y fecha de la próxima auditoría (misma regla del Excel: ≤25 % → 7 días, ≤50 % → 14, ≤80 % → 21, >80 % → 30).
- Descargar o **compartir** el Excel (WhatsApp, Teams, correo) desde el celular.
- Descargar todas las auditorías en un .zip, y respaldo/restauración en .json para pasar los datos a otro equipo.
- Funciona sin conexión después de la primera visita y se puede instalar en la pantalla de inicio.
- Sin librerías externas: no depende de ningún CDN.

## Publicar en GitHub Pages

1. Crea un repositorio en GitHub (por ejemplo `auditoria-5s`).
2. Sube **todo el contenido de esta carpeta** a la raíz del repositorio (`index.html` debe quedar en la raíz).
3. En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, rama `main`, carpeta `/ (root)` → **Save**.
4. En uno o dos minutos la app queda en `https://TU-USUARIO.github.io/auditoria-5s/`.
5. En el celular, abre el enlace y usa **"Agregar a pantalla de inicio"** / **"Instalar app"**.

Cuando publiques cambios, sube el número de `VERSION` en `sw.js` para que los celulares tomen la versión nueva.

## Importante

- Los datos se guardan **en el navegador de cada dispositivo** (no en un servidor). Si se borran los datos del navegador, se pierden: usa *Guardar respaldo* de vez en cuando.
- Para cambiar el formato Excel base, reemplaza la plantilla en `js/template.js` (se genera a partir del .xlsx) y los criterios en `js/data.js`.

## Estructura

```
index.html            página
css/styles.css        estilos (mobile first, modo oscuro automático)
js/app.js             interfaz y lógica
js/data.js            los 32 criterios por sección
js/template.js        el Excel original embebido
js/xlsx-export.js     rellena la plantilla y genera el .xlsx
js/zip.js             empaquetador .zip mínimo
js/db.js              almacenamiento local (IndexedDB)
sw.js, manifest.webmanifest, icons/   app instalable y uso sin conexión
```
