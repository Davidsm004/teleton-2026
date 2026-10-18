# Cómo volver a levantar el sitio de la Teletón 2026

El respaldo (`teleton-2026-respaldo-AAAA-MM-DD-HHMM.zip`) trae dos carpetas:

- `sitio/`: las páginas (dashboard, captura, librerías e informes). Corresponde a la rama **main**.
- `datos/`: los datos y los usuarios, **cifrados** (sin las contraseñas no se pueden leer). Corresponde a la rama **datos**.

## Volver a publicarlo en GitHub (el mismo repositorio u otro nuevo)
1. En GitHub, creá el repositorio (o usá el existente) y subí el contenido de `sitio/` a la rama `main`.
2. Creá la rama `datos` y subí `datos.json` y `usuarios.json` de la carpeta `datos/`.
3. En *Settings → Pages*, elegí la rama `main` y la carpeta raíz. En un par de minutos el sitio queda en línea.
4. Si el repositorio cambió de nombre o de dueño, pedile a quien administra el sitio que actualice la dirección de guardado y el token de GitHub.

Los usuarios y contraseñas siguen siendo los mismos que había al momento del respaldo.
