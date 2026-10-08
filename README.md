# Tablero de indicadores (Excel → página web)

Página web en tonos verdes que se construye **a partir del archivo Excel del tablero**.
No necesita servidor ni compilación: es HTML + JavaScript y el Excel se lee en el navegador.

## Qué muestra

| Sección | Contenido |
|---|---|
| **Resumen** | Réplica de la hoja *Dashboard* (o la hoja con más gráficos): indicadores clave (celdas y cuadros de texto vinculados), los mismos gráficos de Excel en la misma disposición y los botones con macro convertidos en accesos a las secciones. |
| **Una sección por hoja** | Indicadores con último valor, variación y minigráfico; gráfico configurable (barras, apiladas, líneas, área, dona, transponer, elegir columnas); tabla con búsqueda, orden y descarga CSV. Las hojas a las que llevan las macros y las que alimentan el tablero aparecen primero. |
| **Macros** | Módulos y procedimientos VBA: qué hacen (navegar, actualizar tablas dinámicas, filtrar, copiar…), qué hojas usan, botones/controles que los ejecutan y el código. |
| **Archivo y actualización** | Datos del archivo, hojas, nombres definidos e instrucciones. |

Los gráficos se dibujan con la paleta verde; el botón de paleta del encabezado cambia a los colores originales del Excel. Incluye modo oscuro e impresión/PDF.

## Cómo actualizar

1. Actualiza el Excel como siempre (ejecuta las macros y **guárdalo en Excel**, así los gráficos quedan con los valores al día).
2. Elige una opción:
   - **Solo en tu navegador:** botón **Subir Excel** o arrastra el archivo sobre la página. Queda guardado en ese navegador.
   - **Para todos:** copia el archivo a `data/` y pon su nombre en `data/config.json`:
     ```json
     { "organizacion": "Fedepalma", "titulo": "Tablero de indicadores", "archivo": "Dashboard.xlsm" }
     ```
     Para actualizar, reemplaza el archivo en `data/` con el mismo nombre.

## Ver la página

- **GitHub Pages:** *Settings → Pages → Deploy from a branch* y elige la rama y la carpeta raíz.
- **En el equipo:** `python3 -m http.server` en esta carpeta y abre <http://localhost:8000>.
  (Abriendo `index.html` con doble clic también funciona para subir archivos; solo la carga automática desde `data/` necesita un servidor.)

## Estructura

```
index.html
assets/css/styles.css     estilos y paleta verde (modo claro y oscuro)
assets/js/ooxml.js        lee gráficos, cuadros de texto, imágenes y botones del .xlsx/.xlsm
assets/js/analysis.js     detecta tablas e indicadores y resuelve los rangos de los gráficos
assets/js/vba.js          extrae y analiza las macros (vbaProject.bin)
assets/js/app.js          interfaz, secciones y gráficos (Chart.js)
assets/vendor/            SheetJS 0.18.5, JSZip 3.10.1, Chart.js 4.4.1
data/config.json          nombre del archivo publicado y títulos
```
