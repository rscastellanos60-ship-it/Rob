# Tablero de indicadores (Excel → página web)

Página web en tonos verdes que se construye **a partir del archivo Excel del tablero**.
No necesita servidor ni compilación: es HTML + JavaScript y el Excel se lee en el navegador.

## Qué muestra

| Sección | Contenido |
|---|---|
| **Resumen** | Encabezado de la hoja *Dashboard* (logo, título y fecha de corte), indicadores clave porcentuales y una tarjeta por sección con su gráfico principal. |
| **Secciones del Dashboard** | La hoja *Dashboard* se divide usando sus rótulos (Estructura, Mapa Social, Rotación General, Selección, Vacaciones, Ausentismo, Teletrabajo, Desarrollo, Seguridad y Salud en el Trabajo…). Cada sección muestra sus gráficos con el mismo tipo y orden que en Excel (incluidos embudos y cascadas), sus indicadores, sus tablas y las tablas pegadas como imagen (EMF). |
| **Una sección por hoja** | Indicadores con último valor, variación y minigráfico; gráfico configurable (barras, apiladas, líneas, área, dona, transponer, elegir columnas); tabla con búsqueda, orden y descarga CSV. Las hojas a las que llevan las macros y las que alimentan el tablero aparecen primero. |
| **Macros** | Módulos y procedimientos VBA: qué hacen (navegar, actualizar tablas dinámicas, filtrar, copiar…), qué hojas usan, botones/controles que los ejecutan y el código. |
| **Archivo y actualización** | Datos del archivo, hojas, nombres definidos e instrucciones. |

Los gráficos se dibujan con la paleta verde; el botón de paleta del encabezado cambia a los colores originales del Excel. Incluye modo oscuro e impresión/PDF.

## Ortografía

La página corrige al mostrar los errores de los textos del Excel (por ejemplo «Dsitribución» → «Distribución»,
«Antiguedad» → «Antigüedad», «Régimen», «período(s)», «Vacantes Cubiertas») y muestra las fechas en español
(«ene-26» en lugar de «Jan-26»). El archivo no se modifica. Las correcciones están en
`assets/js/correcciones.js` y se pueden agregar más en `data/config.json`:

```json
"correcciones": { "texto como está en Excel": "texto corregido" }
```

## Acceso con clave y privacidad

El Excel publicado está en `data/tablero.enc`, **cifrado** (AES-GCM de 256 bits, llave derivada de la clave
con PBKDF2-SHA256). Al abrir la página se pide la clave; el archivo se descifra solo en el navegador y la clave
no se envía a ningún servidor ni está guardada en el repositorio. Sin la clave, el archivo descargado no se
puede leer. La clave queda recordada solo mientras la pestaña esté abierta; el botón del candado vuelve a bloquear.

- Nunca subas el Excel sin cifrar a `data/`: este repositorio es **público** y contiene datos personales.
- Si la clave se comparte con alguien que no debería verla, cámbiala cifrando de nuevo el archivo
  (ten en cuenta que las versiones anteriores de `tablero.enc` siguen en el historial del repositorio).

## Cómo actualizar

1. Actualiza el Excel como siempre y **guárdalo en Excel** (así los gráficos quedan con los valores al día).
2. Abre `cifrar.html` en la página publicada (por ejemplo `…/Rob/cifrar.html`), elige el Excel, escribe la clave
   dos veces y descarga `tablero.enc`. Con una clave distinta, esa será la nueva clave del tablero.
3. En GitHub reemplaza `data/tablero.enc` (*Add file → Upload files*). En uno o dos minutos se ve publicado.

Para revisar otro archivo solo en tu navegador, entra con la clave y usa **Subir Excel**.

## Ver la página

- **GitHub Pages:** *Settings → Pages → Deploy from a branch* y elige la rama y la carpeta raíz.
- **En el equipo:** `python3 -m http.server` en esta carpeta y abre <http://localhost:8000>.
  (Abriendo `index.html` con doble clic también funciona para subir archivos; solo la carga automática desde `data/` necesita un servidor.)

## Estructura

```
index.html
assets/css/styles.css     estilos y paleta verde (modo claro y oscuro)
assets/js/ooxml.js        lee gráficos, cuadros de texto, imágenes y botones del .xlsx/.xlsm
assets/js/cifrado.js      cifra y descifra el Excel publicado con la clave
assets/js/correcciones.js ortografía y meses en español de los textos del Excel
assets/js/emf.js          dibuja las imágenes EMF (tablas copiadas como imagen) como SVG
assets/js/analysis.js     detecta tablas e indicadores y resuelve los rangos de los gráficos
assets/js/vba.js          extrae y analiza las macros (vbaProject.bin)
assets/js/app.js          interfaz, secciones y gráficos (Chart.js)
assets/vendor/            SheetJS 0.18.5, JSZip 3.10.1, Chart.js 4.4.1
data/config.json          títulos, archivo cifrado y secciones en torta
data/tablero.enc          Excel publicado, cifrado
cifrar.html               genera tablero.enc a partir de un Excel nuevo
```
