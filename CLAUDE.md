# Control de Cargadores Eléctricos — Terra 93 PH (Apps Script)

Proyecto de Google Apps Script (`Code.js` + `Index.html` + `CSS.html` + `JS.html`)
que sirve el formulario web de registro de entrada/salida de carga de vehículos
eléctricos. Se despliega como Web App vía `doGet()`.

## Lección aprendida: Apps Script trunca en silencio la página de doGet()

`HtmlService` en Apps Script tiene un límite de tamaño para el HTML que devuelve
`doGet()`. Cuando el contenido total (HTML + CSS + JS embebidos, y sobre todo
cualquier dato pesado como imágenes/firmas en base64 metidas directo en el
template) se acerca o supera ese límite, Google **no lanza ningún error visible**:
la página simplemente se sirve incompleta o en blanco — solo el HTML estático que
alcanzó a salir antes del corte (por eso se vieron el título y los botones, pero
no el formulario inyectado por JS).

No hay mensaje de error, ni en el servidor ni en la consola del navegador que lo
señale directamente como "límite excedido" — parece un bug de renderizado, y hace
perder mucho tiempo diagnosticando algo que en realidad es tamaño de payload.

## Regla a seguir en este repo

**Cualquier contenido pesado (fotos, firmas, PDFs, imágenes en base64, listas
grandes de datos) se sirve por `google.script.run`, nunca embebido directo en
`Index.html`, `CSS.html` o `JS.html`.**

Esto significa:

- `doGet()` / los archivos de template (`Index.html`, `CSS.html`, `JS.html`)
  deben quedarse livianos: solo estructura, estilos y lógica de UI.
- Fotos, firmas de canvas, PDFs generados, o cualquier blob/base64 se
  transfieren **después** de la carga inicial, mediante llamadas
  `google.script.run.xxx()` desde el cliente hacia funciones del servidor
  (como ya hace `guardarEntrada`, `registrarSalida`, `subirArchivoDrive`).
- Nunca inyectar imágenes o archivos en base64 directamente dentro de un
  template `.html` que se evalúa en `doGet()`.
- Si se necesita mostrar una imagen ya guardada (ej. una foto de Drive), se
  pide bajo demanda con `google.script.run` (como hace
  `obtenerImagenBase64Drive`), no se precarga todo en el HTML inicial.

El workflow de despliegue (`.github/workflows/deploy.yml`) valida el tamaño
combinado de `Index.html` + `CSS.html` + `JS.html` antes de subir a Apps
Script, y falla el despliegue si se pasa del límite seguro — ver ese archivo
para el valor exacto y el mensaje de error.
