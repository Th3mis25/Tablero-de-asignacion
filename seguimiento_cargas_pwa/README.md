# Tablero de seguimiento PWA

Aplicación PWA para seguimiento de cargas en tiempo real, conectada a Google Sheets.

## Cómo correr localmente
- Clonar el repo: `git clone https://github.com/TuUsuario/seguimiento_cargas_pwa.git`
- Copiar el archivo de ejemplo: `cp secure-config.sample.json secure-config.json` y colocar el token/API_TOKEN junto con el `usuario` y `password` que quieras utilizar para la pantalla de inicio de sesión. El archivo real queda fuera del control de versiones.
- Abrir `index.html` con Live Server (VS Code extension).
- Configura la URL de tu Google Apps Script en `config.js` o mediante la variable de entorno `API_BASE` al desplegar.

## Configuración
- URL de Google Apps Script: defínela en `config.js` o establece `API_BASE` como variable de entorno.
- Token de API: proporciónalo mediante `secure-config.json` (no versionado) o con la variable de entorno `API_TOKEN`/un endpoint seguro.
- Para evitar bloqueos por CORS durante el desarrollo local, la PWA envía el token como parámetro `token` en las peticiones (`GET` y `POST`). El backend de Apps Script sigue aceptando el encabezado `Authorization: Bearer` para pruebas manuales o integraciones externas.
- Credenciales de inicio de sesión: especifica `usuario`, `password` y opcionalmente `nombre` en `secure-config.json` para definir el acceso principal. También puedes declarar un arreglo `users` con objetos `{ "usuario": "...", "password": "...", "nombre": "..." }`/`{ "username": "..." }` para múltiples cuentas. Si no hay credenciales configuradas se usará el usuario por defecto `admin`/`admin123`.
- Nombre de la hoja: `Tabla_1`
- Columnas esperadas:
  - Trip, Caja, Referencia, Cliente, Destino, Estatus, Segmento, TR-MX, TR-USA, Cita carga, Llegada carga, Cita entrega, Llegada entrega, Comentarios, Docs, Tracking
- Las fechas deben enviarse en formato `DD/MM/YYYY HH:mm:ss` (ejemplo: `26/08/2025 22:00:00`).

## Inicio de sesión y control de acceso

Al abrir la PWA se muestra un overlay de inicio de sesión que bloquea el tablero hasta proporcionar credenciales válidas. La validación usa las credenciales definidas en alguna de las siguientes fuentes (en orden):

1. `secure-config.json` (generado desde `secure-config.sample.json` y excluido del repositorio).
2. El endpoint auxiliar configurado en `config.js`/`SECURE_CONFIG_URL` (por ejemplo el script `scripts/secure-config-server.js`).
3. Valores embebidos mediante `window.__ENV`.

Cada fuente puede definir un usuario principal con las llaves `usuario`, `password` y `nombre` o declarar un arreglo `users` con múltiples objetos que aceptan tanto `usuario`/`nombre` como `username`/`displayName`. Si ninguna fuente suministra credenciales se habilita el acceso con el usuario por defecto `admin`/`admin123`.

El servidor auxiliar (`scripts/secure-config-server.js`) expone por omisión las variables de entorno `API_TOKEN`, `SECURE_USER`/`USUARIO`, `SECURE_PASSWORD`/`CLAVE` y `SECURE_DISPLAY_NAME`/`NOMBRE`, devolviéndolas como `{ "API_TOKEN": "...", "usuario": "...", "password": "...", "nombre": "..." }`. Puedes ajustarlas según tu despliegue para replicar el comportamiento de `secure-config.json` sin guardar secretos en el repositorio.

## Verificación del Apps Script antes de usar la PWA

1. **Confirma el token configurado en Apps Script.**
   - En el editor de Apps Script ve a `Project Settings` → `Script properties` y asegúrate de que exista una propiedad `API_TOKEN`.
   - Copia ese valor y úsalo en `secure-config.json`, en la variable de entorno `API_TOKEN` o en tu endpoint seguro. Si difieren, la PWA recibirá respuestas `401 Unauthorized`.
2. **Verifica la URL del despliegue (`API_BASE`).**
   - Ingresa a `Deploy` → `Manage deployments`, abre el despliegue vigente y copia el campo **Web app URL**.
   - Usa exactamente esa URL en `config.js` o expórtala como `API_BASE` al generar `runtime-env.js`.
3. **Confirma los permisos del despliegue web.**
   - En el menú `Deploy` → `Manage deployments` selecciona tu despliegue, abre la sección **Web app** y establece **"Who has access"** en *Anyone* (sin necesidad de iniciar sesión). Esto evita respuestas sin encabezados CORS que provocarían errores de conexión en la PWA.
   - Cada vez que modifiques el código en Apps Script crea una nueva versión y actualiza el deployment para conservar estos permisos de acceso públicos.
4. **Prueba manualmente el endpoint de Apps Script.**
   - Ejecuta una petición GET para validar el token y la URL (sustituye los valores entre `<...>`):

     ```bash
     curl -X GET \
       -H "Authorization: Bearer <API_TOKEN>" \
       "https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec"
     ```

     Deberías obtener un JSON con la llave `data`. Si recibes `{"error":"Unauthorized"}`, revisa el token y repite la solicitud.
   - Opcionalmente realiza un POST simple para comprobar que se aceptan escrituras:

     ```bash
     curl -X POST \
       -H "Authorization: Bearer <API_TOKEN>" \
       -H "Content-Type: application/x-www-form-urlencoded" \
       -d "action=add&trip=225001&ejecutivo=QA" \
       "https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec"
     ```

     Usa un número de trip temporal y elimina la fila desde la hoja si no la necesitas.
5. **Actualiza la PWA con los valores verificados.**
   - Guarda el token en `secure-config.json` (o en tu fuente segura) y coloca la URL en `config.js`/`API_BASE` antes de abrir la aplicación.

Siguiendo estos pasos podrás detectar problemas de autenticación o de URL antes de levantar la PWA.

## Ejemplo de JSON de muestra
```json
[
  {"Trip":"229605","Caja":"8484","Referencia":"8151-8155-JN743","Cliente":"Yanfeng","Destino":"Huntsville, AL","Estatus":"Mty yard","Segmento":"OTR","TR-MX":"PATIO08","TR-USA":"","Cita carga":"26/08/2025 22:00:00","Llegada carga":"26/08/2025 9:59:00"},
  {"Trip":"229606","Caja":"8485","Referencia":"8151-8155-JN746","Cliente":"Yanfeng","Destino":"Huntsville, AL","Estatus":"Drop","Segmento":"OTR","TR-MX":"","TR-USA":"","Cita carga":"27/08/2025 21:00:00","Llegada carga":""}
]

```

## Despliegue

Para suministrar la URL de la API en producción puedes generar el archivo `config.js` a partir de una variable de entorno o bien definir `API_BASE` usando un archivo de entorno ejecutado antes del script:

```bash
export API_BASE="https://tu-web-app.example.com"
cat <<'JS' > runtime-env.js
window.__ENV = {
  API_BASE: '${API_BASE}'
};
JS
```

Incluye el archivo `runtime-env.js` antes de `<script src="./config.js"></script>` en tu `index.html` del entorno desplegado.

### Distribución del token `API_TOKEN`

`config.js` obtiene el token en el siguiente orden:
1. Variable de entorno `API_TOKEN` expuesta en `window.__ENV` (o en `process.env` si usas un bundler).
2. Un endpoint seguro definido en `window.__ENV.SECURE_CONFIG_URL` o `process.env.SECURE_CONFIG_URL` que devuelva `{ "API_TOKEN": "..." }`.
3. El archivo local `secure-config.json` (generado a partir de `secure-config.sample.json` y excluido del repositorio).

#### Opciones recomendadas para mantener el token fuera del repositorio

- **Archivo local no versionado**: duplica `secure-config.sample.json`, renómbralo a `secure-config.json` y coloca ahí tu token. El archivo ya está listado en `.gitignore`.
- **Variables de entorno durante el despliegue**:
  1. Exporta los secretos en tu pipeline:
     ```bash
     export API_TOKEN="super-token-secreto"
     export API_BASE="https://tu-web-app.example.com"
     cat <<'JS' > runtime-env.js
     window.__ENV = {
       API_BASE: '${API_BASE}',
       API_TOKEN: '${API_TOKEN}'
     };
     JS
     ```
  2. Asegúrate de servir `runtime-env.js` antes de `config.js` (puedes copiarlo junto al `index.html` generado por CI/CD y referenciarlo con `<script src="./runtime-env.js"></script>`).
- **Endpoint seguro**: hospeda un servicio autenticado que devuelva `{"API_TOKEN":"..."}` y expón su URL mediante `window.__ENV.SECURE_CONFIG_URL`. El token nunca se guardará en archivos estáticos.

### Ingresar el token desde la interfaz

Cuando `secure-config.json` no está disponible o el endpoint seguro devuelve un error, la aplicación mostrará una alerta roja con un formulario para capturar el token. El valor queda almacenado únicamente en el `localStorage` del navegador, por lo que deberás repetir el proceso en cada dispositivo/navegador que utilices.

Recuerda alinear el valor de `API_TOKEN` en tu Google Apps Script (`Project Settings` → `Script properties`) con el que uses en el cliente para evitar errores de autenticación.
