// Configuración global de la aplicación.
// Este archivo puede ser reemplazado en el despliegue para ajustar la URL de la API
// y la forma en la que se obtiene el token seguro.
(function configureApp(global){
  const runtimeEnv = global && typeof global.__ENV === 'object' ? global.__ENV : {};
  const processEnv = typeof process !== 'undefined' && process ? process.env || {} : {};

  const readEnvValue = key => {
    const runtimeValue = runtimeEnv[key];
    if(typeof runtimeValue === 'string' && runtimeValue.trim()){
      return runtimeValue.trim();
    }
    const processValue = processEnv[key];
    if(typeof processValue === 'string' && processValue.trim()){
      return processValue.trim();
    }
    return '';
  };

  // URL predeterminada del Web App activo de Apps Script.
  const DEFAULT_API_BASE = 'https://script.google.com/macros/s/AKfycbw9LokY2zUIyjq30SL4GZQ_KzkH2yGRPCSoWiG7F1tFtDva_0JjKgb3wIrs1b0F2o-g/exec';

  const DEFAULT_API_TOKEN = '11fb2c10d4e9c7231cd84ed83aa4a716c94bd9d31ad6d8b5';

  const apiBase = readEnvValue('API_BASE') || DEFAULT_API_BASE;
  const inlineToken = readEnvValue('API_TOKEN') || DEFAULT_API_TOKEN;
  const secureUrlOverride = readEnvValue('SECURE_CONFIG_URL');

  let secureConfigUrl = secureUrlOverride || 'secure-config.json';

  if(inlineToken){
    const json = JSON.stringify({ API_TOKEN: inlineToken });
    secureConfigUrl = `data:application/json,${encodeURIComponent(json)}`;
  }

  global.APP_CONFIG = {
    // URL del Web App de Google Apps Script.
    API_BASE: apiBase,
    // Archivo o endpoint seguro con los usuarios autorizados y el token de API.
    SECURE_CONFIG_URL: secureConfigUrl
  };
})(typeof window !== 'undefined' ? window : globalThis);
