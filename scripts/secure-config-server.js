const http = require('http');

const port = process.env.PORT || 3000;
const API_TOKEN = process.env.API_TOKEN || '11fb2c10d4e9c7231cd84ed83aa4a716c94bd9d31ad6d8b5';
const USERNAME = process.env.SECURE_USER || process.env.API_USER || process.env.USUARIO || process.env.USERNAME || 'admin';
const PASSWORD = process.env.SECURE_PASSWORD || process.env.API_PASSWORD || process.env.CLAVE || process.env.PASSWORD || 'admin123';
const DISPLAY_NAME = process.env.SECURE_DISPLAY_NAME || process.env.NOMBRE || process.env.DISPLAY_NAME || 'Administrador';

http.createServer((req, res) => {
  if (req.url === '/secure-config.json') {
    res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
    const payload = {
      API_TOKEN,
      usuario: USERNAME,
      password: PASSWORD,
      nombre: DISPLAY_NAME,
      users: [
        {
          usuario: USERNAME,
          password: PASSWORD,
          nombre: DISPLAY_NAME
        }
      ]
    };
    res.end(JSON.stringify(payload));
    return;
  }
  res.writeHead(404);
  res.end('Not Found');
}).listen(port, () => {
  console.log(`Secure config server running on http://localhost:${port}`);
});
