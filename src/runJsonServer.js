'use strict';

const os = require('os');
const jsonServer = require('json-server');
const http = require('http');
const https = require('https');
const fs = require('fs');

const isProduction = Boolean(process.env.NODE_ENV == 'production');
const HOST = isProduction ? os.hostname : 'localhost';
const PORT = isProduction ? (process.env.PORT || 443) : 3001;

function getSslOptions() {
  if (!isProduction) return {};

  return {
    key: fs.readFileSync('../../etc/letsencrypt/live/vbahelpers.ru/privkey.pem'),
    cert: fs.readFileSync('../../etc/letsencrypt/live/vbahelpers.ru/fullchain.pem')
  };
}

function runJsonServer(routes, routedData) {
  const app = jsonServer.create();
  const router = jsonServer.router(routedData);
  const middlewares = jsonServer.defaults({ readOnly: true, host: HOST });
  
  app.use(middlewares);
  app.use(jsonServer.rewriter(routes));
  app.use(router);

  const server = isProduction ? https.createServer(getSslOptions(), app) : http.createServer(app);

  server.listen(
    PORT,
    () => console.log(`JSON server is running at http${isProduction ? 's': ''}://${HOST}:${PORT}`)
  );
}

module.exports = runJsonServer;