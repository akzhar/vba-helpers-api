'use strict';

const os = require('os');
const jsonServer = require('json-server');
const https = require('https');
const fs = require('fs');

const PORT = process.env.PORT || 443;
const isProduction = Boolean(process.env.NODE_ENV == 'production');

const sslKeyPath = isProduction ? '../../etc/ssl/private.key' : './ssl/private.key';
const sslCertPath = isProduction ? '../../etc/ssl/fullchain.crt' : './ssl/fullchain.crt';

const sslOptions = {
  key: fs.readFileSync(sslKeyPath),
  cert: fs.readFileSync(sslCertPath)
};

function runJsonServer(routes, routedData) {
  const app = jsonServer.create();
  const router = jsonServer.router(routedData);
  const middlewares = jsonServer.defaults({ readOnly: true, host: os.hostname });
  
  app.use(middlewares);
  app.use(jsonServer.rewriter(routes));
  app.use(router);

  const server = https.createServer(sslOptions, app);
  server.listen(PORT, () => console.log(`JSON server is running at https://${isProduction ? os.hostname : 'localhost'}:${PORT}`));
}

module.exports = runJsonServer;