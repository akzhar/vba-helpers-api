'use strict';

const os = require('os');
const jsonServer = require('json-server');
const PORT = process.env.PORT || 80;

function runJsonServer(routes, json) {
  const server = jsonServer.create();
  const router = jsonServer.router(json);
  const middlewares = jsonServer.defaults({ readOnly: true, host: os.hostname });

  server.use(middlewares);
  server.use(jsonServer.rewriter(routes));
  server.use(router);
  server.listen(PORT, () => console.log(`JSON server is running at http://${os.hostname}:${PORT}`));
}

module.exports = runJsonServer;