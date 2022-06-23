'use strict';

const fs = require('fs/promises');
const getJson = require('./src/getJson');
const runJsonServer = require('./src/runJsonServer');

async function start() {
  const { JSON_URL } = JSON.parse(await fs.readFile('./consts.json'));
  const routes = JSON.parse(await fs.readFile('./routes.json'));
  const json = JSON.parse(await getJson(JSON_URL));
  runJsonServer(routes, json)
}

start();