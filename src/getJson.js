'use strict';

const https = require('https');

async function getJson(url) {
  return new Promise((resolve, reject) => {
    try {
      https.get(url, (res) => {
        let json = '';
        res.on('data', chunk => json += chunk);
        res.on('end', () => resolve(json));
      });
    } catch(err) {
      reject(`JSON request error: ${err}`);
    }
  });
}

module.exports = getJson;