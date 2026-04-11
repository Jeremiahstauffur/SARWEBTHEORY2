const http = require('http');
const https = require('https');

const proxyUrl = 'https://sarwebtheory2-production.up.railway.app/api/proxy';
const activeMapId = 'test';
const activeMapDomain = 'caltopo.com';

let finalProxyUrl = `${proxyUrl}?mapId=${activeMapId}&domain=${activeMapDomain}`;
console.log('Final URL:', finalProxyUrl);

const url = new URL(finalProxyUrl);
const options = {
    method: 'GET',
    hostname: url.hostname,
    path: url.pathname + url.search
};

console.log('Options:', options);

const req = https.request(options, (res) => {
    console.log('Status:', res.statusCode);
    res.on('data', (d) => process.stdout.write(d));
});

req.on('error', (e) => {
    console.error('Error:', e);
});

req.end();
