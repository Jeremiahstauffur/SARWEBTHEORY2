const http = require('http');
const https = require('https');

const proxyUrl = process.argv[2] || 'http://127.0.0.1:3101/api/proxy';
const activeMapId = process.argv[3] || 'C34BK08';
const activeMapDomain = process.argv[4] || 'caltopo.com';
const credentialId = process.argv[5] || '';
const credentialSecret = process.argv[6] || '';

const url = new URL(proxyUrl);
const transport = url.protocol === 'https:' ? https : http;
const requestBody = {
    mapId: activeMapId,
    domain: activeMapDomain
};

if (credentialId && credentialSecret) {
    requestBody.credentialId = credentialId;
    requestBody.credentialSecret = credentialSecret;
}

const payload = JSON.stringify(requestBody);
const options = {
    method: 'POST',
    hostname: url.hostname,
    port: url.port || (url.protocol === 'https:' ? 443 : 80),
    path: url.pathname + url.search,
    headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
    }
};

console.log('Request URL:', proxyUrl);
console.log('Payload:', JSON.stringify(requestBody, null, 2));

const req = transport.request(options, (res) => {
    let responseBody = '';
    console.log('Status:', res.statusCode);

    res.on('data', (chunk) => {
        responseBody += chunk;
    });

    res.on('end', () => {
        console.log(responseBody || '<empty body>');
    });
});

req.on('error', (error) => {
    console.error('Error:', error);
});

req.write(payload);
req.end();
