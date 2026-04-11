const express = require('express');
const axios = require('axios');
const crypto = require('crypto');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;
const CALTOPO_DEFAULT_DOMAIN = 'caltopo.com';
const CALTOPO_TIMEOUT_MS = 15000;
const CALTOPO_SIGNING_WINDOW_MS = 2 * 60 * 1000;

app.use(cors({
    origin: '*',
    methods: ['GET', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

const ensureHttpsDomain = (domain) => {
    const normalized = (domain || CALTOPO_DEFAULT_DOMAIN).trim().toLowerCase();
    if (!normalized || normalized.includes('/') || normalized.includes('\\') || normalized.includes('?')) {
        return CALTOPO_DEFAULT_DOMAIN;
    }
    return normalized;
};

const resolveCalTopoCredentials = () => {
    const credentialId = (process.env.CALTOPO_CREDENTIAL_ID || process.env.SARTOPO_CREDENTIAL_ID || '').trim();
    const credentialSecret = (process.env.CALTOPO_CREDENTIAL_SECRET || process.env.CALTOPO_SECRET || process.env.SARTOPO_SECRET || '').trim();

    return {
        credentialId,
        credentialSecret,
        configured: Boolean(credentialId && credentialSecret)
    };
};

const signCalTopoRequest = (method, endpoint, payloadString, credentialSecret) => {
    const expires = Date.now() + CALTOPO_SIGNING_WINDOW_MS;
    const message = `${method.toUpperCase()} ${endpoint}\n${expires}\n${payloadString || ''}`;
    const secret = Buffer.from(credentialSecret, 'base64');
    const signature = crypto.createHmac('sha256', secret).update(message).digest('base64');

    return {expires, signature};
};

const normalizeCalTopoState = (payload) => {
    if (!payload || typeof payload !== 'object') {
        return {
            type: 'FeatureCollection',
            features: []
        };
    }

    if (payload.type === 'FeatureCollection' && Array.isArray(payload.features)) {
        return payload;
    }

    if (payload.state && payload.state.type === 'FeatureCollection' && Array.isArray(payload.state.features)) {
        return payload.state;
    }

    if (Array.isArray(payload.features)) {
        return {
            type: 'FeatureCollection',
            features: payload.features,
            ids: payload.ids,
            timestamp: payload.timestamp
        };
    }

    return {
        type: 'FeatureCollection',
        features: []
    };
};

const fetchMapHandler = async (req, res) => {
    const {mapId, domain} = req.query;

    if (!mapId) {
        return res.status(400).json({
            error: 'Missing Map ID',
            message: 'Use /api/proxy?mapId=YOUR_MAP_ID&domain=caltopo.com'
        });
    }

    const trimmedMapId = String(mapId).trim();
    const targetDomain = ensureHttpsDomain(domain);
    const endpoint = `/api/v1/map/${encodeURIComponent(trimmedMapId)}/since/0`;
    const targetUrl = `https://${targetDomain}${endpoint}`;
    const creds = resolveCalTopoCredentials();

    if (!creds.configured) {
        return res.status(500).json({
            error: 'Proxy Not Configured',
            message: 'This proxy must sign CalTopo API requests server-side. Set CALTOPO_CREDENTIAL_ID and CALTOPO_CREDENTIAL_SECRET in the server environment.',
            targetUrl,
            mapId: trimmedMapId,
            signingRequired: true
        });
    }

    const {expires, signature} = signCalTopoRequest('GET', endpoint, '', creds.credentialSecret);

    try {
        console.log(`[PROXY] Fetching shapes from ${targetUrl}`);
        const response = await axios.get(targetUrl, {
            timeout: CALTOPO_TIMEOUT_MS,
            params: {
                id: creds.credentialId,
                expires,
                signature
            }
        });

        const normalizedState = normalizeCalTopoState(response.data);

        res.json({
            type: normalizedState.type,
            features: normalizedState.features || [],
            state: normalizedState,
            source: 'caltopo-signed-proxy',
            mapId: trimmedMapId,
            domain: targetDomain
        });
    } catch (error) {
        console.error(`[PROXY] Error fetching from ${targetUrl}:`, error.message);

        const responseStatus = error.response ? error.response.status : 500;
        const responseBody = error.response && error.response.data ? error.response.data : null;
        const detailMessage = typeof responseBody === 'string'
            ? responseBody.slice(0, 400)
            : responseBody && responseBody.message
                ? responseBody.message
                : error.message;

        res.status(responseStatus).json({
            error: error.response ? `CalTopo Error ${responseStatus}` : 'Proxy Connection Error',
            message: detailMessage,
            targetUrl,
            mapId: trimmedMapId,
            signingRequired: true,
            caltopoResponse: responseBody
        });
    }
};

app.get('/api/health', (req, res) => {
    const creds = resolveCalTopoCredentials();
    res.json({
        status: 'ok',
        version: '1.2.0',
        service: 'SARTopo Proxy',
        message: 'Proxy is live and ready for signed CalTopo requests',
        caltopoSigningConfigured: creds.configured,
        timestamp: new Date().toISOString()
    });
});

app.get('/api/proxy', fetchMapHandler);
app.get('/fetch-map', fetchMapHandler);

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Proxy v1.2.0 is live on port ${PORT}`);
});