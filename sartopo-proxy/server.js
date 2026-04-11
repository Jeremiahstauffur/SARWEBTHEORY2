const express = require('express');
const axios = require('axios');
const crypto = require('crypto');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;
const CALTOPO_DEFAULT_DOMAIN = 'caltopo.com';
const CALTOPO_TIMEOUT_MS = 30000;
const CALTOPO_SIGNING_WINDOW_MS = 5 * 60 * 1000;

const getTrimmedString = (value) => typeof value === 'string' ? value.trim() : '';

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json({limit: '1mb'}));

const ensureHttpsDomain = (domain) => {
    const normalized = (domain || CALTOPO_DEFAULT_DOMAIN).trim().toLowerCase();
    if (!normalized || normalized.includes('/') || normalized.includes('\\') || normalized.includes('?')) {
        return CALTOPO_DEFAULT_DOMAIN;
    }
    return normalized;
};

const resolveCalTopoCredentials = (requestCredentials = {}) => {
    const requestCredentialId = getTrimmedString(requestCredentials.credentialId);
    const requestCredentialSecret = getTrimmedString(requestCredentials.credentialSecret || requestCredentials.secret);
    const envCredentialId = (process.env.CALTOPO_CREDENTIAL_ID || process.env.SARTOPO_CREDENTIAL_ID || '').trim();
    const envCredentialSecret = (process.env.CALTOPO_CREDENTIAL_SECRET || process.env.CALTOPO_SECRET || process.env.SARTOPO_SECRET || '').trim();
    const useRequestCredentials = Boolean(requestCredentialId && requestCredentialSecret);
    const credentialId = useRequestCredentials ? requestCredentialId : envCredentialId;
    const credentialSecret = useRequestCredentials ? requestCredentialSecret : envCredentialSecret;

    return {
        credentialId,
        credentialSecret,
        configured: Boolean(credentialId && credentialSecret),
        source: useRequestCredentials
            ? 'request-body'
            : (envCredentialId && envCredentialSecret ? 'environment' : 'missing')
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

    // 1. Direct FeatureCollection (standard GeoJSON)
    if (payload.type === 'FeatureCollection' && Array.isArray(payload.features)) {
        return payload;
    }

    // 2. Nested FeatureCollection in 'state' (common Team API response)
    if (payload.state && payload.state.type === 'FeatureCollection' && Array.isArray(payload.state.features)) {
        const fc = payload.state;
        if (!fc.ids && payload.ids) fc.ids = payload.ids;
        if (!fc.timestamp && payload.timestamp) fc.timestamp = payload.timestamp;
        return fc;
    }

    // 3. Fallback: Aggregate features from typed arrays or 'state' object
    // CalTopo/SARTopo internal state often uses separate arrays for Marker, Shape, Assignment, etc.
    const state = payload.state || payload;
    const collectedFeatures = [];

    if (Array.isArray(state)) {
        // Direct array of features
        collectedFeatures.push(...state);
    } else if (state && typeof state === 'object') {
        if (Array.isArray(state.features)) {
            collectedFeatures.push(...state.features);
        } else {
            // Look for common typed arrays OR any array that might contain features
            // CalTopo standard types:
            const knownTypes = ['Marker', 'Shape', 'Assignment', 'Track', 'Route', 'Clue', 'Area', 'Line', 'Folder', 'Sector', 'Buffer'];
            
            // First check known types
            knownTypes.forEach(t => {
                if (Array.isArray(state[t])) {
                    state[t].forEach(item => {
                        if (item && typeof item === 'object') {
                            if (!item.type && !item.geometry && !item.class) item.class = t;
                            collectedFeatures.push(item);
                        }
                    });
                }
            });

            // Then check any other arrays (case-insensitive) just in case
            Object.keys(state).forEach(key => {
                if (Array.isArray(state[key]) && !knownTypes.includes(key) && key !== 'features' && key !== 'ids') {
                    state[key].forEach(item => {
                        if (item && typeof item === 'object') {
                            if (!item.type && !item.geometry && !item.class) item.class = key;
                            collectedFeatures.push(item);
                        }
                    });
                }
            });
        }
    }

    return {
        type: 'FeatureCollection',
        features: collectedFeatures,
        ids: payload.ids || (state && typeof state === 'object' ? state.ids : null),
        timestamp: payload.timestamp || (state && typeof state === 'object' ? state.timestamp : null)
    };
};

const fetchMapHandler = async (req, res) => {
    const requestData = (req.method === 'POST' && req.body && typeof req.body === 'object')
        ? req.body
        : req.query;
    const mapId = getTrimmedString(requestData.mapId);
    const domain = getTrimmedString(requestData.domain);
    const usePostToCalTopo = requestData.usePost || false;

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
    const creds = resolveCalTopoCredentials(requestData);

    if (!creds.configured) {
        return res.status(500).json({
            error: 'Proxy Not Configured',
            message: 'Provide credentialId and credentialSecret.',
            targetUrl,
            mapId: trimmedMapId,
            signingRequired: true
        });
    }

    const method = usePostToCalTopo ? 'POST' : 'GET';
    const payload = ''; // Empty for since endpoint
    const {expires, signature} = signCalTopoRequest(method, endpoint, payload, creds.credentialSecret);

    try {
        console.log(`[PROXY] Fetching shapes from ${targetUrl} (Method: ${method})`);
        console.log(`[PROXY] Credentials Source: ${creds.source}, ID: ${creds.credentialId.slice(0, 4)}...`);
        
        const axiosConfig = {
            timeout: CALTOPO_TIMEOUT_MS,
            params: {
                id: creds.credentialId,
                expires,
                signature,
                _: Date.now()
            }
        };

        let response;
        if (usePostToCalTopo) {
            response = await axios.post(targetUrl, {}, axiosConfig);
        } else {
            response = await axios.get(targetUrl, axiosConfig);
        }

        const dataSize = JSON.stringify(response.data).length;
        console.log(`[PROXY] Success! Response size: ${Math.round(dataSize / 1024)} KB, Status: ${response.status}`);

        const normalizedState = normalizeCalTopoState(response.data);

        res.json({
            type: normalizedState.type,
            features: normalizedState.features || [],
            state: normalizedState,
            source: 'caltopo-signed-proxy',
            credentialSource: creds.source,
            mapId: trimmedMapId,
            domain: targetDomain,
            caltopoMethod: method
        });
    } catch (error) {
        console.error(`[PROXY] Error fetching from ${targetUrl} (${method}):`, error.message);
        
        // If GET fails, try POST automatically if not already using it
        if (method === 'GET' && !usePostToCalTopo && (error.response?.status === 405 || error.response?.status === 403 || error.code === 'ECONNRESET')) {
            console.log(`[PROXY] GET failed, retrying with POST...`);
            // We'll call the handler again but with usePost=true
            req.body = { ...requestData, usePost: true };
            return fetchMapHandler(req, res);
        }

        const responseStatus = error.response ? error.response.status : 500;
        const responseBody = error.response && error.response.data ? error.response.data : null;
        
        console.error(`[PROXY] CalTopo Response Code: ${responseStatus}`, responseBody);

        res.status(responseStatus).json({
            error: error.response ? `CalTopo Error ${responseStatus}` : 'Proxy Connection Error',
            message: error.message,
            targetUrl,
            mapId: trimmedMapId,
            signingRequired: true,
            caltopoResponse: responseBody
        });
    }
};

const genericCallHandler = async (req, res) => {
    const { method = 'GET', endpoint, payload, domain } = req.body;
    
    if (!endpoint) {
        return res.status(400).json({ error: 'Missing endpoint' });
    }

    const targetDomain = ensureHttpsDomain(domain || req.body.domain);
    const targetUrl = `https://${targetDomain}${endpoint}`;
    const creds = resolveCalTopoCredentials(req.body);

    if (!creds.configured) {
        return res.status(500).json({ error: 'Proxy Not Configured' });
    }

    const payloadString = payload ? JSON.stringify(payload) : '';
    const { expires, signature } = signCalTopoRequest(method, endpoint, payloadString, creds.credentialSecret);

    try {
        console.log(`[PROXY] Generic ${method} to ${targetUrl}`);
        const axiosConfig = {
            timeout: CALTOPO_TIMEOUT_MS,
            params: {
                id: creds.credentialId,
                expires,
                signature
            },
            headers: {
                'Content-Type': 'application/json'
            }
        };

        let response;
        if (method.toUpperCase() === 'POST') {
            response = await axios.post(targetUrl, payload || {}, axiosConfig);
        } else if (method.toUpperCase() === 'DELETE') {
            response = await axios.delete(targetUrl, axiosConfig);
        } else {
            response = await axios.get(targetUrl, axiosConfig);
        }

        res.json(response.data);
    } catch (error) {
        console.error(`[PROXY] Error in generic call:`, error.message);
        const status = error.response ? error.response.status : 500;
        res.status(status).json(error.response ? error.response.data : { error: error.message });
    }
};

app.get('/api/health', (req, res) => {
    const creds = resolveCalTopoCredentials();
    res.json({
        status: 'ok',
        version: '1.4.0',
        service: 'SARTopo Proxy',
        message: 'Proxy is live and ready for signed CalTopo Team API requests',
        caltopoSigningConfigured: creds.configured,
        caltopoCredentialSource: creds.source,
        supportsClientSuppliedCredentials: true,
        timestamp: new Date().toISOString()
    });
});

app.get('/api/proxy', fetchMapHandler);
app.post('/api/proxy', fetchMapHandler);
app.post('/api/call', genericCallHandler);
app.get('/fetch-map', fetchMapHandler);
app.post('/fetch-map', fetchMapHandler);

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Proxy v1.3.0 is live on port ${PORT}`);
});