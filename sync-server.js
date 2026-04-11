const axios = require('axios');
const crypto = require('crypto');
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'db');
const CALTOPO_DEFAULT_DOMAIN = 'caltopo.com';
const CALTOPO_TIMEOUT_MS = 30000;
const CALTOPO_SIGNING_WINDOW_MS = 5 * 60 * 1000;

const getTrimmedString = (value) => typeof value === 'string' ? value.trim() : '';

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR);
}

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-User-Name', 'X-User-Pin']
}));
app.use(express.json({limit: '50mb'}));

const ensureHttpsDomain = (domain) => {
    const normalized = (domain || CALTOPO_DEFAULT_DOMAIN).trim().toLowerCase();
    if (!normalized || normalized.includes('/') || normalized.includes('\\') || normalized.includes('?')) {
        return CALTOPO_DEFAULT_DOMAIN;
    }
    return normalized;
};

const signCalTopoRequest = (method, endpoint, payloadString, credentialSecret) => {
    const expires = Date.now() + CALTOPO_SIGNING_WINDOW_MS;
    const message = `${method.toUpperCase()} ${endpoint}\n${expires}\n${payloadString || ''}`;
    const secret = Buffer.from(credentialSecret, 'base64');
    const signature = crypto.createHmac('sha256', secret).update(message).digest('base64');

    return {expires, signature};
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

// Helper to get file path
const getFilePath = (bucket, key) => {
    const bucketDir = path.join(DATA_DIR, bucket);
    if (!fs.existsSync(bucketDir)) {
        fs.mkdirSync(bucketDir);
    }
    // Sanitize key to prevent directory traversal
    const safeKey = key.replace(/[^a-z0-9_-]/gi, '_');
    return path.join(bucketDir, `${safeKey}.json`);
};

// List all keys in a bucket
app.get('/api/v1/:bucket', (req, res) => {
    const {bucket} = req.params;
    const bucketDir = path.join(DATA_DIR, bucket);

    if (!fs.existsSync(bucketDir)) {
        return res.json([]);
    }

    try {
        const files = fs.readdirSync(bucketDir);
        const keys = files
            .filter(f => f.endsWith('.json'))
            .map(f => f.replace('.json', ''));
        res.json(keys);
    } catch (err) {
        res.status(500).json({error: 'Failed to list keys'});
    }
});

// Get the most recently updated file in a bucket
app.get('/api/v1/:bucket/latest', (req, res) => {
    const {bucket} = req.params;
    const bucketDir = path.join(DATA_DIR, bucket);

    if (!fs.existsSync(bucketDir)) {
        return res.status(404).json({error: 'Bucket not found'});
    }

    try {
        const files = fs.readdirSync(bucketDir)
            .filter(f => f.endsWith('.json') && f !== 'all-files.json' && f !== 'bundle.json');
        
        if (files.length === 0) {
            // Fallback to bundle.json if it exists
            const bundlePath = path.join(bucketDir, 'bundle.json');
            if (fs.existsSync(bundlePath)) {
                return res.json(JSON.parse(fs.readFileSync(bundlePath, 'utf8')));
            }
            return res.status(404).json({error: 'No data files found'});
        }

        let latestFile = null;
        let latestTime = 0;

        files.forEach(f => {
            const filePath = path.join(bucketDir, f);
            const metaPath = filePath + '.meta';
            let updatedAt = 0;

            if (fs.existsSync(metaPath)) {
                try {
                    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
                    updatedAt = new Date(meta.updatedAt).getTime();
                } catch (e) {
                    updatedAt = fs.statSync(filePath).mtimeMs;
                }
            } else {
                updatedAt = fs.statSync(filePath).mtimeMs;
            }

            if (updatedAt > latestTime) {
                latestTime = updatedAt;
                latestFile = f;
            }
        });

        // Also check bundle.json for its time
        const bundlePath = path.join(bucketDir, 'bundle.json');
        if (fs.existsSync(bundlePath)) {
            const bundleMetaPath = bundlePath + '.meta';
            let bundleTime = 0;
            if (fs.existsSync(bundleMetaPath)) {
                try {
                    bundleTime = new Date(JSON.parse(fs.readFileSync(bundleMetaPath, 'utf8')).updatedAt).getTime();
                } catch (e) {
                    bundleTime = fs.statSync(bundlePath).mtimeMs;
                }
            } else {
                bundleTime = fs.statSync(bundlePath).mtimeMs;
            }

            if (bundleTime > latestTime) {
                latestTime = bundleTime;
                latestFile = 'bundle.json';
            }
        }

        if (latestFile) {
            const data = fs.readFileSync(path.join(bucketDir, latestFile), 'utf8');
            res.json(JSON.parse(data));
        } else {
            res.status(404).json({error: 'No files found'});
        }
    } catch (err) {
        console.error('Error finding latest file:', err);
        res.status(500).json({error: 'Internal server error'});
    }
});

// Get a value
app.get('/api/v1/:bucket/:key', (req, res) => {
    const {bucket, key} = req.params;
    const filePath = getFilePath(bucket, key);

    if (!fs.existsSync(filePath)) {
        return res.status(404).json({error: 'Not found'});
    }

    try {
        const data = fs.readFileSync(filePath, 'utf8');
        res.json(JSON.parse(data));
    } catch (err) {
        res.status(500).json({error: 'Failed to read data'});
    }
});

// Set a value
app.put('/api/v1/:bucket/:key', (req, res) => {
    const {bucket, key} = req.params;
    const filePath = getFilePath(bucket, key);
    const metaPath = filePath + '.meta';

    const userName = req.headers['x-user-name'] || 'Unknown';
    const userPin = req.headers['x-user-pin'] || '';
    const isSuperAdmin = userPin === '1976';

    let incomingLastModified = 0;
    if (req.body) {
        if (req.body.lastModified) {
            incomingLastModified = new Date(req.body.lastModified).getTime();
        } else if (typeof req.body === 'object') {
            // Try to find latest modified time in a collection of files
            for (const key in req.body) {
                if (req.body[key] && req.body[key].lastModified) {
                    const m = new Date(req.body[key].lastModified).getTime();
                    if (m > incomingLastModified) incomingLastModified = m;
                }
            }
        }
    }

    if (fs.existsSync(metaPath)) {
        try {
            const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
            const currentIsSuperAdmin = meta.userPin === '1976';
            const existingLastModified = new Date(meta.updatedAt).getTime();

            // Super-Admin priority
            if (currentIsSuperAdmin && !isSuperAdmin) {
                return res.status(403).json({
                    error: 'Conflict',
                    message: 'Changes by Super-Admin cannot be overwritten by a regular user.'
                });
            }

            // Conflict resolution (same level or Super-Admin overwriting anyone)
            if (isSuperAdmin === currentIsSuperAdmin) {
                if (incomingLastModified < existingLastModified) {
                    return res.status(403).json({
                        error: 'Conflict',
                        message: 'Incoming data is older than server data.'
                    });
                }
                
                if (incomingLastModified === existingLastModified) {
                    const incoming = userName.toLowerCase();
                    const existing = meta.userName.toLowerCase();
                    if (incoming !== existing && incoming > existing) {
                        return res.status(403).json({
                            error: 'Conflict',
                            message: `Changes by ${meta.userName} have priority (alphabetically closer to A).`
                        });
                    }
                }
            }
        } catch (err) {
            console.error('Failed to read meta file:', err);
        }
    }

    const saveTime = (incomingLastModified && incomingLastModified > 0) 
        ? new Date(incomingLastModified).toISOString() 
        : new Date().toISOString();

    try {
        fs.writeFileSync(filePath, JSON.stringify(req.body, null, 2));
        fs.writeFileSync(metaPath, JSON.stringify({
            userName,
            userPin,
            updatedAt: saveTime
        }, null, 2));
        res.json({success: true});
    } catch (err) {
        res.status(500).json({error: 'Failed to save data'});
    }
});

// Root endpoint for health check
app.get('/', (req, res) => {
    res.send('SAR Sync + Proxy Server is running');
});

// Health check endpoint for the proxy
app.get('/api/health', (req, res) => {
    const creds = resolveCalTopoCredentials();
    res.json({
        status: 'ok',
        version: '1.3.0',
        service: 'SAR Proxy + Sync',
        message: 'Unified server is live and ready to sign CalTopo Team API requests',
        caltopoSigningConfigured: creds.configured,
        caltopoCredentialSource: creds.source,
        supportsClientSuppliedCredentials: true,
        timestamp: new Date().toISOString()
    });
});

// CalTopo Proxy endpoint
const fetchMapHandler = async (req, res) => {
    const requestData = req.method === 'POST' && req.body && typeof req.body === 'object'
        ? req.body
        : req.query;
    const mapId = getTrimmedString(requestData.mapId);
    const domain = getTrimmedString(requestData.domain);
    const usePostToCalTopo = requestData.usePost || false;

    if (!mapId) {
        return res.status(400).json({
            error: "Missing mapId parameter",
            message: "Please ensure your Map ID is correctly entered in the Maps page."
        });
    }

    const trimmedMapId = String(mapId).trim();
    const targetDomain = ensureHttpsDomain(domain);
    const endpoint = `/api/v1/map/${trimmedMapId}/since/0`;
    const targetUrl = `https://${targetDomain}${endpoint}`;
    const creds = resolveCalTopoCredentials(requestData);

    if (!creds.configured) {
        return res.status(500).json({
            error: 'Proxy Not Configured',
            message: 'This proxy needs a CalTopo Credential ID and Credential Secret to sign the Team API request. Provide them in the server environment, or send credentialId and credentialSecret in this POST body.',
            targetUrl,
            mapId: trimmedMapId,
            signingRequired: true,
            supportsClientSuppliedCredentials: true
        });
    }

    const method = usePostToCalTopo ? 'POST' : 'GET';
    const payloadString = ''; // Empty for since endpoint
    
    const {expires, signature} = signCalTopoRequest(method, endpoint, payloadString, creds.credentialSecret);

    try {
        console.log(`[PROXY] Fetching shapes from ${targetUrl} (Method: ${method})`);
        
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
        if (method === 'POST') {
            response = await axios.post(targetUrl, payloadString, axiosConfig);
        } else {
            response = await axios.get(targetUrl, axiosConfig);
        }

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
            req.body = { ...requestData, usePost: true };
            return fetchMapHandler(req, res);
        }

        const responseStatus = error.response ? error.response.status : 500;
        const responseBody = error.response && error.response.data ? error.response.data : null;
        
        if (responseStatus === 401 || (typeof responseBody === 'string' && responseBody.includes('Authentication'))) {
            const authMessage = `${method.toUpperCase()} ${endpoint}\n${expires}\n${payloadString}`;
            console.error(`[PROXY] Auth Failure! Method: ${method}, Endpoint: ${endpoint}, Expires: ${expires}`);
            console.error(`[PROXY] Signature: ${signature}`);
            console.error(`[PROXY] Signed Message:\n${authMessage}`);
            
            if (typeof responseBody === 'object' && responseBody !== null) {
                responseBody.proxyDiagnostics = {
                    method: method.toUpperCase(),
                    endpoint,
                    expires,
                    payloadSize: payloadString.length,
                    messageToSign: authMessage
                };
            }
        }

        const detailMessage = typeof responseBody === 'string'
            ? responseBody.slice(0, 400)
            : responseBody && responseBody.message
                ? responseBody.message
                : error.message;

        // Return detailed JSON for better debugging in the website
        res.status(responseStatus).json({
            error: error.response ? `CalTopo Error ${responseStatus}` : "Proxy Connection Error",
            message: detailMessage,
            targetUrl: targetUrl,
            mapId: trimmedMapId,
            signingRequired: true,
            credentialSource: creds.source,
            supportsClientSuppliedCredentials: true,
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

    // CalTopo Team API: if payload is an empty object, sign and send it as an empty string
    // Python's bool({}) is False, so json.dumps(payload) if payload else "" results in ""
    const payloadString = (payload && typeof payload === 'object' && Object.keys(payload).length > 0) 
        ? JSON.stringify(payload) 
        : (typeof payload === 'string' && payload.length > 0 ? payload : '');

    const { expires, signature } = signCalTopoRequest(method, endpoint, payloadString, creds.credentialSecret);

    try {
        console.log(`[PROXY] Generic ${method.toUpperCase()} to ${targetUrl}`);
        if (payloadString) console.log(`[PROXY] Payload size: ${payloadString.length} chars`);
        
        const axiosConfig = {
            timeout: CALTOPO_TIMEOUT_MS,
            params: {
                id: creds.credentialId,
                expires,
                signature
            },
            headers: {}
        };

        if (['POST', 'PUT', 'PATCH'].includes(method.toUpperCase())) {
            axiosConfig.headers['Content-Type'] = 'application/json';
        }

        let response;
        const upperMethod = method.toUpperCase();
        if (upperMethod === 'POST') {
            response = await axios.post(targetUrl, payloadString, axiosConfig);
        } else if (upperMethod === 'PUT') {
            response = await axios.put(targetUrl, payloadString, axiosConfig);
        } else if (upperMethod === 'DELETE') {
            // DELETE can have a body in some APIs, but let's stick to the signature
            response = await axios.delete(targetUrl, { ...axiosConfig, data: payloadString });
        } else {
            response = await axios.get(targetUrl, axiosConfig);
        }

        res.json(response.data);
    } catch (error) {
        console.error(`[PROXY] Error in generic call to ${targetUrl}:`, error.message);
        const status = error.response ? error.response.status : 500;
        const responseData = error.response ? error.response.data : { error: error.message };
        
        if (status === 401 || (typeof responseData === 'string' && responseData.includes('Authentication'))) {
            const authMessage = `${method.toUpperCase()} ${endpoint}\n${expires}\n${payloadString}`;
            console.error(`[PROXY] Auth Failure! Method: ${method}, Endpoint: ${endpoint}, Expires: ${expires}`);
            console.error(`[PROXY] Signature: ${signature}`);
            console.error(`[PROXY] Signed Message:\n${authMessage}`);
            
            // Add diagnostic info to help debug on the client
            if (typeof responseData === 'object') {
                responseData.proxyDiagnostics = {
                    method: method.toUpperCase(),
                    endpoint,
                    expires,
                    payloadSize: payloadString.length,
                    messageToSign: authMessage
                };
            }
        }
        
        res.status(status).json(typeof responseData === 'object' ? responseData : { error: responseData, message: responseData });
    }
};

app.get('/api/proxy', fetchMapHandler);
app.post('/api/proxy', fetchMapHandler);
app.post('/api/call', genericCallHandler);
app.get('/fetch-map', fetchMapHandler); // Alias for compatibility
app.post('/fetch-map', fetchMapHandler); // Alias for compatibility

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Sync server v1.3.0 listening on port ${PORT}`);
});
