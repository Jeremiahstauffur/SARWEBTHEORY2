const axios = require('axios');
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'db');

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
    res.json({
        status: 'ok',
        version: '1.1.0',
        service: 'SAR Proxy + Sync',
        message: 'Unified server is live and well',
        timestamp: new Date().toISOString()
    });
});

// CalTopo Proxy endpoint
const fetchMapHandler = async (req, res) => {
    const {mapId, domain} = req.query;
    if (!mapId) {
        return res.status(400).json({
            error: "Missing mapId parameter",
            message: "Please ensure your Map ID is correctly entered in the Maps page."
        });
    }

    const targetDomain = domain && domain !== 'undefined' ? domain : 'caltopo.com';
    const targetUrl = `https://${targetDomain}/api/v1/map/${mapId}/features`;

    try {
        console.log(`[PROXY] Fetching shapes from ${targetUrl}`);
        const response = await axios.get(targetUrl, {timeout: 15000});
        res.json(response.data);
    } catch (error) {
        console.error(`[PROXY] Error fetching from ${targetUrl}:`, error.message);

        const responseStatus = error.response ? error.response.status : 500;

        // Return detailed JSON for better debugging in the website
        res.status(responseStatus).json({
            error: error.response ? `CalTopo Error ${responseStatus}` : "Proxy Connection Error",
            message: error.message,
            targetUrl: targetUrl,
            mapId: mapId
        });
    }
};

app.get('/api/proxy', fetchMapHandler);
app.get('/fetch-map', fetchMapHandler); // Alias for compatibility

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Sync server v1.1.0 listening on port ${PORT}`);
});
