const axios = require('axios');
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5050;
const DATA_DIR = path.join(__dirname, 'db');

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR);
}

app.use(cors());
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

    try {
        fs.writeFileSync(filePath, JSON.stringify(req.body, null, 2));
        res.json({success: true});
    } catch (err) {
        res.status(500).json({error: 'Failed to save data'});
    }
});

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

// SarTopo Proxy endpoint
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
            error: error.response ? `SARTopo Error ${responseStatus}` : "Proxy Connection Error",
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
