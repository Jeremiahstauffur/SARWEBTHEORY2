const express = require('express');
const axios = require('axios');
const cors = require('cors');
const app = express();

app.use(cors()); // This tells the browser: "It's okay to let my website talk to me!"

// 1. Health endpoint (used by the website to check if the proxy is alive)
app.get('/api/health', (req, res) => {
    res.json({status: 'ok', message: 'Proxy is live and well'});
});

// 2. Fetch Map endpoint (now handles mapId and domain)
app.get('/fetch-map', async (req, res) => {
    const {mapId, domain} = req.query;

    if (!mapId) {
        return res.status(400).send("Missing mapId parameter. Use /fetch-map?mapId=YOUR_MAP_ID&domain=caltopo.com");
    }

    const targetDomain = domain || 'caltopo.com';
    const targetUrl = `https://${targetDomain}/api/v1/map/${mapId}/features`;

    try {
        console.log(`Fetching ${targetUrl} on behalf of the website...`);
        // We are asking SARTopo for the data on BEHALF of your website
        const response = await axios.get(targetUrl);
        res.json(response.data);
    } catch (error) {
        console.error(`Error fetching from ${targetUrl}:`, error.message);
        res.status(500).send("The waiter couldn't get the food: " + error.message);
    }
});

const PORT = process.env.PORT || 3000; // Railway tells your app which port to use
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Proxy is live on port ${PORT}`);
});