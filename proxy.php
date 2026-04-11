<?php
/**
 * SAR CalTopo Proxy (PHP Version)
 * Use this file if your hosting environment does not support Node.js/Express.
 *
 * Supports:
 * 1. Standard CalTopo Maps (via mapId)
 * 2. Team Accounts (via mapId and teamId)
 * 3. Fallback for old map IDs (automatic "S" prefix)
 *
 * To use:
 * 1. Upload this file to your web server (e.g., as proxy.php)
 * 2. In your SAR website Maps Management, set the Map ID and Team ID (if applicable).
 */

// Enable CORS
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Authorization, X-CalTopo-Account-Id, X-CalTopo-Credential-Id, X-CalTopo-Secret");

// Handle preflight OPTIONS request
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    exit;
}

// 1. Health Check Handler
// Allows the website to verify if the proxy is online
if (isset($_GET['health']) || (isset($_SERVER['PATH_INFO']) && $_SERVER['PATH_INFO'] === '/api/health')) {
    header("Content-Type: application/json");
    echo json_encode([
        "status" => "ok",
        "message" => "PHP Proxy is live and well",
        "version" => "1.2.0",
        "timestamp" => date('c')
    ]);
    exit;
}

// 2. Fetch Map Handler
$mapId = isset($_GET['mapId']) ? $_GET['mapId'] : '';
$teamId = isset($_GET['teamId']) ? $_GET['teamId'] : '';
$domain = (isset($_GET['domain']) && $_GET['domain'] !== 'undefined') ? $_GET['domain'] : 'caltopo.com';

// Clean up mapId (remove any trailing slashes or spaces that might cause 404)
$mapId = trim($mapId, "/ \t\n\r\0\x0B");
$teamId = trim($teamId, "/ \t\n\r\0\x0B");

if (empty($mapId)) {
    header("Content-Type: application/json");
    http_response_code(400);
    echo json_encode([
        "error" => "Missing mapId parameter",
        "message" => "Please ensure your Map ID is correctly entered in the Maps page (e.g., ABCDE)."
    ]);
    exit;
}

/**
 * Helper to sign a request
 */
function signRequest($method, $url, $payload, $credentialId, $secret)
{
    $expires = (time() + 30) * 1000; // 30 seconds
    $stringToSign = "$method $url\n$expires\n$payload";
    $signature = base64_encode(hash_hmac('sha256', $stringToSign, base64_decode($secret), true));
    
    return $url . (strpos($url, '?') === false ? '?' : '&') . "id=$credentialId&expires=$expires&signature=" . urlencode($signature);
}

/**
 * Helper to perform the cURL request
 */
function performRequest($url, $method = 'GET', $payload = '', $credentialId = '', $secret = '')
{
    if ($credentialId && $secret) {
        $url = signRequest($method, $url, $payload, $credentialId, $secret);
    }
    
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
    curl_setopt($ch, CURLOPT_TIMEOUT, 15);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, true);
    if ($method === 'POST') {
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $payload);
    }

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error = curl_error($ch);
    curl_close($ch);

    return [
        'response' => $response,
        'httpCode' => $httpCode,
        'error' => $error,
        'url' => $url
    ];
}

$credentialId = isset($_SERVER['HTTP_X_CALTOPO_CREDENTIAL_ID']) ? $_SERVER['HTTP_X_CALTOPO_CREDENTIAL_ID'] : '';
$secret = isset($_SERVER['HTTP_X_CALTOPO_SECRET']) ? $_SERVER['HTTP_X_CALTOPO_SECRET'] : '';

// Construct target URL
if (!empty($teamId)) {
    // Team Account Workspace endpoint
    $targetUrl = "https://caltopo.com/api/v1/acct/{$teamId}/CollaborativeMap/{$mapId}/features";
} else {
    // Standard Map Access endpoint
    $targetUrl = "https://{$domain}/api/v1/map/{$mapId}/features";
}

$result = performRequest($targetUrl, 'GET', '', $credentialId, $secret);

// 3. Fallback logic for "S" prefix (Common for old map IDs)
if ($result['httpCode'] === 404 && empty($teamId) && strpos($mapId, 'S') !== 0 && strlen($mapId) <= 5) {
    $sMapId = 'S' . $mapId;
    $sTargetUrl = "https://caltopo.com/api/v1/map/{$sMapId}/features";
    $sResult = performRequest($sTargetUrl, 'GET', '', $credentialId, $secret);

    if ($sResult['httpCode'] === 200) {
        $result = $sResult;
        $mapId = $sMapId; // Update for error messages
    }
}

$response = $result['response'];
$httpCode = $result['httpCode'];
$error = $result['error'];
$targetUrl = $result['url'];

header("Content-Type: application/json");

// Handle Connection Errors
if ($response === false) {
    http_response_code(500);
    echo json_encode([
        "error" => "Proxy Connection Error",
        "message" => "The PHP proxy could not connect to CalTopo. Details: " . $error,
        "targetUrl" => $targetUrl,
        "mapId" => $mapId
    ]);
    exit;
}

// Handle HTTP response codes from CalTopo
http_response_code($httpCode);

if ($httpCode === 404) {
    echo json_encode([
        "error" => "Map Not Found",
        "message" => "CalTopo couldn't find map ID \"$mapId\". If this is a team map, ensure you entered the Team ID. Also try prepending 'S' to the ID if it's an old map.",
        "targetUrl" => $targetUrl,
        "mapId" => $mapId,
        "teamId" => $teamId
    ]);
} elseif ($httpCode === 403 || $httpCode === 401) {
    echo json_encode([
        "error" => "Private Map (Forbidden)",
        "message" => "CalTopo refused to share this map. This usually means the map is private. Please go to 'Map Settings' in CalTopo and set 'Access' to 'Anyone with the link can view'.",
        "targetUrl" => $targetUrl,
        "mapId" => $mapId
    ]);
} elseif ($httpCode >= 400) {
    echo json_encode([
        "error" => "CalTopo Error $httpCode",
        "message" => "The request to CalTopo failed with status $httpCode.",
        "targetUrl" => $targetUrl,
        "mapId" => $mapId,
        "response_preview" => substr($response, 0, 200)
    ]);
} else {
    // Success - pass through the CalTopo response
    echo $response;
}
