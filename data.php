<?php
// data.php - PHP bridge for SAR database storage using SQLite
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Authorization, X-User-Name, X-User-Pin, X-Last-Modified");
header("Content-Type: application/json");

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    exit;
}

$db_file = __DIR__ . '/db/sar_sync.db';
if (!is_dir(__DIR__ . '/db')) {
    mkdir(__DIR__ . '/db', 0777, true);
}

try {
    $db = new PDO("sqlite:" . $db_file);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // Ensure table exists
    $db->exec("CREATE TABLE IF NOT EXISTS store (
        bucket TEXT,
        key TEXT,
        value TEXT,
        userName TEXT,
        userPin TEXT,
        updatedAt TEXT,
        PRIMARY KEY (bucket, key)
    )");

    $db->exec("CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT,
        pin TEXT
    )");
    
    $db->exec("CREATE TABLE IF NOT EXISTS user_settings (
        username TEXT PRIMARY KEY,
        settings TEXT DEFAULT '{}'
    )");
    
    $db->exec("CREATE TABLE IF NOT EXISTS bucket_history (
        username TEXT,
        bucket TEXT,
        lastAccessed TEXT,
        PRIMARY KEY (username, bucket)
    )");
} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(["error" => "Database connection failed: " . $e->getMessage()]);
    exit;
}

$method = $_SERVER['REQUEST_METHOD'];
$path_info = isset($_SERVER['PATH_INFO']) ? $_SERVER['PATH_INFO'] : (isset($_SERVER['ORIG_PATH_INFO']) ? $_SERVER['ORIG_PATH_INFO'] : '');
$parts = explode('/', trim($path_info, '/'));

if (count($parts) < 2 || $parts[0] !== 'api') {
    http_response_code(400);
    echo json_encode(["error" => "Invalid endpoint."]);
    exit;
}

if ($parts[1] === 'health' && $method === 'GET') {
    echo json_encode(["status" => "ok"]);
    exit;
}

if ($parts[1] === 'auth') {
    $action = isset($parts[2]) ? $parts[2] : '';
    
    $body = file_get_contents('php://input');
    $data = json_decode($body, true) ?: [];
    
    if ($action === 'register' && $method === 'POST') {
        $username = isset($data['username']) ? trim($data['username']) : '';
        $pin = isset($data['pin']) ? trim($data['pin']) : '';
        
        if (!$username || !$pin) {
            http_response_code(400); echo json_encode(["error" => "Username and PIN required"]); exit;
        }
        
        $hashed = hash('sha256', $pin);
        
        $stmt = $db->prepare("SELECT username FROM users WHERE username = ?");
        $stmt->execute([$username]);
        if ($stmt->fetch()) {
            http_response_code(400); echo json_encode(["error" => "User already exists"]); exit;
        }
        
        $stmt = $db->prepare("INSERT INTO users (username, password, pin) VALUES (?, ?, ?)");
        $stmt->execute([$username, $hashed, $pin]);
        echo json_encode(["success" => true, "user" => ["username" => $username, "pin" => $pin]]);
        exit;
    }
    
    if ($action === 'login' && $method === 'POST') {
        $username = isset($data['username']) ? trim($data['username']) : '';
        $pin = isset($data['pin']) ? trim($data['pin']) : '';
        
        if (!$username || !$pin) {
            http_response_code(400); echo json_encode(["error" => "Username and PIN required"]); exit;
        }

        $hashed = hash('sha256', $pin);
        $stmt = $db->prepare("SELECT * FROM users WHERE username = ? AND pin = ?");
        $stmt->execute([$username, $pin]);
        $row = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($row) {
            echo json_encode(["success" => true, "user" => ["username" => $row['username'], "pin" => $row['pin']]]);
        } else {
            http_response_code(401); echo json_encode(["error" => "no matching login found"]);
        }
        exit;
    }
}

// Global Auth Check for all other endpoints
$username = isset($_SERVER['HTTP_X_USER_NAME']) ? $_SERVER['HTTP_X_USER_NAME'] : '';
$password = isset($_SERVER['HTTP_X_USER_PASSWORD']) ? $_SERVER['HTTP_X_USER_PASSWORD'] : (isset($_SERVER['HTTP_X_USER_PIN']) ? $_SERVER['HTTP_X_USER_PIN'] : '');

if (!$username || !$password) {
    http_response_code(401); echo json_encode(["error" => "Not authenticated"]); exit;
}
$hashed = hash('sha256', $password);
$stmt = $db->prepare("SELECT * FROM users WHERE username = ? AND password = ?");
$stmt->execute([$username, $hashed]);
if (!$stmt->fetch()) {
    http_response_code(401); echo json_encode(["error" => "Invalid credentials"]); exit;
}

if ($parts[1] === 'auth') {
    $action = isset($parts[2]) ? $parts[2] : '';
    $body = file_get_contents('php://input');
    $data = json_decode($body, true) ?: [];
    
    if ($action === 'settings') {
        if ($method === 'GET') {
            $stmt = $db->prepare("SELECT settings FROM user_settings WHERE username = ?");
            $stmt->execute([$username]);
            $row = $stmt->fetch(PDO::FETCH_ASSOC);
            echo $row ? $row['settings'] : '{}';
            exit;
        } elseif ($method === 'PUT') {
            $settings = json_encode($data);
            $stmt = $db->prepare("INSERT OR REPLACE INTO user_settings (username, settings) VALUES (?, ?)");
            $stmt->execute([$username, $settings]);
            echo json_encode(["success" => true]);
            exit;
        }
    }
    
    if ($action === 'history' && $method === 'GET') {
        $stmt = $db->prepare("SELECT bucket, lastAccessed FROM bucket_history WHERE username = ? ORDER BY lastAccessed DESC LIMIT 5");
        $stmt->execute([$username]);
        echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        exit;
    }
    
    http_response_code(404);
    echo json_encode(["error" => "Not found"]);
    exit;
}

// Expected path: /api/v1/:bucket/:key
if ($parts[1] !== 'v1' || count($parts) < 3) {
    http_response_code(400);
    echo json_encode(["error" => "Invalid endpoint. Use /api/v1/:bucket/:key"]);
    exit;
}

$bucket = $parts[2];
$key = isset($parts[3]) ? $parts[3] : null;

function trackBucketAccess($db, $bucket) {
    $username = isset($_SERVER['HTTP_X_USER_NAME']) ? $_SERVER['HTTP_X_USER_NAME'] : '';
    if ($username && $bucket) {
        $stmt = $db->prepare("INSERT OR REPLACE INTO bucket_history (username, bucket, lastAccessed) VALUES (?, ?, ?)");
        $stmt->execute([$username, $bucket, date('c')]);
    }
}

if ($method === 'GET') {
    trackBucketAccess($db, $bucket);
    if ($key === 'latest') {
        $stmt = $db->prepare("SELECT value FROM store WHERE bucket = ? ORDER BY updatedAt DESC LIMIT 1");
        $stmt->execute([$bucket]);
        $row = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($row) {
            echo $row['value'];
        } else {
            http_response_code(404);
            echo json_encode(["error" => "No data found"]);
        }
    } elseif ($key === 'all-files') {
        $stmt = $db->prepare("SELECT key, updatedAt FROM store WHERE bucket = ?");
        $stmt->execute([$bucket]);
        $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $files = [];
        foreach ($rows as $row) {
            $files[$row['key']] = ["lastModified" => $row['updatedAt']];
        }
        echo json_encode($files);
    } elseif ($key) {
        $stmt = $db->prepare("SELECT value FROM store WHERE bucket = ? AND key = ?");
        $stmt->execute([$bucket, $key]);
        $row = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($row) {
            echo $row['value'];
        } else {
            http_response_code(404);
            echo json_encode(["error" => "Not found"]);
        }
    } else {
        $stmt = $db->prepare("SELECT key FROM store WHERE bucket = ?");
        $stmt->execute([$bucket]);
        echo json_encode($stmt->fetchAll(PDO::FETCH_COLUMN));
    }
} elseif ($method === 'PUT') {
    trackBucketAccess($db, $bucket);
    if (!$key) {
        http_response_code(400);
        echo json_encode(["error" => "Key required for PUT"]);
        exit;
    }
    
    $body = file_get_contents('php://input');
    $data = json_decode($body, true);
    
    $userName = isset($_SERVER['HTTP_X_USER_NAME']) ? $_SERVER['HTTP_X_USER_NAME'] : 'Unknown';
    $userPin = isset($_SERVER['HTTP_X_USER_PIN']) ? $_SERVER['HTTP_X_USER_PIN'] : (isset($_SERVER['HTTP_X_USER_PASSWORD']) ? $_SERVER['HTTP_X_USER_PASSWORD'] : '');
    $isSuperAdmin = ($userPin === '1976');
    
    $incomingLastModified = time() * 1000;
    if (isset($_SERVER['HTTP_X_LAST_MODIFIED'])) {
        $incomingLastModified = strtotime($_SERVER['HTTP_X_LAST_MODIFIED']) * 1000;
    } elseif (isset($data['lastModified'])) {
        $incomingLastModified = strtotime($data['lastModified']) * 1000;
    }
    
    // Check for existing record
    $stmt = $db->prepare("SELECT userPin, updatedAt FROM store WHERE bucket = ? AND key = ?");
    $stmt->execute([$bucket, $key]);
    $row = $stmt->fetch(PDO::FETCH_ASSOC);
    
    if ($row) {
        $currentIsSuperAdmin = ($row['userPin'] === '1976');
        $existingLastModified = strtotime($row['updatedAt']) * 1000;
        
        if ($currentIsSuperAdmin && !$isSuperAdmin) {
            http_response_code(403);
            echo json_encode(["error" => "Conflict", "message" => "Changes by Super-Admin cannot be overwritten."]);
            exit;
        }
        
        if ($isSuperAdmin === $currentIsSuperAdmin) {
            if ($incomingLastModified < $existingLastModified) {
                http_response_code(403);
                echo json_encode(["error" => "Conflict", "message" => "Incoming data is older than server data."]);
                exit;
            }
        }
    }
    
    $saveTime = date('c', $incomingLastModified / 1000);
    
    $stmt = $db->prepare("INSERT OR REPLACE INTO store (bucket, key, value, userName, userPin, updatedAt) VALUES (?, ?, ?, ?, ?, ?)");
    $stmt->execute([$bucket, $key, $body, $userName, $userPin, $saveTime]);

    echo json_encode(["success" => true]);
} elseif ($method === 'DELETE') {
    trackBucketAccess($db, $bucket);
    if (!$key) {
        http_response_code(400);
        echo json_encode(["error" => "Key required for DELETE"]);
        exit;
    }
    
    $userPin = isset($_SERVER['HTTP_X_USER_PIN']) ? $_SERVER['HTTP_X_USER_PIN'] : (isset($_SERVER['HTTP_X_USER_PASSWORD']) ? $_SERVER['HTTP_X_USER_PASSWORD'] : '');
    $isSuperAdmin = ($userPin === '1976');
    
    $stmt = $db->prepare("SELECT userPin FROM store WHERE bucket = ? AND key = ?");
    $stmt->execute([$bucket, $key]);
    $row = $stmt->fetch(PDO::FETCH_ASSOC);
    
    if ($row && $row['userPin'] === '1976' && !$isSuperAdmin) {
        http_response_code(403);
        echo json_encode(["error" => "Conflict", "message" => "Cannot delete Super-Admin created files."]);
        exit;
    }
    
    $stmt = $db->prepare("DELETE FROM store WHERE bucket = ? AND key = ?");
    $stmt->execute([$bucket, $key]);
    
    echo json_encode(["success" => true]);
}
