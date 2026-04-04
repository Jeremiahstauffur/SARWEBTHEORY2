const ROWS = 1;
const COLS = 5;
const DEFAULT_DYNAMIC_COLS = 3;
const BUNDLE_STORAGE_KEY = 'pill-table-bundle-v1';
const FILE_LIST_STORAGE_KEY = 'sar-saved-files-v1';
const DEFAULT_FILE_NAME = 'us-pill-data.json';
const SYNC_URL_STORAGE_KEY = 'sar-sync-url-v1';
const SYNC_BUCKET_STORAGE_KEY = 'sar-sync-bucket-v1';
const SARTOPO_PROXY_STORAGE_KEY = 'sar-sartopo-proxy-v1';
const DEVICE_ID_STORAGE_KEY = 'sar-device-id-v1';
const KVDB_BASE_URL = 'https://kvdb.io';
const DEFAULT_BUCKET = 'sar-sync-' + Math.random().toString(36).substr(2, 6);

function getSyncBucket() {
    let bucket = localStorage.getItem(SYNC_BUCKET_STORAGE_KEY);
    if (!bucket) {
        // If they had an old sync URL, maybe we can derive a bucket from it? 
        // But better to just start fresh or let them set it.
        bucket = DEFAULT_BUCKET;
        localStorage.setItem(SYNC_BUCKET_STORAGE_KEY, bucket);
    }
    return bucket;
}

function getSartopoProxy() {
    return localStorage.getItem(SARTOPO_PROXY_STORAGE_KEY) || 'http://localhost:5050/api/proxy';
}

function setSartopoProxy(url) {
    if (url) localStorage.setItem(SARTOPO_PROXY_STORAGE_KEY, url);
    else localStorage.removeItem(SARTOPO_PROXY_STORAGE_KEY);
}

function getDeviceId() {
    let id = localStorage.getItem(DEVICE_ID_STORAGE_KEY);
    if (!id) {
        id = 'device-' + Math.random().toString(36).substr(2, 9) + '-' + Date.now();
        localStorage.setItem(DEVICE_ID_STORAGE_KEY, id);
    }
    return id;
}

const PAGE_DEFS = [
  { key: 'index', title: 'Regions', href: 'index.html' },
  { key: 'page2', title: 'Segments', href: 'page2.html' },
  { key: 'page3', title: 'Personnel', href: 'page3.html' },
  { key: 'page4', title: 'Search Log', href: 'page4.html' },
  { key: 'page5', title: 'Forms', href: 'page5.html' },
  { key: 'page6', title: 'Incident', href: 'page6.html' },
  { key: 'page7', title: 'Uploads', href: 'page7.html' },
  { key: 'page8', title: 'User Account', href: 'page8.html' },
  { key: 'page10', title: 'Maps', href: 'page10.html' },
  { key: 'settings', title: 'Settings', href: 'settings.html' }
];

const BRAND_NAME = 'Search & Rescue Theory Software';

let newlyImportedSegments = new Set();

const HIGHLIGHT_COLORS = {
    orange: '#ffa500',
    yellow: '#ffff00',
    red: '#ff0000',
    blue: '#0000ff',
    green: '#008000',
    purple: '#800080',
    brown: '#a52a2a',
    black: '#000000',
    white: '#ffffff',
    grey: '#808080',
    maroon: '#800000'
};

function pageKey() {
  return document.body.dataset.page || 'index';
}

function isHomePage() {
  return pageKey() === 'home';
}

function isRegionsPage() {
  return pageKey() === 'index';
}

function isSettingsPage() {
  return pageKey() === 'settings';
}

function isSegmentsPage() {
  const pk = pageKey();
  return pk === 'page2' || pk === 'index';
}

function isPersonnelPage() {
  return pageKey() === 'page3';
}

function isSearchLogPage() {
  return pageKey() === 'page4';
}

function isUploadsPage() {
  const pk = pageKey();
  return pk === 'page7' || pk === 'uploads';
}

function isFormsPage() {
  return pageKey() === 'page5';
}

function isProfilePage() {
  return pageKey() === 'page6';
}

function isPage8() {
  return pageKey() === 'page8';
}

function isPage9() {
  return false;
}

function isMapsPage() {
  return pageKey() === 'page10';
}

function isAccountsPage() {
  return pageKey() === 'page8';
}

function getCurrentUser() {
  const userJson = sessionStorage.getItem('sar-current-user');
  if (!userJson) return null;
  return JSON.parse(userJson);
}

function isUserAdmin(user) {
    return !!user;
}

function getAccountName(user) {
    if (!user) return '';
    if (user.pin === '1976') return 'Super-Admin';
    return (user.firstName + (user.lastName ? ' ' + (user.lastName || '') : '')).trim();
}

function setupAutoFormatDate(input) {
  input.oninput = () => {
    let val = input.value.replace(/[^\d]/g, '');
    if (val.length > 8) val = val.slice(0, 8);
    let formatted = val;
    if (val.length > 4) {
      formatted = val.slice(0, 2) + '-' + val.slice(2, 4) + '-' + val.slice(4);
    } else if (val.length > 2) {
      formatted = val.slice(0, 2) + '-' + val.slice(2);
    }
    input.value = formatted;
  };
}

function setupAutoFormatTime(input) {
  input.oninput = () => {
    let val = input.value.replace(/[^\d]/g, '');
    if (val.length > 4) val = val.slice(0, 4);
    let formatted = val;
    if (val.length > 2) {
      formatted = val.slice(0, 2) + ':' + val.slice(2);
    }
    input.value = formatted;
  };
}

function showTimePrompt(title, onConfirm, onCancel) {
    const popup = createPopup(title, null, onCancel);
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');

    const inputs = document.createElement('div');
    inputs.className = 'popup-input-container';

    const now = new Date();
    const defaultDate = `${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}-${now.getFullYear()}`;
    const defaultTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;

    const dateInput = document.createElement('input');
    dateInput.className = 'pill-input';
    dateInput.placeholder = 'MM-DD-YYYY';
    dateInput.value = defaultDate;
    setupAutoFormatDate(dateInput);
    inputs.appendChild(dateInput);

    const timeInput = document.createElement('input');
    timeInput.className = 'pill-input';
    timeInput.placeholder = 'hh:mm';
    timeInput.value = defaultTime;
    setupAutoFormatTime(timeInput);
    inputs.appendChild(timeInput);

    content.insertBefore(inputs, btnContainer);

    const updateBtn = document.createElement('button');
    updateBtn.className = 'popup-btn primary';
    updateBtn.textContent = 'Update';
    updateBtn.onclick = () => {
        onConfirm(dateInput.value, timeInput.value);
        popup.remove();
    };
    btnContainer.appendChild(updateBtn);
}

function setCurrentUser(user) {
  if (user) {
    sessionStorage.setItem('sar-current-user', JSON.stringify(user));
    notifyActiveUser(user);
  } else {
    sessionStorage.removeItem('sar-current-user');
  }
}

function checkAccess() {
  const user = getCurrentUser();
  const page = pageKey();
  const bundle = loadBundle();

  if (!user) {
    if (page !== 'index') navigateToPage('index.html');
    return;
  }

  // Refresh user data from bundle to ensure visiblePages are up to date
  const actualUser = (bundle.accounts || []).find(a => a.pin === user.pin);
  if (actualUser) {
      setCurrentUser(actualUser);
      if (isUserAdmin(actualUser)) return; // Admin has access to everything
  }

  if (page === 'page9') {
      // Everyone is an admin now
      return;
  }

    if (actualUser && actualUser.visiblePages) {
        // Everyone is allowed access to everything now
        return;
    }
}

function defaultSearchLogData() {
  return Array.from({ length: ROWS }, () =>
    Array.from({ length: 10 }, () => '')
  );
}

function defaultPersonnelData() {
  return Array.from({ length: ROWS }, () =>
    Array.from({ length: 7 }, () => '')
  );
}

function defaultSegmentsData() {
  return Array.from({ length: ROWS }, () =>
    Array.from({ length: 9 }, () => '')
  );
}

function defaultData() {
  return Array.from({ length: ROWS }, () =>
    Array.from({ length: COLS }, () => '')
  );
}

function defaultRegionsData() {
  return {
    headers: ['Region', ...Array.from({ length: DEFAULT_DYNAMIC_COLS }, (_, i) => `Voter ${i + 1}`), 'Consensus'],
    rows: Array.from({ length: ROWS }, () => Array.from({ length: DEFAULT_DYNAMIC_COLS + 2 }, () => ''))
  };
}

function defaultBundle() {
  return {
    fileName: DEFAULT_FILE_NAME,
    deleteMode: false,
    theme: 'dark',
    showTips: true,
    background: 'assets/us-night.jpg',
    activityLog: [],
    currentAssignments: {},
    teamStatuses: {},
    parChecks: {},
    teamLeaveTimes: {},
    teamAssignmentTimes: {},
    parCheckFrequency: 20,
    dismissedNotifications: [],
    arrivedTeams: [],
    forms: {},
    uploads: [],
    maps: [],
    accounts: [
      { firstName: 'Super', lastName: 'Admin', pin: '1976', color: 'none', handle: 'Super-Admin', isFileManager: true, theme: 'dark', visiblePages: ['index', 'page2', 'page3', 'page4', 'page5', 'page6', 'page7', 'settings', 'home', 'page8', 'page9', 'page10'] }
    ],
    profile: {
      incidentNumber: '',
      lostPersonName: '',
      lostPersonAge: '',
      lostPersonGender: '',
      lostPersonDescription: '',
      lostPersonClothing: '',
      lostPersonPhysical: ''
    },
    pages: {
      index: defaultRegionsData(),
      page2: defaultSegmentsData(),
      page3: defaultPersonnelData(),
      page4: defaultSearchLogData(),
      page5: defaultData(),
      page6: defaultData(),
      page7: defaultData()
    }
  };
}

function sanitizeStandardData(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return defaultData();
  return parsed.map(row =>
    Array.from({ length: COLS }, (_, c) => (row?.[c] ?? '').toString())
  );
}

function sanitizeRegionsData(parsed) {
  const fallback = defaultRegionsData();
  if (!parsed || typeof parsed !== 'object') return fallback;

  const headers = Array.isArray(parsed.headers) ? parsed.headers.slice() : fallback.headers.slice();
  const rawRows = Array.isArray(parsed.rows) ? parsed.rows : fallback.rows;
  const rowCount = Math.max(1, rawRows.length);
  const dynamicCount = Math.max(1, headers.length - 2 || DEFAULT_DYNAMIC_COLS);

  const safeHeaders = ['Region'];
  for (let i = 0; i < dynamicCount; i++) {
    safeHeaders.push((headers[i + 1] ?? `Voter ${i + 1}`).toString());
  }
  safeHeaders.push('Consensus');

  const safeRows = Array.from({ length: rowCount }, (_, rowIndex) => {
    const sourceRow = Array.isArray(rawRows[rowIndex]) ? rawRows[rowIndex] : [];
    // Ensure row has enough columns (Region + dynamicCount + Consensus)
    return Array.from({ length: dynamicCount + 2 }, (_, colIndex) => (sourceRow[colIndex] ?? '').toString());
  });

  return { headers: safeHeaders, rows: safeRows };
}

function sanitizeSegmentsData(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return defaultSegmentsData();
  return parsed.map(row => {
    const sourceLen = row.length;
    const targetRow = Array.from({ length: 9 }, (_, c) => (row?.[c] ?? '').toString());
    if (sourceLen === 7) {
       // Old index 6 was PSR. Now it's PSRi and we initialize PSRc with it.
       targetRow[7] = targetRow[6];
    }
    return targetRow;
  });
}

function sanitizePersonnelData(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return defaultPersonnelData();
  return parsed.map(row => {
    // Keep at least 9 columns to preserve PIN link at index 8
    const r = Array.from({ length: Math.max(7, row?.length || 0) }, (_, c) => (row?.[c] ?? '').toString());
    // Clear team/lead if off-scene
    if (r[6] === 'false' || r[6] === '') {
      r[1] = '';
      r[2] = '';
    }
    return r;
  });
}

function sanitizeSearchLogData(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return defaultSearchLogData();
  return parsed.map(row =>
    Array.from({ length: 10 }, (_, c) => (row?.[c] ?? '').toString())
  );
}

function sanitizeBundle(bundle) {
  const fallback = defaultBundle();
  if (!bundle || typeof bundle !== 'object') return fallback;

  const fileName = typeof bundle.fileName === 'string' && bundle.fileName.trim()
    ? bundle.fileName.trim()
    : DEFAULT_FILE_NAME;

  const deleteMode = typeof bundle.deleteMode === 'boolean' ? bundle.deleteMode : false;
  const background = typeof bundle.background === 'string' && bundle.background.trim()
    ? bundle.background.trim()
    : 'assets/us-night.jpg';

  const activityLog = Array.isArray(bundle.activityLog) ? bundle.activityLog : [];
  const currentAssignments = (bundle.currentAssignments && typeof bundle.currentAssignments === 'object')
    ? bundle.currentAssignments
    : {};

  const teamStatuses = (bundle.teamStatuses && typeof bundle.teamStatuses === 'object')
    ? bundle.teamStatuses
    : {};

  const parChecks = (bundle.parChecks && typeof bundle.parChecks === 'object')
    ? bundle.parChecks
    : {};

  const teamLeaveTimes = (bundle.teamLeaveTimes && typeof bundle.teamLeaveTimes === 'object')
    ? bundle.teamLeaveTimes
    : {};

  const teamAssignmentTimes = (bundle.teamAssignmentTimes && typeof bundle.teamAssignmentTimes === 'object')
    ? bundle.teamAssignmentTimes
    : {};

  const arrivedTeams = Array.isArray(bundle.arrivedTeams) ? bundle.arrivedTeams : [];
  const dismissedNotifications = Array.isArray(bundle.dismissedNotifications) ? bundle.dismissedNotifications : [];

  const pages = {};
  for (const page of PAGE_DEFS) {
    const rawPage = bundle.pages?.[page.key];
    if (page.key === 'index') {
      pages[page.key] = sanitizeRegionsData(rawPage);
    } else if (page.key === 'page2') {
      pages[page.key] = sanitizeSegmentsData(rawPage);
    } else if (page.key === 'page3') {
      pages[page.key] = sanitizePersonnelData(rawPage);
    } else if (page.key === 'page4') {
      pages[page.key] = sanitizeSearchLogData(rawPage);
    } else {
      pages[page.key] = sanitizeStandardData(rawPage);
    }
  }

  // Ensure all teams in Personnel (page3) have a status of "at base" if not already set
  const personnelData = pages.page3 || [];
  personnelData.forEach(row => {
    const team = (row[1] || '').trim();
    const onScene = row[6] === 'true';
    if (team && onScene && !teamStatuses[team]) {
        const now = new Date();
        const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
        teamStatuses[team] = `at base (${timeStr})`;
    }
  });

  const parCheckFrequency = (bundle.parCheckFrequency !== undefined) ? bundle.parCheckFrequency : 20;
  const showTips = (bundle.showTips !== undefined) ? bundle.showTips : true;
  const theme = bundle.theme || 'dark';
  const forms = bundle.forms || {};
  const profile = bundle.profile || fallback.profile;
  const uploads = Array.isArray(bundle.uploads) ? bundle.uploads : [];
  const maps = Array.isArray(bundle.maps) ? bundle.maps : [];

  let accounts = Array.isArray(bundle.accounts) ? bundle.accounts : fallback.accounts;

  // 1. Ensure only one Super Admin exists
  const superAdmin = accounts.find(a => a.pin === '1976') || fallback.accounts[0];
  superAdmin.firstName = 'Super';
  superAdmin.lastName = 'Admin';
  superAdmin.handle = 'Super-Admin';
  superAdmin.pin = '1976';
  
  if (superAdmin.visiblePages && !superAdmin.visiblePages.includes('page10')) {
    superAdmin.visiblePages.push('page10');
  }

  // 2. Filter out any other accounts that might be pretending to be super admin
  accounts = accounts.filter(a => a.pin !== '1976');
  
  // 3. Sync with Personnel (page3)
  const personnel = pages.page3 || [];

  const syncedAccounts = [];
  syncedAccounts.push(superAdmin);

  // Helper to find next available PIN starting from 1400
  const getNextPin = (currentSynced) => {
    let next = 1400;
    while (currentSynced.some(a => a.pin === next.toString())) {
      next++;
    }
    return next.toString();
  };

  personnel.forEach(row => {
    const name = (row[0] || '').trim();
    if (!name) return;

    // Try to find by PIN link first (column 8)
    const rowPin = (row[8] || '').trim();
    let existing = null;
    
    if (rowPin) {
        existing = accounts.find(a => a.pin === rowPin);
    }

    // Fallback to name match
    if (!existing) {
        existing = accounts.find(a => a.handle === name || (a.firstName + ' ' + (a.lastName || '')).trim() === name);
    }

    if (existing) {
      // Update account name from Personnel list (Personnel is source of truth for name unless changed via User Account page which also updates Personnel)
      const parts = name.split(' ');
      existing.firstName = parts[0];
      existing.lastName = parts.slice(1).join(' ');
      existing.handle = name;
      
      // Ensure PIN link is established in Personnel list
      row[8] = existing.pin;
      
      syncedAccounts.push(existing);
      // Remove from accounts to avoid double-processing
      accounts = accounts.filter(a => a !== existing);
    } else {
      // Create new account
      const newPin = rowPin || getNextPin(syncedAccounts);
      const parts = name.split(' ');
      const newAcc = {
        firstName: parts[0],
        lastName: parts.slice(1).join(' '),
        pin: newPin,
        color: 'none',
        handle: name,
        isFileManager: false,
        theme: 'dark',
        visiblePages: ['index', 'page2', 'page3', 'page4', 'page5', 'page6', 'page7', 'settings', 'home', 'page8', 'page10']
      };
      syncedAccounts.push(newAcc);
      row[8] = newPin; // Establish link
    }
  });

  // Also keep any other accounts that are Admins or were not matched (to prevent accidental loss)
  // This prevents accidental loss of the currently logged-in user or admins not yet in Personnel list
  const currentU = getCurrentUser();
  accounts.forEach(a => {
      const isAlreadySynced = syncedAccounts.some(sa => sa.pin === a.pin);
      const isCurrUser = currentU && a.pin === currentU.pin;
      const isAdmin = isUserAdmin(a) || a.isFileManager;

      if (!isAlreadySynced && (isAdmin || isCurrUser)) {
          syncedAccounts.push(a);
      }
  });

  return { 
    fileName, 
    deleteMode, 
    theme, 
    background, 
    activityLog, 
    currentAssignments, 
    teamStatuses, 
    parChecks, 
    teamLeaveTimes, 
    teamAssignmentTimes, 
    arrivedTeams, 
    dismissedNotifications,
    parCheckFrequency, 
    showTips, 
    pages, 
    forms, 
    profile, 
    uploads, 
    maps,
    accounts: syncedAccounts 
  };
}

function loadBundle() {
  try {
    const raw = localStorage.getItem(BUNDLE_STORAGE_KEY);
    if (!raw) return defaultBundle();
    return sanitizeBundle(JSON.parse(raw));
  } catch {
    return defaultBundle();
  }
}

function saveBundle(bundle) {
  const sanitized = sanitizeBundle(bundle);
  const oldBundle = loadBundle();
  const oldName = oldBundle.fileName;
  const newName = sanitized.fileName;

  localStorage.setItem(BUNDLE_STORAGE_KEY, JSON.stringify(sanitized));
  pushBundleToServer(sanitized);
  
  // Also update in the list if it's there
  const files = getSavedFiles();
  if (oldName && files[oldName] && oldName !== newName) {
      // Rename case
      files[newName] = files[oldName];
      files[newName].bundle = sanitized;
      files[newName].lastModified = new Date().toLocaleString();
      delete files[oldName];
      localStorage.setItem(FILE_LIST_STORAGE_KEY, JSON.stringify(files));
  } else if (files[newName]) {
      // Update case
      files[newName].bundle = sanitized;
      files[newName].lastModified = new Date().toLocaleString();
      localStorage.setItem(FILE_LIST_STORAGE_KEY, JSON.stringify(files));
  }
  updateFileNameDisplay();
}

function getSavedFiles() {
    const raw = localStorage.getItem(FILE_LIST_STORAGE_KEY);
    if (!raw) return {};
    try {
        return JSON.parse(raw);
    } catch {
        return {};
    }
}

function saveFileToList(fileName, bundle) {
    const files = getSavedFiles();
    if (!files[fileName]) {
        logCreation('File', fileName, bundle);
    }
    files[fileName] = {
        bundle: sanitizeBundle(bundle),
        lastModified: new Date().toLocaleString()
    };
    localStorage.setItem(FILE_LIST_STORAGE_KEY, JSON.stringify(files));
}

function deleteFileFromList(fileName) {
    const files = getSavedFiles();
    logDeletion('File', fileName);
    delete files[fileName];
    localStorage.setItem(FILE_LIST_STORAGE_KEY, JSON.stringify(files));
}

function confirmDeleteRow(rowElement, onConfirm) {
  const bundle = loadBundle();
  if (bundle.deleteMode) {
    onConfirm();
    return;
  }

  rowElement.classList.add('delete-highlight');
  
  const onCancel = () => {
    rowElement.classList.remove('delete-highlight');
  };

  const overlay = createPopup("Are you sure you want to delete this row?", rowElement, onCancel);
  const content = overlay.querySelector('.popup-content');
  const btnContainer = overlay.querySelector('.popup-buttons');
  
  // Apply slide-in animation
  content.classList.remove('expanding');
  content.classList.add('slide-in-top');

  const deleteBtn = document.createElement('button');
  deleteBtn.className = 'popup-btn primary';
  deleteBtn.style.background = '#ff4444';
  deleteBtn.style.borderColor = '#ff4444';
  deleteBtn.textContent = 'Delete';
  deleteBtn.onclick = () => {
    overlay.classList.add('fade-out-slow');
    setTimeout(() => {
      onConfirm();
      overlay.remove();
    }, 1000);
  };

  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'popup-btn';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.onclick = () => {
    onCancel();
    closePopup(overlay);
  };

  btnContainer.appendChild(deleteBtn);
  btnContainer.appendChild(cancelBtn);
}

const PERMANENT_PERSONNEL_KEY = 'permanent_personnel_global';

function getPermanentPersonnel() {
  const stored = localStorage.getItem(PERMANENT_PERSONNEL_KEY);
  return stored ? JSON.parse(stored) : {};
}

function setPermanentPersonnel(data) {
  localStorage.setItem(PERMANENT_PERSONNEL_KEY, JSON.stringify(data));
}

function syncPersonnelData(fileData) {
  const global = getPermanentPersonnel();
  const merged = [];
  const processedNames = new Set();

  // 1. Process data from the file
  fileData.forEach(row => {
    const name = row[0];
    if (!name) {
      merged.push([...row]);
      return;
    }
    processedNames.add(name);
    
    if (global[name]) {
      merged.push([
        name,
        row[1] || '',
        row[2] || '',
        global[name].gps || row[3] || '',
        global[name].radio || row[4] || '',
        global[name].medic || row[5] || '',
        row[6] || 'false'
      ]);
    } else {
      global[name] = {
        gps: row[3] || '',
        radio: row[4] || '',
        medic: row[5] || ''
      };
      merged.push([...row]);
    }
  });

  // 2. Add members from global that were NOT in the file
  for (const name in global) {
    if (!processedNames.has(name)) {
      merged.push([
        name,
        '', 
        '', 
        global[name].gps || '',
        global[name].radio || '',
        global[name].medic || '',
        'false'
      ]);
    }
  }

  setPermanentPersonnel(global);
  return merged;
}

function splitPersonnelData(mergedData) {
  const global = getPermanentPersonnel();
  const filePart = [];

  mergedData.forEach(row => {
    const name = row[0];
    if (!name) {
      filePart.push([...row]);
      return;
    }

    global[name] = {
      gps: row[3] || '',
      radio: row[4] || '',
      medic: row[5] || ''
    };

    filePart.push([
      name,
      row[1] || '',
      row[2] || '',
      '', 
      '', 
      '', 
      row[6] || 'false'
    ]);
  });

  setPermanentPersonnel(global);
  return filePart;
}

function loadData() {
  const bundle = loadBundle();
  const key = pageKey();
  if (bundle.pages[key]) {
    if (key === 'page3') return syncPersonnelData(bundle.pages[key]);
    return bundle.pages[key];
  }
  if (isRegionsPage()) return defaultRegionsData();
  if (isSegmentsPage()) return defaultSegmentsData();
  if (isPersonnelPage()) return syncPersonnelData(defaultPersonnelData());
  if (isSearchLogPage()) return defaultSearchLogData();
  return defaultData();
}

function saveCurrentPageData(data) {
  const bundle = loadBundle();
  const key = pageKey();
  if (key === 'page3') {
    bundle.pages[key] = splitPersonnelData(data);
  } else {
    bundle.pages[key] = data;
  }
  saveBundle(bundle);
  const status = document.getElementById('save-status');
  if (status) {
    const now = new Date();
    status.textContent = `Saved automatically at ${now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}`;
  }
}

function updateFileNameDisplay() {
  const brandEl = document.querySelector('.brand');
  if (brandEl) brandEl.textContent = BRAND_NAME;

  const bundle = loadBundle();
  document.querySelectorAll('[data-file-name]').forEach((el) => {
    el.textContent = bundle.fileName;
  });
}

function downloadTextFile(filename, content, mimeType = 'application/json') {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

function placeCaretAtEnd(el) {
  const range = document.createRange();
  const sel = window.getSelection();
  range.selectNodeContents(el);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

function focusSelector(selector) {
  const next = document.querySelector(selector);
  if (!next) return;
  next.focus();
  placeCaretAtEnd(next);
}

function focusCell(row, col) {
  focusSelector(`.pill-cell[data-row="${row}"][data-col="${col}"]`);
}


let highlightedRowIndex = -1;

function animateNewRow(tr, index) {
  if (index === highlightedRowIndex) {
    tr.classList.add('new-row-highlight');
    // We don't need to remove it here as it's an animation that finishes on its own
    // But we should reset the index for the next table build
    setTimeout(() => {
      highlightedRowIndex = -1;
    }, 100);
  }
}

function animateArrivedRow(tr, teamName) {
  const bundle = loadBundle();
  if (bundle.arrivedTeams && bundle.arrivedTeams.includes(teamName)) {
    tr.classList.add('new-row-highlight');
    // Remove from arrivedTeams after highlighting
    bundle.arrivedTeams = bundle.arrivedTeams.filter(t => t !== teamName);
    saveBundle(bundle);
  }
}

function buildStandardTable() {
  const tableBody = document.getElementById('table-body');
  const clearBtn = document.getElementById('clear-table');
  const data = loadData();

  if (!tableBody) return;
  tableBody.innerHTML = '';

  for (let r = 0; r < data.length; r++) {
    const tr = document.createElement('tr');
    animateNewRow(tr, r);

    for (let c = 0; c < COLS; c++) {
      const td = document.createElement('td');
      td.dataset.label = `Column ${c + 1}`;
      const cellContainer = document.createElement('div');
      cellContainer.className = 'pill-cell-container';

      const cell = document.createElement('div');
      cell.className = 'pill-cell';
      cell.contentEditable = 'true';
      cell.spellcheck = false;
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);
      cell.textContent = data[r]?.[c] ?? '';
      cell.setAttribute('role', 'textbox');
      cell.setAttribute('aria-label', `Row ${r + 1}, Column ${c + 1}`);

      cell.addEventListener('blur', () => {
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);
        data[row][col] = cell.textContent.trim();
        saveCurrentPageData(data);
      });

      cell.addEventListener('keydown', (event) => {
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);

        if (event.key === 'Enter') {
          event.preventDefault();
          cell.blur();
          focusCell(Math.min(row + 1, data.length - 1), col);
          return;
        }

        if (event.key === 'Tab') {
          event.preventDefault();
          cell.blur();
          const nextCol = event.shiftKey ? Math.max(col - 1, 0) : Math.min(col + 1, COLS - 1);
          focusCell(row, nextCol);
        }
      });

      cellContainer.appendChild(cell);
      td.appendChild(cellContainer);
      tr.appendChild(td);
    }

    // New Delete Column
    const deleteTd = document.createElement('td');
    deleteTd.dataset.label = 'Delete';
    const deleteContainer = document.createElement('div');
    const delBtn = document.createElement('button');
    delBtn.className = 'row-delete-btn';
    delBtn.textContent = 'Delete';
    delBtn.type = 'button';
    delBtn.onclick = () => {
      confirmDeleteRow(tr, () => {
        const rowContent = (data[r] || []).filter(Boolean).join(', ') || 'empty row';
        data.splice(r, 1);
        logDeletion('row', rowContent);
        if (data.length === 0) data.push(Array.from({ length: COLS }, () => ''));
        saveCurrentPageData(data);
        buildStandardTable();
      });
    };
    deleteContainer.appendChild(delBtn);
    deleteTd.appendChild(deleteContainer);
    tr.appendChild(deleteTd);

    tableBody.appendChild(tr);
  }

  // Add Row button
  const addRowContainer = document.createElement('div');
  addRowContainer.className = 'add-row-container';
  const addRowBtn = document.createElement('button');
  addRowBtn.className = 'add-row-btn';
  addRowBtn.textContent = '+ Add new row';
  addRowBtn.onclick = () => {
    data.push(Array.from({ length: COLS }, () => ''));
    logCreation('row', 'new empty row');
    saveCurrentPageData(data);
    highlightedRowIndex = data.length - 1;
    buildStandardTable();
    focusCell(data.length - 1, 0);
  };
  addRowContainer.appendChild(addRowBtn);

  const existing = document.querySelector('.add-row-container');
  if (existing) existing.remove();
  tableBody.parentElement.after(addRowContainer);

  if (clearBtn) {
    clearBtn.remove();
  }
}

function parseVote(value) {
  const num = Number(String(value).trim());
  return Number.isFinite(num) ? num : null;
}

function parseNumeric(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  // Remove commas and handle units like "ac", "mi", "ft", "hr"
  // parseFloat handles trailing text, but not commas in the middle
  const cleaned = String(val).replace(/,/g, '');
  const num = parseFloat(cleaned);
  return isFinite(num) ? num : 0;
}

function formatUnit(val, unit) {
  if (!val) return '';
  const cleaned = String(val).replace(/,/g, '');
  const num = parseFloat(cleaned);
  if (isNaN(num)) return val;
  return `${num} ${unit}`;
}

function computeConsensus(data, rowIndex) {
  const dynamicCount = data.headers.length - 2;
  if (dynamicCount <= 0) return '';

  let totalNormalized = 0;
  let usedColumns = 0;

  for (let voteCol = 0; voteCol < dynamicCount; voteCol++) {
    let columnSum = 0;
    for (let r = 0; r < data.rows.length; r++) {
      const value = parseVote(data.rows[r][voteCol + 1]);
      if (value !== null) columnSum += value;
    }

    const rowValue = parseVote(data.rows[rowIndex][voteCol + 1]);
    if (rowValue === null || columnSum === 0) continue;

    totalNormalized += rowValue / columnSum;
    usedColumns += 1;
  }

  if (!usedColumns) return '';
  return (totalNormalized / usedColumns).toFixed(4);
}

function updateConsensusCells(data) {
  // Update data in memory first for all rows
  for (let r = 0; r < data.rows.length; r++) {
    data.rows[r][data.headers.length - 1] = computeConsensus(data, r);
  }
  // Then update DOM if present
  document.querySelectorAll('.consensus-cell').forEach((cell) => {
    const row = Number(cell.dataset.row);
    if (data.rows[row]) {
      cell.textContent = data.rows[row][data.headers.length - 1];
    }
  });
}

function saveRegionsAndRefresh(data) {
  updateConsensusCells(data);
  saveCurrentPageData(data);
  recalculateEverything();
}

function recalculateEverything() {
  const bundle = loadBundle();
  const segmentsData = bundle.pages.page2 || [];
  const searchLogData = bundle.pages.page4 || [];
  const regionsData = bundle.pages.index;

  // 1. Recalculate Regions Consensus
  updateConsensusCells(regionsData);

  // Helper for share calculation
  const getInitialShare = (region, area) => {
    const regionRowIndex = (regionsData.rows || []).findIndex(r => r[0] === region);
    if (regionRowIndex === -1) return 0;
    const consensus = parseFloat(computeConsensus(regionsData, regionRowIndex)) || 0;
    
    let sumOfAreas = 0;
    segmentsData.forEach(r => {
      if (r[0] === region) {
        sumOfAreas += parseNumeric(r[2]);
      }
    });
    if (sumOfAreas <= 0) return 0;
    return (consensus * area / sumOfAreas);
  };

  const segInfoMap = new Map();

  // 2. Recalculate Segments PSRi and Initialize Shares
  for (let r = 0; r < segmentsData.length; r++) {
    if (segmentsData[r][0] && segmentsData[r][1]) {
       const length = parseNumeric(segmentsData[r][3]);
       const manualTime = parseNumeric(segmentsData[r][8]);
       if (manualTime > 0) {
          segmentsData[r][5] = segmentsData[r][8];
       } else if (length > 0) {
          segmentsData[r][5] = (length / 0.5).toFixed(2) + ' hr';
       } else {
          segmentsData[r][5] = '';
       }
       segmentsData[r][6] = calculatePSR(segmentsData, r, bundle);
       
       const share = getInitialShare(segmentsData[r][0], parseNumeric(segmentsData[r][2]));
       segInfoMap.set(`${segmentsData[r][0]}|${segmentsData[r][1]}`, { share });
    }
  }

  // 3. Recalculate Search Log row by row, smallest task # first
  const sortedSearchLog = [...searchLogData].filter(row => row[0]).sort((a, b) => {
    const taskA = parseInt(a[0].replace('#', '')) || 0;
    const taskB = parseInt(b[0].replace('#', '')) || 0;
    return taskA - taskB;
  });

  sortedSearchLog.forEach(logRow => {
    const region = logRow[3];
    const segment = logRow[4];
    const key = `${region}|${segment}`;
    const info = segInfoMap.get(key);
    
    const segRow = segmentsData.find(s => s[0] === region && s[1] === segment);
    if (info && segRow) {
      const area = parseNumeric(segRow[2]);
      const length = parseNumeric(segRow[3]);
      const timePerSweep = parseNumeric(segRow[5]);
      const sweepWidth = parseNumeric(logRow[8]);
      const numSweeps = parseNumeric(logRow[9]);
      const teamInfo = logRow[7] || '';
      const match = teamInfo.match(/\((\d+)\)/);
      const numMembers = match ? parseInt(match[1]) : 0;

      // PSR Before using search sweep
      const psrBefore = (length / timePerSweep * sweepWidth * info.share) / (area / 640);
      logRow[5] = isFinite(psrBefore) ? psrBefore.toFixed(4) : '';

      if (sweepWidth > 0 && numSweeps > 0 && numMembers > 0 && area > 0 && length > 0) {
        const z = sweepWidth / ((area / 640 / length / numSweeps / numMembers) * 5280);
        info.share *= Math.exp(-z);
      }

      // PSR After using search sweep
      const psrAfter = (length / timePerSweep * sweepWidth * info.share) / (area / 640);
      logRow[6] = isFinite(psrAfter) ? psrAfter.toFixed(4) : '';
    }
  });

  // 4. Update segment's PSRc with final share using segment's default sweep
  for (let r = 0; r < segmentsData.length; r++) {
    const region = segmentsData[r][0];
    const segment = segmentsData[r][1];
    const info = segInfoMap.get(`${region}|${segment}`);
    if (info) {
      const length = parseNumeric(segmentsData[r][3]);
      const timePerSweep = parseNumeric(segmentsData[r][5]);
      const sweep = parseNumeric(segmentsData[r][4]);
      const area = parseNumeric(segmentsData[r][2]);
      
      const psrc = (length / timePerSweep * sweep * info.share) / (area / 640);
      segmentsData[r][7] = isFinite(psrc) ? psrc.toFixed(4) : '';
    }
  }

  // Map back to original searchLogData
  searchLogData.forEach(row => {
     const sortedRow = sortedSearchLog.find(sr => sr[0] === row[0]);
     if (sortedRow) {
       row[5] = sortedRow[5];
       row[6] = sortedRow[6];
     }
  });

  saveBundle(bundle);
}

function buildRegionsTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  const clearBtn = document.getElementById('clear-table');
  const addBtn = document.getElementById('add-column');
  const deleteBtn = document.getElementById('delete-column');
  const data = loadData();
  const dynamicCount = data.headers.length - 2;

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headerRow = document.createElement('tr');
  for (let c = 0; c < data.headers.length; c++) {
    const th = document.createElement('th');
    const isFixed = c === 0 || c === data.headers.length - 1;

    if (isFixed) {
      th.textContent = data.headers[c];
      th.className = 'fixed-header';
    } else {
      const headerPill = document.createElement('div');
      headerPill.className = 'pill-cell header-pill';
      headerPill.contentEditable = 'true';
      headerPill.spellcheck = false;
      headerPill.dataset.headerCol = String(c);
      headerPill.textContent = data.headers[c];
      headerPill.setAttribute('role', 'textbox');
      headerPill.setAttribute('aria-label', `Voter header ${c}`);

      headerPill.addEventListener('blur', () => {
        data.headers[c] = headerPill.textContent.trim() || `Voter ${c}`;
        saveRegionsAndRefresh(data);
      });

      headerPill.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          headerPill.blur();
          focusSelector(`.pill-cell[data-row="0"][data-col="${c}"]`);
          return;
        }

        if (event.key === 'Tab') {
          event.preventDefault();
          headerPill.blur();
          const targetCol = event.shiftKey ? Math.max(c - 1, 1) : Math.min(c + 1, data.headers.length - 2);
          focusSelector(`.header-pill[data-header-col="${targetCol}"]`);
        }
      });

      th.appendChild(headerPill);
    }

    headerRow.appendChild(th);
  }

  // Delete Header
  const deleteTh = document.createElement('th');
  deleteTh.textContent = 'Delete';
  deleteTh.className = 'fixed-header';
  headerRow.appendChild(deleteTh);

  tableHead.appendChild(headerRow);

  for (let r = 0; r < data.rows.length; r++) {
    const tr = document.createElement('tr');
    animateNewRow(tr, r);

    for (let c = 0; c < data.headers.length; c++) {
      const td = document.createElement('td');
      td.dataset.label = data.headers[c];

      if (c === data.headers.length - 1) {
        const consensus = document.createElement('div');
        consensus.className = 'pill-cell consensus-cell readonly-pill';
        consensus.dataset.row = String(r);
        consensus.textContent = computeConsensus(data, r);
        td.appendChild(consensus);
        tr.appendChild(td);
        continue;
      }

      const cellContainer = document.createElement('div');
      cellContainer.className = 'pill-cell-container';

      const cell = document.createElement('div');
      cell.className = 'pill-cell';
      cell.contentEditable = 'true';
      cell.spellcheck = false;
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);
      cell.textContent = data.rows[r]?.[c] ?? '';
      cell.setAttribute('role', 'textbox');
      cell.setAttribute('aria-label', `Row ${r + 1}, ${data.headers[c]}`);

      cell.addEventListener('blur', () => {
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);
        data.rows[row][col] = cell.textContent.trim();
        saveRegionsAndRefresh(data);
      });

      cell.addEventListener('keydown', (event) => {
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);

        if (event.key === 'Enter') {
          event.preventDefault();
          cell.blur();
          focusCell(Math.min(row + 1, data.rows.length - 1), col);
          return;
        }

        if (event.key === 'Tab') {
          event.preventDefault();
          cell.blur();
          const maxEditableCol = data.headers.length - 2;
          const nextCol = event.shiftKey ? Math.max(col - 1, 0) : Math.min(col + 1, maxEditableCol);
          focusCell(row, nextCol);
        }
      });

      cellContainer.appendChild(cell);
      td.appendChild(cellContainer);
      tr.appendChild(td);
    }

    // New Delete Column
    const deleteTd = document.createElement('td');
    deleteTd.dataset.label = 'Delete';
    const deleteContainer = document.createElement('div');
    deleteContainer.className = 'pill-cell-container';
    const delBtn = document.createElement('button');
    delBtn.className = 'row-delete-btn';
    delBtn.textContent = 'Delete';
    delBtn.type = 'button';
    delBtn.onclick = () => {
      confirmDeleteRow(tr, () => {
        const regionName = (data.rows[r] && data.rows[r][0]) || 'unnamed region';
        data.rows.splice(r, 1);
        logDeletion('Region', regionName);
        if (data.rows.length === 0) {
          data.rows.push(Array.from({ length: data.headers.length }, () => ''));
        }
        saveCurrentPageData(data);
        buildRegionsTable();
      });
    };
    deleteContainer.appendChild(delBtn);
    deleteTd.appendChild(deleteContainer);
    tr.appendChild(deleteTd);

    tableBody.appendChild(tr);
  }

  // Add Row button
  const addRowContainer = document.createElement('div');
  addRowContainer.className = 'add-row-container';
  const addRowBtn = document.createElement('button');
  addRowBtn.className = 'add-row-btn';
  addRowBtn.textContent = '+ Add new row';
  addRowBtn.onclick = () => {
    data.rows.push(Array.from({ length: data.headers.length }, () => ''));
    logCreation('Region', 'new empty region');
    saveRegionsAndRefresh(data);
    highlightedRowIndex = data.rows.length - 1;
    buildRegionsTable();
    focusCell(data.rows.length - 1, 0);
  };
  addRowContainer.appendChild(addRowBtn);

  const existing = document.querySelector('.add-row-container');
  if (existing) existing.remove();
  tableBody.parentElement.after(addRowContainer);

  updateConsensusCells(data);
  saveCurrentPageData(data);

  addBtn.onclick = () => {
    const insertAt = data.headers.length - 1;
    const newHeaderName = `Voter ${dynamicCount + 1}`;
    data.headers.splice(insertAt, 0, newHeaderName);
    data.rows.forEach((row) => row.splice(insertAt, 0, ''));
    logCreation('Voter Column', newHeaderName);
    saveCurrentPageData(data);
    buildRegionsTable();
  };

  deleteBtn.onclick = () => {
    const currentDynamicCount = data.headers.length - 2;
    if (currentDynamicCount <= 1) return;
    const removeAt = data.headers.length - 2;
    const headerName = data.headers[removeAt];
    data.headers.splice(removeAt, 1);
    data.rows.forEach((row) => row.splice(removeAt, 1));
    logDeletion('Voter Column', headerName);
    saveCurrentPageData(data);
    buildRegionsTable();
  };

  if (clearBtn) {
    clearBtn.remove();
  }
}

function calculatePSR(data, rowIndex, bundle) {
  const row = data[rowIndex];
  const regionName = row[0];
  const area = parseNumeric(row[2]);
  const length = parseNumeric(row[3]);
  const sweep = parseNumeric(row[4]);
  
  // Calculate Time per Sweep as Length / 0.5 if no manual override
  const manualTime = parseNumeric(row[8]);
  const timePerSweep = manualTime > 0 ? manualTime : (length / 0.5);

  if (!regionName || area <= 0 || length <= 0 || sweep <= 0 || timePerSweep <= 0) return '';

  const regionsData = bundle.pages.index;
  const regionRowIndex = (regionsData.rows || []).findIndex(r => r[0] === regionName);
  if (regionRowIndex === -1) return '';

  const consensus = parseFloat(computeConsensus(regionsData, regionRowIndex)) || 0;
  if (consensus <= 0) return '';

  // Sum of Areas for all segments in this region
  let sumOfAreas = 0;
  data.forEach(r => {
    if (r[0] === regionName) {
      sumOfAreas += parseNumeric(r[2]);
    }
  });

  if (sumOfAreas <= 0) return '';

  // Formula: PSR = ((Sweep * Length ) / Time) * (Consensus * (Area / Sum of Areas)) / (Area / 640)
  const psr = ((sweep * length ) / timePerSweep) * (consensus * (area / sumOfAreas)) / (area / 640);

  return isFinite(psr) ? psr.toFixed(4) : '';
}

// updateAllPSRs removed, logic moved to recalculateEverything

function buildSegmentsTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  const clearBtn = document.getElementById('clear-table');
  const sortToggle = document.getElementById('sort-toggle');
  const sortLabel = document.getElementById('sort-label');
  
  recalculateEverything();
  let data = loadData();
  const bundle = loadBundle();

  const isPSRDescending = sortToggle && sortToggle.checked;
  if (sortLabel) {
    sortLabel.textContent = isPSRDescending ? 'Sorted by PSRc (Descending)' : 'Sorted by Region then Segment';
  }

  const sortedData = [...data].sort((a, b) => {
    if (isPSRDescending) {
      const psrA = parseFloat(a[7]) || 0;
      const psrB = parseFloat(b[7]) || 0;
      return psrB - psrA;
    } else {
      // Region then Segment
      const regionA = (a[0] || '').toLowerCase();
      const regionB = (b[0] || '').toLowerCase();
      if (regionA < regionB) return -1;
      if (regionA > regionB) return 1;

      // Same region, check segment (index 1)
      const segA = parseNumeric(a[1]);
      const segB = parseNumeric(b[1]);
      return segA - segB;
    }
  });

  const activeSegments = new Set();
  if (bundle.currentAssignments && bundle.teamStatuses) {
    for (const team in bundle.currentAssignments) {
      const status = bundle.teamStatuses[team] || '';
      const assignment = bundle.currentAssignments[team] || '';
      if (!status.includes('at base') && assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
        const match = assignment.match(/#\d+ (.+) - (.+)/);
        if (match) {
          activeSegments.add(`${match[1]}|${match[2]}`);
        }
      }
    }
  }

  const logSweepsDue = getLogSweepsDue();
  const dueSegments = new Set(logSweepsDue.map(d => `${d.region}|${d.segment}`));

  const regionsData = bundle.pages.index;
  const regionNames = (regionsData.rows || []).map(r => r[0]).filter(name => name && name.trim() !== '');

  const actionsContainer = document.getElementById('segment-header-actions');
  if (actionsContainer) {
      actionsContainer.innerHTML = '';
      const importBtn = document.createElement('button');
      importBtn.className = 'clear-btn';
      importBtn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 8px; vertical-align: middle;"><path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M8 5H4a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-4"/></svg>Import JSON';
      importBtn.onclick = showImportSegmentsPopup;
      actionsContainer.appendChild(importBtn);
  }

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headers = ['Region', 'Segment', 'Area (acres)', 'Length (mi)', 'Sweep (ft)', 'Time per Sweep (hr)', 'PSRi', 'PSRc', 'Delete'];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.className = 'fixed-header';
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  for (let r = 0; r < sortedData.length; r++) {
    const tr = document.createElement('tr');
    animateNewRow(tr, r);
    const segKey = `${sortedData[r][0]}|${sortedData[r][1]}`;
    if (activeSegments.has(segKey)) {
      tr.classList.add('unfinished-row');
    } else if (dueSegments.has(segKey)) {
      tr.classList.add('log-sweeps-due');
    }

    if (newlyImportedSegments.has(segKey)) {
      tr.classList.add('new-import-row');
    }

    const headers = ['Region', 'Segment', 'Area (acres)', 'Length (mi)', 'Sweep (ft)', 'Time per Sweep (hr)', 'PSRi', 'PSRc', 'Delete'];
    for (let c = 0; c < 8; c++) {
      const td = document.createElement('td');
      td.dataset.label = headers[c];
      const cellContainer = document.createElement('div');
      cellContainer.className = 'pill-cell-container';

      if (c === 0) {
        // Region Dropdown
        const select = document.createElement('select');
        select.className = 'pill-cell segment-select';
        if (newlyImportedSegments.has(segKey)) select.classList.add('new-import-highlight');
        select.dataset.row = String(r);
        select.dataset.col = String(c);
        select.style.width = '100%';

        const emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = '-- Select --';
        select.appendChild(emptyOpt);

        regionNames.forEach(name => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          if (sortedData[r][c] === name) opt.selected = true;
          select.appendChild(opt);
        });

        select.onchange = () => {
          const originalRow = sortedData[r];
          originalRow[c] = select.value;
          saveCurrentPageData(data);
          buildSegmentsTable();
        };

        select.addEventListener('keydown', (event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            focusCell(Math.min(r + 1, sortedData.length - 1), c);
          }
          if (event.key === 'Tab') {
            event.preventDefault();
            const nextCol = event.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, 7);
            focusCell(r, nextCol);
          }
        });

        cellContainer.appendChild(select);
      } else if (c === 5) {
        // Time per Sweep (hr) with manual override logic
        if (sortedData[r][8]) {
          const resetBtn = document.createElement('button');
          resetBtn.className = 'reset-pill-btn';
          resetBtn.textContent = '↺';
          resetBtn.title = 'Reset to calculated value';
          resetBtn.onclick = (e) => {
             e.stopPropagation();
             const originalRowIndex = data.indexOf(sortedData[r]);
             data[originalRowIndex][8] = '';
             saveCurrentPageData(data);
             buildSegmentsTable();
          };
          cellContainer.appendChild(resetBtn);
        }

        const cell = document.createElement('div');
        cell.className = 'pill-cell';
        if (newlyImportedSegments.has(segKey)) cell.classList.add('new-import-highlight');
        cell.contentEditable = 'true';
        cell.spellcheck = false;
        cell.dataset.row = String(r);
        cell.dataset.col = String(c);
        cell.textContent = sortedData[r][c] || '';

        cell.addEventListener('blur', () => {
          let val = cell.textContent.trim();
          const originalRowIndex = data.indexOf(sortedData[r]);
          if (val) {
            val = formatUnit(val, 'hr');
            // Check if it's different from the calculated value (length / 0.5)
            const length = parseNumeric(sortedData[r][3]);
            const calc = (length / 0.5).toFixed(2) + ' hr';
            if (val !== calc) {
              data[originalRowIndex][8] = val; // Store manual override
            } else {
              data[originalRowIndex][8] = ''; // Clear override
            }
          } else {
            data[originalRowIndex][8] = '';
          }
          saveCurrentPageData(data);
          buildSegmentsTable();
        });

        cell.addEventListener('keydown', (event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            focusCell(Math.min(r + 1, sortedData.length - 1), c);
          }
          if (event.key === 'Tab') {
            event.preventDefault();
            const nextCol = event.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, 7);
            focusCell(r, nextCol);
          }
        });

        cellContainer.appendChild(cell);
      } else if (c === 6 || c === 7) {
        // PSR Columns (Read-only)
        cellContainer.className = 'pill-cell-container psr-cell-container';

        const cell = document.createElement('div');
        cell.className = 'pill-cell readonly-pill';
        if (newlyImportedSegments.has(segKey)) cell.classList.add('new-import-highlight');
        cell.contentEditable = 'false';
        cell.spellcheck = false;
        cell.dataset.row = String(r);
        cell.dataset.col = String(c);
        cell.textContent = sortedData[r][c] || '';
        
        cell.addEventListener('keydown', (event) => {
          if (event.key === 'Tab') {
            event.preventDefault();
            const nextCol = event.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, 7);
            focusCell(r, nextCol);
          }
        });

        cellContainer.appendChild(cell);

        if (c === 7) {
          const sweepsDue = getLogSweepsDue();
          const isDue = sweepsDue.find(d => d.region === sortedData[r][0] && d.segment === sortedData[r][1]);
          
          const actionBtn = document.createElement('button');
          actionBtn.className = 'row-search-btn';
          actionBtn.type = 'button';
          
          if (isDue) {
            actionBtn.textContent = 'log sweeps';
            actionBtn.classList.add('log-sweeps-active');
            actionBtn.onclick = () => showLogSweepsPopup(isDue.taskNum);
          } else {
            actionBtn.textContent = 'search';
            actionBtn.onclick = () => {
              showTeamSelectionPopup((teamName) => {
                showMissingStepsPopup(teamName, null, () => {
                  const region = sortedData[r][0] || '';
                  const segment = sortedData[r][1] || '';
                  const taskNumber = addAutoSearchLogEntry(teamName, region, segment);
                  const assignmentStr = `#${taskNumber} ${region} - ${segment}`;
                  
                  const bundle2 = loadBundle();
                  bundle2.currentAssignments[teamName] = assignmentStr;
                  bundle2.teamAssignmentTimes[teamName] = Date.now();
                  bundle2.teamStatuses[teamName] = 'assigned';
                  saveBundle(bundle2);
                  addActivityLogEntry(teamName, `Started search on ${assignmentStr}`);
                  
                  navigateToPage('page4.html?scroll=latest');
                });
              });
            };
          }
          cellContainer.appendChild(actionBtn);
        }
      } else {
        const cell = document.createElement('div');
        cell.className = 'pill-cell';
        if (newlyImportedSegments.has(segKey)) cell.classList.add('new-import-highlight');
        cell.contentEditable = 'true';
        cell.spellcheck = false;
        cell.dataset.row = String(r);
        cell.dataset.col = String(c);
        cell.textContent = sortedData[r]?.[c] ?? '';

        cell.addEventListener('blur', () => {
          let val = cell.textContent.trim();
          if (val) {
            if (c === 2) val = formatUnit(val, 'ac');
            else if (c === 3) val = formatUnit(val, 'mi');
            else if (c === 4) val = formatUnit(val, 'ft');
            else if (c === 5) val = formatUnit(val, 'hr');
          }
          const originalRow = sortedData[r];
          originalRow[c] = val;
          saveCurrentPageData(data);
          buildSegmentsTable();
        });

        cell.addEventListener('keydown', (event) => {
          const row = Number(cell.dataset.row);
          const col = Number(cell.dataset.col);
          if (event.key === 'Enter') {
            event.preventDefault();
            cell.blur();
            focusCell(Math.min(row + 1, sortedData.length - 1), col);
            return;
          }
          if (event.key === 'Tab') {
            event.preventDefault();
            cell.blur();
            const nextCol = event.shiftKey ? Math.max(col - 1, 0) : Math.min(col + 1, 7);
            focusCell(row, nextCol);
          }
        });

        cellContainer.appendChild(cell);
      }
      td.appendChild(cellContainer);
      tr.appendChild(td);
    }

    const deleteTd = document.createElement('td');
    deleteTd.dataset.label = 'Delete';
    const deleteContainer = document.createElement('div');
    deleteContainer.className = 'pill-cell-container';
    const delBtn = document.createElement('button');
    delBtn.className = 'row-delete-btn';
    delBtn.textContent = 'Delete';
    delBtn.type = 'button';
    delBtn.onclick = () => {
      confirmDeleteRow(tr, () => {
        const segName = (sortedData[r] && sortedData[r][1]) || 'unnamed segment';
        const indexInData = data.indexOf(sortedData[r]);
        if (indexInData > -1) {
          data.splice(indexInData, 1);
          logDeletion('Segment', segName);
          if (data.length === 0) data.push(Array.from({ length: 8 }, () => ''));
          saveCurrentPageData(data);
          buildSegmentsTable();
        }
      });
    };
    deleteContainer.appendChild(delBtn);
    deleteTd.appendChild(deleteContainer);
    tr.appendChild(deleteTd);

    tableBody.appendChild(tr);
  }

  const addRowContainer = document.createElement('div');
  addRowContainer.className = 'add-row-container';
  const addRowBtn = document.createElement('button');
  addRowBtn.className = 'add-row-btn';
  addRowBtn.textContent = '+ Add new segment';
  addRowBtn.onclick = () => {
    data.push(Array.from({ length: 8 }, () => ''));
    logCreation('Segment', 'new empty segment');
    saveCurrentPageData(data);
    highlightedRowIndex = data.length - 1;
    buildSegmentsTable();
    focusCell(data.length - 1, 1);
  };
  addRowContainer.appendChild(addRowBtn);
  const existing = document.querySelector('.add-row-container');
  if (existing) existing.remove();
  tableBody.parentElement.after(addRowContainer);

  if (sortToggle) {
    sortToggle.onchange = () => {
      buildSegmentsTable();
    };
  }

  if (clearBtn) {
    clearBtn.remove();
  }
}

let currentPersonnelSubpage = 'activity';

function buildPersonnelTable() {
  const btnAll = document.getElementById('btn-all-members');
  const btnAct = document.getElementById('btn-activity');
  const btnTeamRep = document.getElementById('btn-team-reports');
  const btnMemRep = document.getElementById('btn-member-reports');

  const logContainer = document.getElementById('activity-log-container');
  const baseContainer = document.getElementById('base-teams-container-header');
  const teamReportsContainer = document.getElementById('team-reports-container');
  const memberReportsContainer = document.getElementById('member-reports-container');
  const searchTeamsContainer = document.getElementById('search-teams-container');
  const controls = document.getElementById('all-members-controls');

  const subNavBtns = [btnAll, btnAct, btnTeamRep, btnMemRep];
  const containers = [controls, baseContainer, teamReportsContainer, memberReportsContainer, searchTeamsContainer];

  const user = getCurrentUser();
  const isSubAllowed = (sub) => {
      return true;
  };

  if (btnAct) btnAct.style.display = isSubAllowed('activity') ? 'inline-block' : 'none';
  if (btnTeamRep) btnTeamRep.style.display = isSubAllowed('team-reports') ? 'inline-block' : 'none';
  if (btnMemRep) btnMemRep.style.display = isSubAllowed('member-reports') ? 'inline-block' : 'none';
  if (btnAll) btnAll.style.display = isSubAllowed('all-members') ? 'inline-block' : 'none';

  // If current subpage is not allowed, switch to first allowed
  if (!isSubAllowed(currentPersonnelSubpage)) {
      if (isSubAllowed('activity')) currentPersonnelSubpage = 'activity';
      else if (isSubAllowed('all-members')) currentPersonnelSubpage = 'all-members';
      else if (isSubAllowed('team-reports')) currentPersonnelSubpage = 'team-reports';
      else if (isSubAllowed('member-reports')) currentPersonnelSubpage = 'member-reports';
  }

  function hideAll() {
    containers.forEach(c => { if (c) c.style.display = 'none'; });
    subNavBtns.forEach(b => { if (b) b.classList.remove('active'); });
  }

  if (btnAll) {
    btnAll.onclick = () => {
      currentPersonnelSubpage = 'all-members';
      buildPersonnelTable();
    };
  }
  if (btnAct) {
    btnAct.onclick = () => {
      currentPersonnelSubpage = 'activity';
      buildPersonnelTable();
    };
  }
  if (btnTeamRep) {
    btnTeamRep.onclick = () => {
      currentPersonnelSubpage = 'team-reports';
      buildPersonnelTable();
    };
  }
  if (btnMemRep) {
    btnMemRep.onclick = () => {
      currentPersonnelSubpage = 'member-reports';
      buildPersonnelTable();
    };
  }

  const btnPrintTeam = document.getElementById('print-team-reports');
  const btnPrintAllTeam = document.getElementById('print-all-team-reports');
  const btnPrintMem = document.getElementById('print-member-reports');
  const btnPrintAllMem = document.getElementById('print-all-member-reports');

  if (btnPrintTeam) btnPrintTeam.onclick = () => printCurrentReport('team');
  if (btnPrintAllTeam) btnPrintAllTeam.onclick = () => printAllReports('team');
  if (btnPrintMem) btnPrintMem.onclick = () => printCurrentReport('member');
  if (btnPrintAllMem) btnPrintAllMem.onclick = () => printAllReports('member');

  const btnReset = document.getElementById('btn-reset-members');
  if (btnReset) {
    btnReset.onclick = () => {
      const popup = createPopup('Delete All Members?');
      const content = popup.querySelector('.popup-content');
      const btnContainer = popup.querySelector('.popup-buttons');
      
      // Semi-transparent red theme
      content.style.background = 'rgba(150, 0, 0, 0.9)';
      content.style.borderColor = '#ff4444';
      content.style.color = '#fff';

      const msg = document.createElement('p');
      msg.textContent = 'This will permanently delete all members from the list. Please type your name to confirm:';
      msg.style.marginBottom = '15px';
      content.insertBefore(msg, btnContainer);

      const confirmInput = document.createElement('input');
      confirmInput.type = 'text';
      confirmInput.placeholder = 'Your Name';
      confirmInput.className = 'cell-edit-input';
      confirmInput.style.width = '100%';
      confirmInput.style.marginBottom = '20px';
      confirmInput.style.background = 'rgba(255, 255, 255, 0.1)';
      confirmInput.style.color = '#fff';
      content.insertBefore(confirmInput, btnContainer);

      const confirmBtn = document.createElement('button');
      confirmBtn.className = 'popup-btn primary';
      confirmBtn.style.background = '#ff4444';
      confirmBtn.textContent = 'Delete All Members';
      confirmBtn.disabled = true;

      confirmInput.oninput = () => {
        confirmBtn.disabled = confirmInput.value.trim() === '';
      };

      confirmBtn.onclick = () => {
        const bundle = loadBundle();
        const allMembersData = bundle.pages.page3 || [];
        const memberNames = allMembersData.map(m => m[0]).filter(n => n);

        // 1. Mark finish all steps for all teams
        const teams = Object.keys(bundle.teamStatuses);
        teams.forEach(team => {
          const currentStatus = bundle.teamStatuses[team] || '';
          if (!currentStatus.startsWith('at base')) {
            const sequence = [
              { id: 'headed to assignment', log: 'Leaving base for assignment' },
              { id: 'searching', log: 'Beginning assignment' },
              { id: 'finished segment', log: 'Finished assignment' },
              { id: 'returning', log: 'Returning to base' },
              { id: 'at base', log: 'Arrived at base' }
            ];
            
            const getIndex = (s) => {
              if (s === 'assigned') return -1;
              if (s === 'headed to assignment') return 0;
              if (s === 'searching') return 1;
              if (s === 'finished segment') return 2;
              if (s === 'returning') return 3;
              if (s && s.startsWith('at base')) return 4;
              return -1;
            };

            const currentIndex = getIndex(currentStatus);
            for (let i = currentIndex + 1; i < sequence.length; i++) {
              const step = sequence[i];
              if (step.id === 'at base') {
                 const now = new Date();
                 const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
                 bundle.teamStatuses[team] = `at base (${timeStr})`;
                 bundle.currentAssignments[team] = 'Base';
                 bundle.teamAssignmentTimes[team] = Date.now();
                 if (bundle.parChecks) delete bundle.parChecks[team];
                 if (bundle.teamLeaveTimes) delete bundle.teamLeaveTimes[team];
              } else {
                 bundle.teamStatuses[team] = step.id;
                 if (!bundle.parChecks) bundle.parChecks = {};
                 bundle.parChecks[team] = { lastTime: Date.now() };
                 if (step.id === 'headed to assignment') bundle.teamLeaveTimes[team] = Date.now();
              }
              addActivityLogEntry(team, step.log, bundle);
            }
          }
        });

        // 2. Log deleted members
        if (memberNames.length > 0) {
          addActivityLogEntry('System', 'Deleted all members: ' + memberNames.join(', '), bundle);
        }

        // 3. Clear data
        localStorage.removeItem(PERMANENT_PERSONNEL_KEY);
        bundle.pages.page3 = [];
        
        saveBundle(bundle);
        closePopup(popup);
        buildPersonnelTable();
      };
      btnContainer.appendChild(confirmBtn);
    };
  }

  const btnCallAll = document.getElementById('btn-call-all-to-base');
  if (btnCallAll) {
    btnCallAll.onclick = callAllTeamsToBase;
  }

  hideAll();

  if (logContainer) logContainer.style.display = 'block';

  if (currentPersonnelSubpage === 'all-members') {
    if (controls) controls.style.display = 'flex';
    if (searchTeamsContainer) {
      searchTeamsContainer.style.display = 'block';
      const h2 = searchTeamsContainer.querySelector('h2');
      if (h2) h2.textContent = 'All Members';
    }
    if (btnAll) btnAll.classList.add('active');
    if (btnReset) btnReset.style.display = 'block';
    buildPersonnelAllMembersTable();
  } else if (currentPersonnelSubpage === 'activity') {
    if (searchTeamsContainer) {
      searchTeamsContainer.style.display = 'block';
      const h2 = searchTeamsContainer.querySelector('h2');
      if (h2) h2.textContent = 'Search Teams';
    }
    if (baseContainer) baseContainer.style.display = 'block';
    if (btnAct) btnAct.classList.add('active');
    buildPersonnelActivityTable();
  } else if (currentPersonnelSubpage === 'team-reports') {
    if (teamReportsContainer) teamReportsContainer.style.display = 'block';
    if (btnTeamRep) btnTeamRep.classList.add('active');
    buildTeamReports();
  } else if (currentPersonnelSubpage === 'member-reports') {
    if (memberReportsContainer) memberReportsContainer.style.display = 'block';
    if (btnMemRep) btnMemRep.classList.add('active');
    buildMemberReports();
  }
  updateActivityLogUI();
}

function buildPersonnelAllMembersTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  const onSceneToggle = document.getElementById('on-scene-toggle');
  const onSceneLabel = document.getElementById('on-scene-label');
  const sortToggle = document.getElementById('personnel-sort-toggle');
  const sortLabel = document.getElementById('personnel-sort-label');
  const data = loadData();

  // Ensure first row has 'Off Duty' if it's empty
  if (data.length > 0 && !data[0][0] && !data[0][1]) {
    data[0][1] = 'Off Duty';
    saveCurrentPageData(data);
  }

  const filterOnScene = onSceneToggle && onSceneToggle.checked;
  const sortByTeam = sortToggle && sortToggle.checked;

  if (onSceneLabel) onSceneLabel.textContent = filterOnScene ? 'Filter: On Scene' : 'Filter: All Members';
  if (sortLabel) sortLabel.textContent = sortByTeam ? 'Sort: Team then Name' : 'Sort: Name';

  if (onSceneToggle && !onSceneToggle.dataset.listenerAdded) {
    onSceneToggle.addEventListener('change', buildPersonnelTable);
    onSceneToggle.dataset.listenerAdded = 'true';
  }
  if (sortToggle && !sortToggle.dataset.listenerAdded) {
    sortToggle.addEventListener('change', buildPersonnelTable);
    sortToggle.dataset.listenerAdded = 'true';
  }

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headers = ['Name', 'Team', 'GPS', 'Radio', 'Medic', 'On Scene', 'Delete'];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.className = 'fixed-header';
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  const teamOptions = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot', 'Golf', 'Command', 'Off Duty', 'Base Support'];
  
  let filteredData = [...data];
  if (filterOnScene) {
    filteredData = filteredData.filter(row => row[6] === 'true');
  }

  filteredData.sort((a, b) => {
    if (sortByTeam) {
      const teamA = (a[1] || '').toLowerCase();
      const teamB = (b[1] || '').toLowerCase();
      if (teamA < teamB) return -1;
      if (teamA > teamB) return 1;
    }
    const nameA = (a[0] || '').toLowerCase();
    const nameB = (b[0] || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  if (highlightedRowIndex === -2 && window.lastAddedRow) {
    highlightedRowIndex = filteredData.indexOf(window.lastAddedRow);
    if (highlightedRowIndex === -1) {
      const targetStr = JSON.stringify(window.lastAddedRow);
      highlightedRowIndex = filteredData.findIndex(r => JSON.stringify(r) === targetStr);
    }
    window.lastAddedRow = null;
  }

  for (let r = 0; r < filteredData.length; r++) {
    const tr = document.createElement('tr');
    animateNewRow(tr, r);
    const originalRowIndex = data.indexOf(filteredData[r]);
    
    const rowHeaders = ['Name', 'Team', 'GPS', 'Radio', 'Medic', 'On Scene', 'Delete'];
    for (let c = 0; c < 7; c++) {
      if (c === 2) continue; // Skip Team Leader column
      const td = document.createElement('td');
      td.dataset.label = rowHeaders[c < 2 ? c : c - 1];
      const cellContainer = document.createElement('div');
      cellContainer.className = 'pill-cell-container';

      if (c === 0) {
        const cell = document.createElement('div');
        cell.className = 'mini-pill';
        cell.style.width = '100%';
        cell.style.cursor = 'text';
        cell.contentEditable = 'true';
        cell.spellcheck = false;
        cell.textContent = filteredData[r][c] || '';
        cell.addEventListener('blur', () => {
          const newName = cell.textContent.trim();
          if (data[originalRowIndex][c] !== newName) {
            data[originalRowIndex][c] = newName;
            saveCurrentPageData(data);
            if (!window.pendingEmptyCellFocus) {
               buildPersonnelTable();
            }
          }
        });
        cell.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            const newName = cell.textContent.trim();
            if (data[originalRowIndex][c] !== newName) {
              data[originalRowIndex][c] = newName;
              saveCurrentPageData(data);
            }
            window.pendingEmptyCellFocus = { colLabel: 'Name' };
            buildPersonnelTable();
          }
        });
        cellContainer.appendChild(cell);
      } else if (c === 1) {
        const select = document.createElement('select');
        select.className = 'mini-pill personnel-select';
        select.style.width = '100%';
        select.style.cursor = 'pointer';
        const emptyOpt = document.createElement('option');
        emptyOpt.value = ''; emptyOpt.textContent = '-- Select Team --';
        select.appendChild(emptyOpt);
        teamOptions.forEach(team => {
          const opt = document.createElement('option');
          opt.value = team; opt.textContent = team;
          if (filteredData[r][c] === team) opt.selected = true;
          select.appendChild(opt);
        });
        select.onchange = () => {
          const newTeam = select.value;
          const oldTeam = data[originalRowIndex][c];
          const memberName = data[originalRowIndex][0];
          data[originalRowIndex][c] = newTeam;
          
          // Auto-switch team lead to new team's lead
          const existingTeamMember = data.find(row => row[1] === newTeam && row[2]);
          const newLead = existingTeamMember ? existingTeamMember[2] : '';
          data[originalRowIndex][2] = newLead;
          
          saveCurrentPageData(data);
          addActivityLogEntry(newTeam || 'Personnel', `${memberName} moved from team ${oldTeam || 'none'} to ${newTeam || 'none'}`);
          buildPersonnelTable();
        };
        cellContainer.appendChild(select);
      } else if (c === 2) {
        const selectLead = document.createElement('select');
        selectLead.className = 'mini-pill personnel-select';
        selectLead.style.width = '100%';
        selectLead.style.cursor = 'pointer';
        const emptyOptLead = document.createElement('option');
        emptyOptLead.value = ''; emptyOptLead.textContent = '-- Select Leader --';
        selectLead.appendChild(emptyOptLead);
        
        const currentTeam = filteredData[r][1];
        const possibleLeads = data.filter(row => row[1] === currentTeam && row[0]);
        
        possibleLeads.forEach(row => {
          const opt = document.createElement('option');
          opt.value = row[0]; opt.textContent = row[0];
          if (filteredData[r][c] === row[0]) opt.selected = true;
          selectLead.appendChild(opt);
        });
        selectLead.onchange = () => {
          const newLead = selectLead.value;
          const targetTeam = data[originalRowIndex][1];
          const oldLead = data[originalRowIndex][c];
          
          // Update everyone on the same team to have this new lead
          data.forEach(row => {
            if (row[1] === targetTeam) {
              row[2] = newLead;
            }
          });
          
          saveCurrentPageData(data);
          addActivityLogEntry(targetTeam || 'Personnel', `Team lead changed from ${oldLead || 'none'} to ${newLead || 'none'}`);
          buildPersonnelTable();
        };
        cellContainer.appendChild(selectLead);
      } else {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'pill-checkbox';
        checkbox.checked = filteredData[r][c] === 'true';
        checkbox.onchange = () => {
          const isChecked = checkbox.checked;
          const memberName = data[originalRowIndex][0];
          
          if (c === 6) { // On Scene column
            showTimePrompt(isChecked ? 'Mark On Scene' : 'Mark Off Scene', (date, time) => {
              data[originalRowIndex][c] = isChecked ? 'true' : 'false';
              if (!isChecked) {
                data[originalRowIndex][1] = ''; // Clear Team
                data[originalRowIndex][2] = ''; // Clear Team Lead
              }
              saveCurrentPageData(data);
              addActivityLogEntry('Personnel', `${memberName} is now ${isChecked ? 'On Scene' : 'Off Scene'} at ${date} ${time}`, null, memberName);
              buildPersonnelTable();
            }, () => {
              checkbox.checked = !isChecked;
            });
          } else { // GPS, Radio, Medic columns
            data[originalRowIndex][c] = isChecked ? 'true' : 'false';
            saveCurrentPageData(data);
          }
        };
        cellContainer.appendChild(checkbox);
      }
      td.appendChild(cellContainer);
      tr.appendChild(td);
    }

    const deleteTd = document.createElement('td');
    deleteTd.dataset.label = 'Delete';
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'row-delete-btn';
    deleteBtn.textContent = 'Delete';
    deleteBtn.onclick = () => {
      confirmDeleteRow(tr, () => {
        const memberName = (data[originalRowIndex] && data[originalRowIndex][0]) || 'unnamed person';
        data.splice(originalRowIndex, 1);
        logDeletion('Personnel', memberName);
        if (data.length === 0) data.push(Array.from({ length: 7 }, () => ''));
        saveCurrentPageData(data);
        buildPersonnelTable();
      });
    };
    deleteTd.appendChild(deleteBtn);
    tr.appendChild(deleteTd);
    tableBody.appendChild(tr);
  }

  const addRowContainer = document.createElement('div');
  addRowContainer.className = 'add-row-container';
  const addRowBtn = document.createElement('button');
  addRowBtn.className = 'add-row-btn';
  addRowBtn.textContent = '+ Add new person';
  addRowBtn.onclick = () => {
    const newRow = Array.from({ length: 7 }, () => '');
    newRow[1] = 'Off Duty';
    newRow[6] = 'true'; // Default to On Scene so it shows up if filter is on
    
    data.push(newRow);
    logCreation('Personnel', 'New Person');
    saveCurrentPageData(data);
    
    // Clear filters to ensure new row is visible
    if (onSceneToggle) {
        onSceneToggle.checked = false;
        if (onSceneLabel) onSceneLabel.textContent = 'Filter: All Members';
    }
    
    highlightedRowIndex = -2;
    window.lastAddedRow = newRow;
    buildPersonnelTable();
    
    // Focus the Name cell of the new row and select text
    setTimeout(() => {
      const highlightedRow = tableBody.querySelector('.new-row-highlight');
      if (highlightedRow) {
        const nameCell = highlightedRow.querySelector('.mini-pill[contenteditable="true"]');
        if (nameCell) {
          nameCell.focus();
          const range = document.createRange();
          range.selectNodeContents(nameCell);
          const selection = window.getSelection();
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }
    }, 100);
  };
  addRowContainer.appendChild(addRowBtn);
  const existingAllMem = document.querySelector('.add-row-container');
  if (existingAllMem) existingAllMem.remove();
  tableBody.parentElement.after(addRowContainer);

  if (window.pendingEmptyCellFocus) {
    const colLabel = window.pendingEmptyCellFocus.colLabel;
    window.pendingEmptyCellFocus = null;
    setTimeout(() => {
      const rows = tableBody.querySelectorAll('tr');
      for (const row of rows) {
        const targetCell = row.querySelector(`td[data-label="${colLabel}"] .mini-pill[contenteditable="true"]`);
        if (targetCell && !targetCell.textContent.trim()) {
          targetCell.focus();
          return;
        }
      }
    }, 100);
  }
}

function getLatestPSR(region, segment) {
  const bundle = loadBundle();
  const segData = bundle.pages.page2 || [];
  for (let i = 0; i < segData.length; i++) {
    if (segData[i][0] === region && segData[i][1] === segment) {
      // Index 7 is PSRc. If empty, fallback to index 6 (PSRi)
      return segData[i][7] || segData[i][6] || '';
    }
  }
  return '';
}

function calculatePSRAfter(row, bundle, segDataOverride = null) {
  const region = row[3];
  const segment = row[4];
  const teamInfo = row[7] || '';
  const sweepWidth = parseNumeric(row[8]);
  const numSweeps = parseNumeric(row[9]);

  if (!region || !segment || sweepWidth <= 0 || numSweeps <= 0) return '';

  // Extract team members count from e.g. "Team Alpha (3)"
  const match = teamInfo.match(/\((\d+)\)/);
  const numMembers = match ? parseInt(match[1]) : 0;
  if (numMembers <= 0) return '';

  // Get segment info from Segments page data
  const segData = segDataOverride || bundle.pages.page2 || [];
  const segRow = segData.find(r => r[0] === region && r[1] === segment);
  if (!segRow) return '';

  const area = parseNumeric(segRow[2]);
  const length = parseNumeric(segRow[3]);
  const timePerSweep = parseNumeric(segRow[5]);

  if (area <= 0 || length <= 0 || timePerSweep <= 0) return '';

  // Get Consensus and Sum of Areas from Regions page data
  const regionsData = bundle.pages.index;
  const regionRowIndex = (regionsData.rows || []).findIndex(r => r[0] === region);
  if (regionRowIndex === -1) return '';

  const consensus = parseFloat(computeConsensus(regionsData, regionRowIndex)) || 0;
  if (consensus <= 0) return '';

  let sumOfAreas = 0;
  segData.forEach(r => {
    if (r[0] === region) {
      sumOfAreas += parseNumeric(r[2]);
    }
  });
  if (sumOfAreas <= 0) return '';

  // Formula as requested:
  // PSR = Length / TimePerSweep * SweepWidth * ((Consensus * Area / SumOfAreas) - ((Consensus * Area / SumOfAreas) * (1 - EXP(-(SweepWidth / (Area / 640 / Length / NumOfSweeps / NumTeamMembers) * 5280)))))) / (Area / 640)

  const baseValue = (consensus * area / sumOfAreas);
  const z = sweepWidth / ((area / 640 / length / numSweeps / numMembers) * 5280);
  const psrAfter = (length / timePerSweep * sweepWidth * (baseValue - (baseValue * (1 - Math.exp(-z))))) / (area / 640);

  return isFinite(psrAfter) ? psrAfter.toFixed(4) : '';
}

let lastKnownProgress = JSON.parse(sessionStorage.getItem('lastKnownProgress') || '{}');
const updatedTasks = new Set();

function markTaskUpdated(teamName) {
  updatedTasks.add(String(teamName));
  const bundle = loadBundle();
  const assignment = bundle.currentAssignments[teamName] || '';
  const match = assignment.match(/#(\d+)/);
  if (match) {
    updatedTasks.add(String(match[0])); // e.g. "#1"
    updatedTasks.add(String(match[1])); // e.g. "1"
  }
}

function getTaskProgressPercent(status) {
  if (status === 'assigned') return 16.6;
  if (status === 'headed to assignment') return 33.3;
  if (status === 'searching') return 50;
  if (status === 'finished segment') return 66.6;
  if (status === 'returning') return 83.3;
  if (status && status.startsWith('at base')) return 100;
  return 0;
}

function createProgressBar(progress, keyRaw) {
  const key = String(keyRaw);
  const progFill = document.createElement('div');
  progFill.className = 'progress-fill-bg';
  const prev = lastKnownProgress[key] || 0;
  progFill.style.width = prev + '%';
  
  if (updatedTasks.has(key)) {
      progFill.classList.add('animate-progress');
      // Smooth fill from previous level
      setTimeout(() => {
        progFill.style.width = progress + '%';
        if (progress === 100) {
            progFill.classList.add('filling');
            // Wait for fill to finish (0.6s), then turn green, wait 5s, fade out and reset.
            setTimeout(() => {
                progFill.classList.add('completed-green');
                setTimeout(() => {
                    progFill.classList.add('fade-out');
                    setTimeout(() => {
                        progFill.style.width = '0%';
                        progFill.classList.remove('animate-progress', 'filling', 'completed-green', 'fade-out');
                        lastKnownProgress[key] = 0;
                        sessionStorage.setItem('lastKnownProgress', JSON.stringify(lastKnownProgress));
                    }, 1000);
                }, 3000);
            }, 600);
        }
      }, 50);
  } else {
      progFill.style.width = progress + '%';
      if (progress === 100) {
          progFill.classList.add('finished');
      }
  }
  
  if (progress < 100) {
      lastKnownProgress[key] = progress;
      sessionStorage.setItem('lastKnownProgress', JSON.stringify(lastKnownProgress));
  }

  return progFill;
}

function getTeamMembers(teamName) {
  const bundle = loadBundle();
  const data = bundle.pages.page3 || [];
  return data.filter(row => row[1] === teamName && row[6] === 'true');
}

function isParCheckDue(teamName, bundle) {
  const baseTeamNames = ['Base Support', 'Off Duty', 'Command'];
  if (baseTeamNames.includes(teamName)) return false;
  
  const status = bundle.teamStatuses[teamName] || '';
  if (status.startsWith('at base')) return false;

  const lastPar = bundle.parChecks?.[teamName];
  const leaveTime = bundle.teamLeaveTimes?.[teamName];
  const assignTime = bundle.teamAssignmentTimes?.[teamName];
  
  let startTime = 0;
  if (lastPar) startTime = Math.max(startTime, lastPar.lastTime);
  if (leaveTime) startTime = Math.max(startTime, leaveTime);
  if (assignTime) startTime = Math.max(startTime, assignTime);

  if (startTime > 0) {
    const elapsedMs = Date.now() - startTime;
    const freqMs = (bundle.parCheckFrequency || 20) * 60 * 1000;
    return (freqMs - elapsedMs) <= 0;
  }
  return false;
}

function getSegmentInfo(region, segment) {
  const bundle = loadBundle();
  const segData = bundle.pages.page2 || [];
  const row = segData.find(r => r[0] === region && r[1] === segment);
  if (row) {
    return {
      area: row[2],
      length: row[3],
      sweep: row[4],
      time: row[5],
      psr: row[6]
    };
  }
  return { area: '', length: '', sweep: '', time: '', psr: '' };
}

function getNextTaskNumber() {
  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  let max = 0;
  searchLog.forEach(row => {
    if (row[0]) {
      const num = parseInt(row[0].replace('#', ''));
      if (!isNaN(num) && num > max) max = num;
    }
  });
  return max + 1;
}

function addAutoSearchLogEntry(teamName, region, segment) {
  const bundle = loadBundle();
  const logData = bundle.pages.page4 || [];
  const now = new Date();
  const dateStr = `${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}-${now.getFullYear()}`;
  const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
  
  const psrBefore = getLatestPSR(region, segment);
  const teamMembers = getTeamMembers(teamName);
  const teamPillText = `${teamName} (${teamMembers.length})`;
  const segInfo = getSegmentInfo(region, segment);
  const taskNumber = getNextTaskNumber();

  const newRow = [
    `#${taskNumber}`,
    dateStr,
    timeStr,
    region,
    segment,
    psrBefore,
    '', // PSR After
    teamPillText,
    segInfo.sweep,
    '' // Num of Sweeps
  ];
  
  logData.push(newRow);
  bundle.pages.page4 = logData;
  logCreation('Search Log Entry', '#' + taskNumber, bundle);
  saveBundle(bundle);
  return taskNumber;
}

function buildPersonnelActivityTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  const baseTableHead = document.getElementById('base-table-head');
  const baseTableBody = document.getElementById('base-table-body');
  const baseContainer = document.getElementById('base-teams-container-header');
  const searchTeamsContainer = document.getElementById('search-teams-container');
  
  const bundle = loadBundle();
  const data = bundle.pages.page3 || [];

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';
  if (baseTableHead) baseTableHead.innerHTML = '';
  if (baseTableBody) baseTableBody.innerHTML = '';

  const headers = ['Team / Lead', 'Members', 'Assignment / Status', 'Update / Par Check'];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.className = 'fixed-header';
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  if (baseTableHead) {
    const baseHeaderRow = document.createElement('tr');
    ['Team', 'Members'].forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      th.className = 'fixed-header';
      baseHeaderRow.appendChild(th);
    });
    baseTableHead.appendChild(baseHeaderRow);
  }

  const teamsMap = new Map();
  const baseTeamsMap = new Map();
  const baseTeamNames = ['Base Support', 'Off Duty', 'Command'];

  data.forEach((row, idx) => {
    if (row[1] && row[6] === 'true') {
      if (baseTeamNames.includes(row[1])) {
        if (!baseTeamsMap.has(row[1])) baseTeamsMap.set(row[1], []);
        baseTeamsMap.get(row[1]).push(row);
      } else {
        if (!teamsMap.has(row[1])) teamsMap.set(row[1], []);
        teamsMap.get(row[1]).push(row);
      }
    }
  });

  if (baseContainer) {
    baseContainer.style.display = 'block';
  }
  if (searchTeamsContainer) {
    searchTeamsContainer.style.display = 'block';
  }

  const sortedBaseTeams = Array.from(baseTeamsMap.keys()).sort();
  sortedBaseTeams.forEach((teamName, idx) => {
    const members = baseTeamsMap.get(teamName);

    const tr = document.createElement('tr');
    animateNewRow(tr, idx);
    animateArrivedRow(tr, teamName);

    const tdTeam = document.createElement('td');
    tdTeam.dataset.label = 'Team';
    const teamPill = document.createElement('div');
    teamPill.className = 'mini-pill readonly-pill';
    teamPill.textContent = teamName;
    tdTeam.appendChild(teamPill);
    tr.appendChild(tdTeam);

    const tdMembers = document.createElement('td');
    tdMembers.dataset.label = 'Members';
    const membersContainer = document.createElement('div');
    membersContainer.className = 'pill-container';
    members.forEach(member => {
      const pill = document.createElement('div');
      pill.className = 'mini-pill';
      pill.textContent = member[0];
      pill.onclick = () => showReassignPopup(member, teamName);
      membersContainer.appendChild(pill);
    });
    tdMembers.appendChild(membersContainer);
    tr.appendChild(tdMembers);

    if (baseTableBody) baseTableBody.appendChild(tr);
  });

  const sortedTeams = Array.from(teamsMap.keys()).sort();
  sortedTeams.forEach((teamName, idx) => {
    const members = teamsMap.get(teamName);
    const teamLeadRow = members.find(row => row[2] === row[0]) || members[0];
    const teamLead = teamLeadRow ? teamLeadRow[0] : '';

    const tr = document.createElement('tr');
    animateNewRow(tr, idx);
    animateArrivedRow(tr, teamName);

    const tdTeamLead = document.createElement('td');
    tdTeamLead.dataset.label = 'Team / Lead';
    const teamLeadContainer = document.createElement('div');
    teamLeadContainer.className = 'pill-cell-container stacked';

    const teamPill = document.createElement('div');
    teamPill.className = 'mini-pill readonly-pill';
    teamPill.textContent = teamName;
    teamLeadContainer.appendChild(teamPill);

    const teamLeadPill = document.createElement('div');
    teamLeadPill.className = 'mini-pill clickable-pill';
    teamLeadPill.textContent = teamLead || 'No Lead';
    teamLeadPill.onclick = () => {
        if (teamLeadRow) showTeamLeadSwapPopup(teamLeadRow, teamName);
    };
    teamLeadContainer.appendChild(teamLeadPill);
    tdTeamLead.appendChild(teamLeadContainer);
    tr.appendChild(tdTeamLead);

    const tdMembers = document.createElement('td');
    tdMembers.dataset.label = 'Members';
    const membersContainer = document.createElement('div');
    membersContainer.className = 'pill-container';
    members.forEach(member => {
      // Members who are team leads do not need to be in the Members column
      if (member[0] === teamLead) return;

      const pill = document.createElement('div');
      pill.className = 'mini-pill';
      pill.textContent = member[0];
      pill.onclick = () => showReassignPopup(member, teamName);
      membersContainer.appendChild(pill);
    });
    tdMembers.appendChild(membersContainer);
    tr.appendChild(tdMembers);

    const tdAssignStatus = document.createElement('td');
    tdAssignStatus.dataset.label = 'Assignment / Status';
    const assignStatusContainer = document.createElement('div');
    assignStatusContainer.className = 'pill-cell-container stacked';

    const assignmentText = bundle.currentAssignments[teamName] || 'None';
    const assignPill = document.createElement('div');
    assignPill.className = 'mini-pill clickable-pill';
    assignPill.textContent = assignmentText;
    assignPill.onclick = () => showTeamUpdatePopup(teamName);
    assignStatusContainer.appendChild(assignPill);

    const statusPill = document.createElement('div');
    statusPill.className = 'mini-pill clickable-pill';
    const statusText = bundle.teamStatuses[teamName] || '';
    statusPill.textContent = statusText;
    statusPill.onclick = () => showTeamUpdatePopup(teamName);
    
    if (assignmentText !== 'None' && assignmentText !== 'Base' && assignmentText !== '') {
      const progress = getTaskProgressPercent(statusText);
      statusPill.appendChild(createProgressBar(progress, teamName));
    }
    
    assignStatusContainer.appendChild(statusPill);
    tdAssignStatus.appendChild(assignStatusContainer);
    tr.appendChild(tdAssignStatus);

    const tdUpdatePar = document.createElement('td');
    tdUpdatePar.dataset.label = 'Update / Par Check';
    const updateParContainer = document.createElement('div');
    updateParContainer.className = 'pill-cell-container stacked';

    const updatePill = document.createElement('div');
    updatePill.className = 'mini-pill update-pill';
    updatePill.textContent = 'Update';
    updatePill.onclick = () => showTeamUpdatePopup(teamName);
    updateParContainer.appendChild(updatePill);

    const parPill = document.createElement('div');
    parPill.className = 'mini-pill readonly-pill';
    
    const lastPar = bundle.parChecks?.[teamName];
    const leaveTime = bundle.teamLeaveTimes?.[teamName];
    const assignTime = bundle.teamAssignmentTimes?.[teamName];
    
    let startTime = 0;
    if (lastPar) startTime = Math.max(startTime, lastPar.lastTime);
    if (leaveTime) startTime = Math.max(startTime, leaveTime);
    if (assignTime) startTime = Math.max(startTime, assignTime);

    const status = bundle.teamStatuses[teamName] || '';
    if (status.startsWith('at base')) {
      parPill.textContent = ' - ';
    } else if (startTime > 0) {
      const elapsedMs = Date.now() - startTime;
      const freqMs = (bundle.parCheckFrequency || 20) * 60 * 1000;
      const remainingMs = freqMs - elapsedMs;

      if (remainingMs <= 0) {
        parPill.textContent = 'Par Check Due';
        parPill.classList.add('par-check-due');
      } else {
        const remainingMin = Math.ceil(remainingMs / 60000);
        parPill.textContent = `${remainingMin}m`;
      }
    } else {
      parPill.textContent = 'N/A';
    }

    parPill.onclick = () => showParCheckPopup(teamName);
    updateParContainer.appendChild(parPill);
    tdUpdatePar.appendChild(updateParContainer);
    tr.appendChild(tdUpdatePar);

    tableBody.appendChild(tr);
  });

  updateActivityLogUI();
  const existingAddRow = document.querySelector('.add-row-container');
  if (existingAddRow) existingAddRow.remove();
}

function showReassignPopup(member, currentTeam) {
  const teamOptions = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot', 'Golf', 'Command', 'Off Duty', 'Base Support'];
  const popup = createPopup('Reassign ' + member[0]);
  const btnContainer = popup.querySelector('.popup-buttons');
  btnContainer.style.display = 'flex';
  btnContainer.style.flexWrap = 'wrap';
  btnContainer.style.gap = '10px';

  teamOptions.forEach(team => {
    if (team === currentTeam) return;
    const btn = document.createElement('button');
    btn.className = 'mini-pill';
    btn.textContent = 'Move to ' + team;
    btn.onclick = () => {
      const bundle = loadBundle();
      const data = bundle.pages.page3;
      const originalRow = data.find(row => row[0] === member[0]);
      if (originalRow) {
        originalRow[1] = team;
        
        // Auto-switch team lead to new team's lead
        const existingTeamMember = data.find(row => row[1] === team && row[2]);
        originalRow[2] = existingTeamMember ? existingTeamMember[2] : '';
        
        saveBundle(bundle);
        addActivityLogEntry(team, `${member[0]} reassigned from ${currentTeam} to ${team}`);
        refreshCurrentPageTable();
      }
      closePopup(popup);
    };
    btnContainer.appendChild(btn);
  });

  // Leave Scene button in reassign popup
  const leaveBtn = document.createElement('button');
  leaveBtn.className = 'mini-pill';
  leaveBtn.style.borderColor = 'rgba(255, 69, 58, 0.4)';
  leaveBtn.style.color = '#ff453a';
  leaveBtn.textContent = 'Leave Scene';
  leaveBtn.onclick = () => {
    closePopup(popup);
    promptMemberOffScene(member);
  };
  btnContainer.appendChild(leaveBtn);
}

function showTeamLeadSwapPopup(currentLeadRow, teamName) {
  const bundle = loadBundle();
  const data = bundle.pages.page3 || [];
  const teamMembers = data.filter(row => row[1] === teamName && row[0] !== currentLeadRow[0] && row[6] === 'true');
  
  const popup = createPopup(`Change Team Lead for ${teamName}`);
  const btnContainer = popup.querySelector('.popup-buttons');
  btnContainer.style.display = 'flex';
  btnContainer.style.flexWrap = 'wrap';
  btnContainer.style.gap = '10px';

  if (teamMembers.length === 0) {
    const msg = document.createElement('div');
    msg.textContent = "No other members in this team to promote.";
    msg.style.marginBottom = "10px";
    btnContainer.appendChild(msg);
  }

  teamMembers.forEach(member => {
    const btn = document.createElement('button');
    btn.className = 'mini-pill';
    btn.textContent = 'Promote ' + member[0];
    btn.onclick = () => {
      const b = loadBundle();
      const d = b.pages.page3;
      const newLeadName = member[0];
      const oldLeadName = currentLeadRow[0];
      
      // Update all team members to the new lead name
      d.forEach(row => {
        if (row[1] === teamName) {
          row[2] = newLeadName;
        }
      });
      
      saveBundle(b);
      addActivityLogEntry(teamName, `${newLeadName} is now Team Lead (previously ${oldLeadName})`);
      closePopup(popup);
      refreshCurrentPageTable();
    };
    btnContainer.appendChild(btn);
  });
}

function getStatusIndex(status) {
    if (status === 'assigned') return -1;
    if (status === 'headed to assignment') return 0;
    if (status === 'searching') return 1;
    if (status === 'finished segment') return 2;
    if (status === 'returning') return 3;
    if (status && (status === 'at base' || status.startsWith('at base'))) return 4;
    return -1;
}

function showTeamUpdatePopup(teamName) {
  const popup = createPopup('Team ' + teamName + ' Update');
  const btnContainer = popup.querySelector('.popup-buttons');

  const bundle = loadBundle();
  const currentStatus = bundle.teamStatuses[teamName] || '';
  const currentIndex = getStatusIndex(currentStatus);

  const updateStatus = (newStatus, logAction) => {
    showMissingStepsPopup(teamName, newStatus, () => {
      markTaskUpdated(teamName);
      const b = loadBundle();
      b.teamStatuses[teamName] = newStatus;
      if (newStatus === 'headed to assignment') {
        b.teamLeaveTimes[teamName] = Date.now();
      }
      if (!b.parChecks) b.parChecks = {};
      b.parChecks[teamName] = { lastTime: Date.now() };
      saveBundle(b);
      addActivityLogEntry(teamName, logAction);
      popup.remove();
      refreshCurrentPageTable();
    });
  };

  const statusActions = [
    { id: 'headed to assignment', label: 'Leave Base', action: () => updateStatus('headed to assignment', 'Leaving base for assignment') },
    { id: 'searching', label: 'Begin Assignment', action: () => updateStatus('searching', 'Beginning assignment') },
    { id: 'finished segment', label: 'Finish Assignment', action: () => updateStatus('finished segment', 'Finished assignment') },
    { id: 'returning', label: 'Return to Base', action: () => updateStatus('returning', 'Returning to base') },
    { id: 'at base', label: 'Arrived at Base', action: () => {
        showMissingStepsPopup(teamName, 'at base', () => {
            markTaskUpdated(teamName);
            const b = loadBundle();
            const now = new Date();
            const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
            b.teamStatuses[teamName] = `at base (${timeStr})`;
            b.currentAssignments[teamName] = 'Base';
            b.teamAssignmentTimes[teamName] = Date.now();
            
            if (b.parChecks) delete b.parChecks[teamName];
            if (b.teamLeaveTimes) delete b.teamLeaveTimes[teamName];
            if (b.teamAssignmentTimes) delete b.teamAssignmentTimes[teamName];
            
            if (!b.arrivedTeams) b.arrivedTeams = [];
            if (!b.arrivedTeams.includes(teamName)) b.arrivedTeams.push(teamName);
            
            saveBundle(b);
            addActivityLogEntry(teamName, `Arrived at base at ${timeStr}`);
            popup.remove();
            refreshCurrentPageTable();
        });
    }}
  ];

  // Assign New Task (Always available)
  const btnNew = document.createElement('button');
  btnNew.className = 'popup-btn';
  btnNew.textContent = 'Assign New Task';
  btnNew.onclick = () => {
    showNewSegmentPopup(teamName, popup);
  };
  btnContainer.appendChild(btnNew);

  // Status Buttons
  statusActions.forEach((step, idx) => {
    const btn = document.createElement('button');
    btn.className = 'popup-btn';
    btn.textContent = step.label;
    if (idx <= currentIndex) {
      btn.style.opacity = '0.4';
      btn.style.pointerEvents = 'none';
    }
    btn.onclick = step.action;
    btnContainer.appendChild(btn);
  });

  // Par Check
  const btnPar = document.createElement('button');
  btnPar.className = 'popup-btn';
  if (isParCheckDue(teamName, bundle)) {
    btnPar.textContent = 'Par Check Due';
    btnPar.classList.add('par-check-due');
  } else {
    btnPar.textContent = 'Par Check';
  }
  btnPar.onclick = () => {
    showParCheckPopup(teamName, popup);
  };
  btnContainer.appendChild(btnPar);
}

function refreshCurrentPageTable() {
  if (isPersonnelPage()) buildPersonnelTable();
  else if (isSearchLogPage()) buildSearchLogTable();
  else if (isSegmentsPage()) buildSegmentsTable();
  else if (isFormsPage()) buildFormsPage();
  else if (isRegionsPage()) buildRegionsTable();
  else if (isHomePage()) buildHomePage();

  // Ensure header status (like par checks) is updated immediately, but avoid infinite loop
  checkParChecksAndNotify(true);
}

function showParCheckPopup(teamName, parentPopup) {
  if (parentPopup) closePopup(parentPopup);
  const popup = createPopup('Par Check - Team ' + teamName, null, () => refreshCurrentPageTable());
  const content = popup.querySelector('.popup-content');
  
  const inputContainer = document.createElement('div');
  inputContainer.className = 'popup-input-container';
  
  const textarea = document.createElement('textarea');
  textarea.className = 'popup-textarea';
  textarea.placeholder = 'Enter par check update...';
  inputContainer.appendChild(textarea);
  
  const updateBtn = document.createElement('button');
  updateBtn.className = 'popup-btn primary';
  updateBtn.textContent = 'Update';
  updateBtn.onclick = () => {
    const bundle = loadBundle();
    const now = new Date();
    const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    const fullNote = `[${timeStr}] Par Check: ${textarea.value}`;
    bundle.parChecks[teamName] = {
      lastTime: Date.now(),
      lastNote: textarea.value
    };
    saveBundle(bundle);
    addActivityLogEntry(teamName, fullNote);
    closePopup(popup);
    refreshCurrentPageTable();
  };
  inputContainer.appendChild(updateBtn);
  
  content.appendChild(inputContainer);
}

function showNewSegmentPopup(teamName, parentPopup) {
  if (parentPopup) parentPopup.remove();
  const popup = createPopup('Assign New Task - Team ' + teamName);
  const content = popup.querySelector('.popup-content');
  
  const bundle = loadBundle();
  const segments = (bundle.pages.page2 || []).filter(s => s[0] && s[1]);
  
  // Sort by PSR descending
  segments.sort((a, b) => {
    const psrA = parseFloat(getLatestPSR(a[0], a[1])) || 0;
    const psrB = parseFloat(getLatestPSR(b[0], b[1])) || 0;
    return psrB - psrA;
  });

  if (segments.length === 0) {
    const p = document.createElement('p');
    p.textContent = 'No segments defined.';
    content.appendChild(p);
  } else {
    const segmentsGrid = document.createElement('div');
    segmentsGrid.className = 'popup-segments-grid';
    
    segments.forEach(seg => {
      const region = seg[0];
      const segment = seg[1];
      const psr = getLatestPSR(region, segment);
      const val = `${region} - ${segment}`;
      
      const btn = document.createElement('button');
      btn.className = 'mini-pill';
      btn.style.padding = '8px 12px';
      btn.style.textAlign = 'center';
      btn.style.display = 'flex';
      btn.style.flexDirection = 'column';
      btn.style.justifyContent = 'center';
      btn.style.alignItems = 'center';
      btn.style.gap = '2px';
      btn.style.width = '100%';
      btn.style.cursor = 'pointer';
      btn.style.fontSize = '0.85rem';

      const title = document.createElement('div');
      title.style.fontWeight = '700';
      title.textContent = val;
      
      const details = document.createElement('div');
      details.style.fontSize = '0.7rem';
      details.style.opacity = '0.8';
      details.textContent = `PSR: ${psr || 'N/A'}`;
      
      btn.appendChild(title);
      btn.appendChild(details);

      // Highlight if currently being searched
      const isSearching = Object.entries(bundle.currentAssignments).some(([t, ass]) => {
          if (!ass.includes(val)) return false;
          const status = bundle.teamStatuses[t] || '';
          return !status.includes('finished segment') && !status.startsWith('at base');
      });

      if (isSearching) {
          btn.style.background = 'rgba(235, 87, 87, 0.25)';
          btn.style.borderColor = 'rgba(235, 87, 87, 0.5)';
          btn.style.color = '#ff4d4d';
      }

      btn.onclick = () => {
        showMissingStepsPopup(teamName, null, () => {
          const taskNumber = addAutoSearchLogEntry(teamName, region, segment);
          const fullAssignment = `#${taskNumber} ${val}`;
          
          const b2 = loadBundle();
          b2.currentAssignments[teamName] = fullAssignment;
          b2.teamAssignmentTimes[teamName] = Date.now();
          b2.teamStatuses[teamName] = 'assigned';
          if (!b2.parChecks) b2.parChecks = {};
          b2.parChecks[teamName] = { lastTime: Date.now() };
          saveBundle(b2);
          markTaskUpdated(teamName);
          addActivityLogEntry(teamName, 'Assigned to segment: ' + fullAssignment);
          closePopup(popup);
          refreshCurrentPageTable();
        });
      };
      segmentsGrid.appendChild(btn);
    });
    content.appendChild(segmentsGrid);
  }
}

function showTeamSelectionPopup(onTeamSelected) {
  const bundle = loadBundle();
  const data = bundle.pages.page3 || [];
  const teamsMap = new Map();
  data.forEach(row => {
    if (row[1] && row[6] === 'true') {
      teamsMap.set(row[1], true);
    }
  });
  const sortedTeams = Array.from(teamsMap.keys())
    .filter(team => team !== 'Base Support' && team !== 'Command')
    .sort();

  const popup = createPopup('Select Team for Search');
  const content = popup.querySelector('.popup-content');
  
  const segmentsGrid = document.createElement('div');
  segmentsGrid.className = 'popup-segments-grid';
  
  if (sortedTeams.length === 0) {
    const p = document.createElement('p');
    p.textContent = 'No teams currently on scene.';
    segmentsGrid.appendChild(p);
  }

  sortedTeams.forEach(team => {
    const status = bundle.teamStatuses[team] || '';
    const assignment = bundle.currentAssignments[team] || '';
    
    const btn = document.createElement('button');
    btn.className = 'mini-pill';
    btn.style.width = '100%';
    btn.style.padding = '12px';
    btn.style.display = 'flex';
    btn.style.flexDirection = 'column';
    btn.style.alignItems = 'center';
    btn.style.gap = '4px';
    
    const nameDiv = document.createElement('div');
    nameDiv.style.fontWeight = 'bold';
    nameDiv.textContent = 'Team ' + team;
    btn.appendChild(nameDiv);

    if (status && status.startsWith('at base')) {
      btn.style.background = 'rgba(255, 255, 255, 0.05)';
      const pill = document.createElement('div');
      pill.style.fontSize = '0.75rem';
      pill.style.opacity = '0.7';
      pill.textContent = status;
      btn.appendChild(pill);
    } else if (status !== '') {
      btn.style.background = 'rgba(125, 198, 255, 0.15)';
      btn.style.borderColor = 'rgba(125, 198, 255, 0.4)';
      const match = assignment.match(/#\d+/);
      if (match) {
        const pill = document.createElement('div');
        pill.style.fontSize = '0.75rem';
        pill.style.color = 'var(--accent)';
        pill.textContent = 'Task ' + match[0];
        btn.appendChild(pill);
      }
    }

    btn.onclick = () => {
      onTeamSelected(team);
      closePopup(popup);
    };
    segmentsGrid.appendChild(btn);
  });
  content.appendChild(segmentsGrid);
}

function showMissingStepsPopup(teamName, targetStatus, onComplete) {
  const bundle = loadBundle();
  const currentStatus = bundle.teamStatuses[teamName] || '';
  
  const sequence = [
    { id: 'headed to assignment', label: 'Leave Base', log: 'Leaving base for assignment' },
    { id: 'searching', label: 'Begin Assignment', log: 'Beginning assignment' },
    { id: 'finished segment', label: 'Finish Assignment', log: 'Finished assignment' },
    { id: 'returning', label: 'Return to Base', log: 'Returning to base' },
    { id: 'at base', label: 'Arrived at Base', log: 'Arrived at base' }
  ];

  const currentIndex = getStatusIndex(currentStatus);
  const targetIndex = getStatusIndex(targetStatus);
  
  const missingSteps = sequence.slice(currentIndex + 1, targetIndex);
  
  if (missingSteps.length === 0) {
    onComplete();
    return;
  }
  
  const popup = createPopup('Missing Steps: ' + teamName);
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');
  
  const msg = document.createElement('p');
  msg.textContent = 'Please provide missing step information:';
  msg.style.marginBottom = '15px';
  content.insertBefore(msg, btnContainer);
  
  const list = document.createElement('div');
  list.style.textAlign = 'left';
  list.style.marginBottom = '20px';
  list.style.display = 'flex';
  list.style.flexDirection = 'column';
  list.style.gap = '15px';
  list.style.maxHeight = '300px';
  list.style.overflowY = 'auto';
  list.style.paddingRight = '5px';
  
  const stepData = [];
  
  missingSteps.forEach((step, i) => {
    const item = document.createElement('div');
    item.style.display = 'flex';
    item.style.alignItems = 'center';
    item.style.justifyContent = 'space-between';
    item.style.gap = '15px';
    
    // Step Pill Button
    const stepPill = document.createElement('div');
    stepPill.className = 'popup-btn';
    stepPill.textContent = step.label;
    stepPill.style.flex = '1';
    stepPill.style.padding = '8px 16px';
    stepPill.style.fontSize = '0.9rem';
    stepPill.style.cursor = 'default';
    stepPill.style.textAlign = 'center';
    item.appendChild(stepPill);

    // Date/Time Stamp
    const now = new Date();
    const initialDate = `${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}-${now.getFullYear()}`;
    const initialTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    
    const stamp = document.createElement('div');
    stamp.className = 'sub-nav-btn';
    stamp.style.padding = '8px 12px';
    stamp.style.fontSize = '0.85rem';
    stamp.style.whiteSpace = 'nowrap';
    stamp.style.cursor = 'pointer';
    stamp.textContent = `${initialDate} ${initialTime}`;
    
    const currentStepData = { step, date: initialDate, time: initialTime };
    stepData.push(currentStepData);

    stamp.onclick = () => {
        showEditTimePopup(currentStepData, (newData) => {
            stamp.textContent = `${newData.date} ${newData.time}`;
        });
    };

    item.appendChild(stamp);
    list.appendChild(item);
  });
  
  content.insertBefore(list, btnContainer);
  
  const submitBtn = document.createElement('button');
  submitBtn.className = 'popup-btn primary';
  submitBtn.textContent = 'Submit';
  submitBtn.onclick = () => {
    markTaskUpdated(teamName);
    const b = loadBundle();
    stepData.forEach(item => {
        let logText = item.step.log;
        const d = item.date;
        const t = item.time;

        if (item.step.id === 'at base') {
          b.teamStatuses[teamName] = `at base (${t})`;
          b.currentAssignments[teamName] = 'Base';
          b.teamAssignmentTimes[teamName] = Date.now();
        } else {
          b.teamStatuses[teamName] = item.step.id;
          if (item.step.id === 'headed to assignment') {
            b.teamLeaveTimes[teamName] = Date.now();
          }
        }

        if (item.step.id === 'finished segment') {
            const assignment = b.currentAssignments[teamName] || '';
            const match = assignment.match(/#\d+/);
            if (match) {
                const tag = match[0];
                const searchLog = b.pages.page4 || [];
                const logEntry = searchLog.find(r => r[0] === tag);
                if (logEntry) {
                    logEntry[1] = d;
                    logEntry[2] = t;
                }
            }
        }

        addActivityLogEntry(teamName, logText, b, null, d, t);
    });
    saveBundle(b);
    closePopup(popup);
    onComplete();
  };
  btnContainer.appendChild(submitBtn);
}

function showEditTimePopup(data, onSave) {
  const popup = createPopup('Edit Time');
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');
  
  const inputContainer = document.createElement('div');
  inputContainer.style.display = 'flex';
  inputContainer.style.flexDirection = 'column';
  inputContainer.style.gap = '15px';
  inputContainer.style.marginBottom = '20px';
  inputContainer.style.marginTop = '10px';
  
  const dateInput = document.createElement('input');
  dateInput.className = 'pill-input';
  dateInput.value = data.date;
  dateInput.placeholder = 'MM-DD-YYYY';
  setupAutoFormatDate(dateInput);
  
  const timeInput = document.createElement('input');
  timeInput.className = 'pill-input';
  timeInput.value = data.time;
  timeInput.placeholder = 'hh:mm';
  setupAutoFormatTime(timeInput);
  
  inputContainer.appendChild(dateInput);
  inputContainer.appendChild(timeInput);
  content.insertBefore(inputContainer, btnContainer);
  
  const saveBtn = document.createElement('button');
  saveBtn.className = 'popup-btn primary';
  saveBtn.textContent = 'Save';
  saveBtn.onclick = () => {
    data.date = dateInput.value;
    data.time = timeInput.value;
    onSave(data);
    closePopup(popup);
  };
  btnContainer.appendChild(saveBtn);
}

function callAllTeamsToBase() {
  const commander = prompt("Enter the name of the person commanding this to all teams:");
  if (!commander) return;

  const bundle = loadBundle();
  const searchTeams = [];
  const personnel = bundle.pages.page3 || [];
  personnel.forEach(row => {
    const team = row[1];
    const onScene = row[6] === 'true';
    if (team && onScene && !['Base Support', 'Off Duty', 'Command'].includes(team)) {
      if (!searchTeams.includes(team)) searchTeams.push(team);
    }
  });

  if (searchTeams.length === 0) {
    alert("No search teams currently on scene.");
    return;
  }

  const sequence = [
    { id: 'headed to assignment', label: 'Leave Base', log: 'Leaving base for assignment' },
    { id: 'searching', label: 'Beginning Assignment', log: 'Beginning assignment' },
    { id: 'finished segment', label: 'Finish Assignment', log: 'Finished assignment' },
    { id: 'returning', label: 'Return to Base', log: 'Returning to base' }
  ];

  function getStatusIndex(status) {
    if (status === 'assigned') return -1;
    if (status === 'headed to assignment') return 0;
    if (status === 'searching') return 1;
    if (status === 'finished segment') return 2;
    if (status === 'returning') return 3;
    if (status && status.startsWith('at base')) return 4;
    return -1;
  }

  searchTeams.forEach(teamName => {
    const currentStatus = bundle.teamStatuses[teamName] || '';
    const currentIndex = getStatusIndex(currentStatus);
    const targetIndex = 3; // 'returning'

    if (currentIndex < targetIndex) {
      // Complete all steps in between
      for (let i = currentIndex + 1; i <= targetIndex; i++) {
        const step = sequence[i];
        bundle.teamStatuses[teamName] = step.id;
        if (!bundle.parChecks) bundle.parChecks = {};
        bundle.parChecks[teamName] = { lastTime: Date.now() };
        if (step.id === 'headed to assignment') {
          bundle.teamLeaveTimes[teamName] = Date.now();
        }
        addActivityLogEntry(teamName, `${step.log} (Commanded by: ${commander})`, bundle);
      }
      if (!bundle.arrivedTeams) bundle.arrivedTeams = [];
      if (!bundle.arrivedTeams.includes(teamName)) bundle.arrivedTeams.push(teamName);
    }
  });

  saveBundle(bundle);
  refreshCurrentPageTable();
}

function closePopup(overlay) {
  if (!overlay) return;
  overlay.classList.add('fade-out');
  const content = overlay.querySelector('.popup-content');
  if (content) content.classList.add('fade-out');
  setTimeout(() => overlay.remove(), 200);
}

function createPopup(titleText, originElement = null, onClose = null) {
  const overlay = document.createElement('div');
  overlay.className = 'popup-overlay';
  
  const content = document.createElement('div');
  content.className = 'popup-content expanding';
  
  // Close button (x)
  const closeBtn = document.createElement('button');
  closeBtn.className = 'popup-close-btn';
  closeBtn.innerHTML = '✕';
  closeBtn.onclick = () => {
    if (onClose) onClose();
    closePopup(overlay);
  };
  content.appendChild(closeBtn);
  
  const origin = originElement || document.activeElement;
  if (origin && typeof origin.getBoundingClientRect === 'function') {
    const rect = origin.getBoundingClientRect();
    if (rect.width > 0 && rect.height > 0) {
        const x = rect.left + rect.width / 2;
        const y = rect.top + rect.height / 2;
        content.style.transformOrigin = `${x}px ${y}px`;
    }
  }
  
  const title = document.createElement('div');
  title.className = 'popup-title';
  title.textContent = titleText;
  content.appendChild(title);
  
  const btnContainer = document.createElement('div');
  btnContainer.className = 'popup-buttons';
  content.appendChild(btnContainer);
  
  overlay.appendChild(content);
  document.body.appendChild(overlay);
  return overlay;
}

function addActivityLogEntry(team, action, bundle = null, membersOverride = null, customDate = null, customTime = null) {
  const b = bundle || loadBundle();
  const now = new Date();
  const dateStr = customDate || `${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}-${now.getFullYear()}`;
  const timeStr = customTime || `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
  
  // Determine tag: "base" or "#task"
  let tag = 'base';
  const assignment = b.currentAssignments[team] || '';
  if (assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
    const match = assignment.match(/#\d+/);
    if (match) {
      tag = match[0];
    }
  }

  const currentUser = getCurrentUser();
  const userTag = currentUser ? ` - ${currentUser.handle || (currentUser.firstName + ' ' + (currentUser.lastName || '')).trim()}` : '';

  const members = membersOverride || getTeamMembers(team).map(m => {
    const name = m[0];
    const leadName = m[2];
    const isLead = (leadName === name);
    return isLead ? name + '*' : name;
  }).join(', ');

  const ts = (customDate && customTime) ? 
    new Date(`${customDate.split('-')[2]}-${customDate.split('-')[0]}-${customDate.split('-')[1]}T${customTime}:00`).getTime() :
    Date.now();

  b.activityLog.unshift({
    id: 'log-' + ts + '-' + Math.floor(Math.random() * 1000),
    date: dateStr,
    time: timeStr,
    tag: tag + userTag,
    team: team,
    members: members,
    action: action,
    timestamp: ts
  });
  
  if (!bundle) {
    saveBundle(b);
  }
  if (isHomePage()) {
    buildHomePage();
  }
}

function logDeletion(type, name, bundle = null) {
  addActivityLogEntry('System', `Deleted ${type}: ${name || 'unknown'}`, bundle);
}

function logCreation(type, name, bundle = null) {
  addActivityLogEntry('System', `Created ${type}: ${name || 'unknown'}`, bundle);
}

function showLoginPopup() {
  const onCancel = () => {
    if (!getCurrentUser()) {
        alert('You must select a team member to continue');
        showLoginPopup();
    }
  };
  const popup = createPopup('Select Team Member', null, onCancel);
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');
  
  const bundle = loadBundle();
  const accounts = bundle.accounts || [];
  let selectedUser = getCurrentUser();

  const inputs = document.createElement('div');
  inputs.className = 'popup-input-container';
  inputs.style.flexDirection = 'column';
  inputs.style.gap = '15px';
  
  const userSelectContainer = document.createElement('div');
  userSelectContainer.style.width = '100%';
  userSelectContainer.style.display = 'flex';
  userSelectContainer.style.flexDirection = 'column';
  userSelectContainer.style.alignItems = 'center';

  const userSearchInput = document.createElement('input');
  userSearchInput.type = 'text';
  userSearchInput.placeholder = 'Type to search...';
  userSearchInput.className = 'pill-input';
  userSearchInput.style.textAlign = 'center';
  userSearchInput.value = '';
  userSelectContainer.appendChild(userSearchInput);

  const pillsContainer = document.createElement('div');
  pillsContainer.style.display = 'flex';
  pillsContainer.style.flexWrap = 'wrap';
  pillsContainer.style.justifyContent = 'center';
  pillsContainer.style.gap = '5px';
  pillsContainer.style.marginTop = '10px';
  userSelectContainer.appendChild(pillsContainer);

  const updatePills = () => {
    pillsContainer.innerHTML = '';
    const query = userSearchInput.value.toLowerCase();
    const filtered = accounts.filter(acc => 
      `${acc.firstName} ${acc.lastName}`.toLowerCase().includes(query)
    );

    filtered.forEach(acc => {
      const pill = document.createElement('button');
      pill.className = 'mini-pill';
      pill.textContent = `${acc.firstName} ${acc.lastName}`;
      if (selectedUser && selectedUser.firstName === acc.firstName && selectedUser.lastName === acc.lastName) {
          pill.style.background = 'var(--pill-bg-hover)';
          pill.style.borderColor = 'var(--accent)';
      }
      pill.onclick = () => {
        setCurrentUser(acc);
        closePopup(popup);
        window.location.reload();
      };
      pillsContainer.appendChild(pill);
    });
  };

  userSearchInput.oninput = updatePills;
  updatePills();
  inputs.appendChild(userSelectContainer);
  
  content.insertBefore(inputs, btnContainer);
  
  // Remove the Login button from btnContainer as selection immediately switches user
  btnContainer.innerHTML = '';
  
  userSearchInput.focus();
}

function showAccountManager() {
  const user = getCurrentUser();
  if (!user) return;

  const popup = createPopup('Account Manager');
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');
  content.style.maxWidth = '600px';

  const bundle = loadBundle();
  const accounts = bundle.accounts || [];

  const list = document.createElement('div');
  list.style.maxHeight = '300px';
  list.style.overflowY = 'auto';
  list.style.marginBottom = '20px';

  accounts.forEach((acc, idx) => {
    const row = document.createElement('div');
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.justifyContent = 'space-between';
    row.style.padding = '10px';
    row.style.borderBottom = '1px solid var(--line)';

    const info = document.createElement('div');
    info.textContent = `${acc.firstName} ${acc.lastName} (@${acc.handle || 'no-handle'})`;
    row.appendChild(info);

    const actions = document.createElement('div');
    actions.style.display = 'flex';
    actions.style.gap = '5px';

    const editBtn = document.createElement('button');
    editBtn.className = 'mini-pill';
    editBtn.textContent = 'Edit';
    editBtn.onclick = () => {
        popup.remove();
        showEditAccountPopup(acc, idx);
    };
    actions.appendChild(editBtn);

    if (!isUserAdmin(acc)) {
        const delBtn = document.createElement('button');
        delBtn.className = 'mini-pill';
        delBtn.style.color = 'var(--accent)';
        delBtn.textContent = 'Delete';
        delBtn.onclick = () => {
            const b = loadBundle();
            const doDelete = () => {
                bundle.accounts.splice(idx, 1);
                saveBundle(bundle);
                popup.remove();
                showAccountManager();
            };
            if (b.deleteMode) {
                doDelete();
            } else if (confirm(`Delete account ${acc.firstName}?`)) {
                doDelete();
            }
        };
        actions.appendChild(delBtn);
    }
    row.appendChild(actions);
    list.appendChild(row);
  });

  content.insertBefore(list, btnContainer);

  const addBtn = document.createElement('button');
  addBtn.className = 'popup-btn primary';
  addBtn.textContent = '+ Create New Account';
  addBtn.onclick = () => {
    popup.remove();
    showEditAccountPopup(null);
  };
  btnContainer.appendChild(addBtn);

  const logoutBtn = document.createElement('button');
  logoutBtn.className = 'popup-btn';
  logoutBtn.style.marginTop = '10px';
  logoutBtn.textContent = 'Logout';
  logoutBtn.onclick = () => {
      setCurrentUser(null);
      window.location.reload();
  };
  btnContainer.appendChild(logoutBtn);
}

function showEditAccountPopup(acc, index = -1) {
    const popup = createPopup(acc ? 'Edit Account' : 'Create Account', null, () => showAccountManager());
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');

    const inputs = document.createElement('div');
    inputs.className = 'popup-input-container';

    const fName = document.createElement('input');
    fName.className = 'pill-input';
    fName.placeholder = 'First Name';
    fName.value = acc ? acc.firstName : '';
    inputs.appendChild(fName);

    const lName = document.createElement('input');
    lName.className = 'pill-input';
    lName.placeholder = 'Last Name';
    lName.value = acc ? acc.lastName : '';
    inputs.appendChild(lName);

    const handle = document.createElement('input');
    handle.className = 'pill-input';
    handle.placeholder = 'Handle (for Log Activity)';
    handle.value = acc ? acc.handle || '' : '';
    inputs.appendChild(handle);

    const pinLabel = document.createElement('div');
    pinLabel.textContent = 'Account ID (formerly PIN):';
    pinLabel.style.marginTop = '10px';
    pinLabel.style.fontSize = '0.8rem';
    pinLabel.style.color = 'var(--muted)';
    inputs.appendChild(pinLabel);

    const pin = document.createElement('input');
    pin.className = 'pill-input';
    pin.placeholder = 'Account ID';
    pin.value = acc ? acc.pin : '';
    pin.readOnly = true; // Make it read-only as we don't want people worrying about PINs
    inputs.appendChild(pin);

    const color = document.createElement('input');
    color.type = 'color';
    color.style.width = '100%';
    color.style.height = '40px';
    color.style.border = 'none';
    color.style.background = 'transparent';
    color.value = acc ? acc.color : '#7dc6ff';
    inputs.appendChild(color);

    const pagesLabel = document.createElement('div');
    pagesLabel.textContent = 'Visible Pages:';
    pagesLabel.style.marginTop = '10px';
    pagesLabel.style.fontWeight = '700';
    inputs.appendChild(pagesLabel);

    const pagesContainer = document.createElement('div');
    pagesContainer.style.display = 'grid';
    pagesContainer.style.gridTemplateColumns = '1fr 1fr';
    pagesContainer.style.gap = '5px';

    const allPages = [
        { id: 'index', name: 'Regions' },
        { id: 'page2', name: 'Segments' },
        { id: 'page3', name: 'Personnel' },
        { id: 'sub-activity', name: 'Personnel: Activity' },
        { id: 'sub-team-reports', name: 'Personnel: Team Reports' },
        { id: 'sub-member-reports', name: 'Personnel: Member Reports' },
        { id: 'sub-all-members', name: 'Personnel: All Members' },
        { id: 'page4', name: 'Search Log' },
        { id: 'page5', name: 'Forms' },
        { id: 'page6', name: 'Profile' },
        { id: 'page7', name: 'Uploads' },
        { id: 'page10', name: 'Maps' },
        { id: 'settings', name: 'Settings' }
    ];

    allPages.forEach(p => {
        const wrap = document.createElement('label');
        wrap.style.display = 'flex';
        wrap.style.alignItems = 'center';
        wrap.style.gap = '5px';
        wrap.style.fontSize = '0.9rem';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = p.id;
        cb.checked = acc ? acc.visiblePages.includes(p.id) : true;
        wrap.appendChild(cb);
        wrap.appendChild(document.createTextNode(p.name));
        pagesContainer.appendChild(wrap);
    });
    inputs.appendChild(pagesContainer);

    content.insertBefore(inputs, btnContainer);

    const saveBtn = document.createElement('button');
    saveBtn.className = 'popup-btn primary';
    saveBtn.textContent = 'Save';
    saveBtn.onclick = () => {
        const bundle = loadBundle();
        let finalPin = pin.value;
        if (!finalPin) {
            // Generate next sequential PIN if not provided
            let next = 1400;
            while (bundle.accounts.some(a => a.pin === next.toString())) {
                next++;
            }
            finalPin = next.toString();
        }
        const newAcc = {
            firstName: fName.value,
            lastName: lName.value,
            handle: handle.value,
            pin: finalPin,
            color: color.value,
            visiblePages: Array.from(pagesContainer.querySelectorAll('input:checked')).map(i => i.value).concat(['home'])
        };
        
        // Sync to Personnel list
        if (bundle.pages && bundle.pages.page3) {
            const newName = newAcc.handle || (newAcc.firstName + ' ' + (newAcc.lastName || '')).trim();
            if (acc) {
                const oldPin = acc.pin;
                const oldHandle = acc.handle;
                const oldFullName = (acc.firstName + ' ' + (acc.lastName || '')).trim();
                bundle.pages.page3.forEach(row => {
                    const rowName = (row[0] || '').trim();
                    const rowPin = (row[8] || '').trim();
                    if ((rowPin && rowPin === oldPin) || rowName === oldHandle || rowName === oldFullName) {
                        row[0] = newName;
                        row[8] = newAcc.pin;
                    }
                });
            } else {
                // New account - add to Personnel if not already there
                const exists = bundle.pages.page3.some(row => (row[0] || '').trim() === newName);
                if (!exists) {
                    bundle.pages.page3.push([newName, '', '', '', '', '', '', '', newAcc.pin]);
                } else {
                    // Link to existing Personnel row
                    bundle.pages.page3.forEach(row => {
                        if ((row[0] || '').trim() === newName) {
                            row[8] = newAcc.pin;
                        }
                    });
                }
            }
        }

        if (index >= 0) {
            bundle.accounts[index] = newAcc;
        } else {
            bundle.accounts.push(newAcc);
        }
        saveBundle(bundle);
        closePopup(popup);
        showAccountManager();
    };
    btnContainer.appendChild(saveBtn);
}

function updateHeaderProfile() {
    const user = getCurrentUser();
    const btn = document.getElementById('profile-btn');
    if (!btn) return;

    // Reset classes
    btn.className = 'profile-nav-btn';
    btn.style.background = '';
    btn.style.color = '';
    btn.style.borderColor = '';

    if (user) {
        const f = (user.firstName || '').trim().charAt(0).toUpperCase();
        const l = (user.lastName || '').trim().charAt(0).toUpperCase();
        btn.innerHTML = (f + l) || '??';
        
        if (pageKey() === 'page8') {
            btn.classList.add('active');
        }
        
        if (user.color && user.color !== 'none') {
            btn.classList.add(`profile-highlight-${user.color}`);
        } else {
            btn.style.background = 'transparent';
            btn.style.color = 'var(--text)';
        }

        btn.onclick = (e) => {
            e.preventDefault();
            navigateToPage('page8.html');
        };
    } else {
        btn.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>`;
        btn.onclick = (e) => {
            e.preventDefault();
            showLoginPopup();
        };
    }

    // Hide restricted nav items
    const navItems = document.querySelectorAll('nav a');
    navItems.forEach(item => {
        const href = item.getAttribute('href');
        if (!href) return;
        const key = href.replace('.html', '');
        
        // Always show these
        if (key === 'home' || key === 'page8' || key === 'page10' || href === '#') return;
        
        // Page 9 is gone, now part of Page 8
        if (key === 'page9') {
            item.style.display = 'none';
            return;
        }
        
        if (user && user.visiblePages && !user.visiblePages.includes(key)) {
            item.style.display = 'none';
        }
    });
}

function showProfileSettingsPopup(user) {
    const popup = createPopup('Profile Settings');
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');

    const inputs = document.createElement('div');
    inputs.className = 'popup-input-container';

    const colorLabel = document.createElement('div');
    colorLabel.textContent = 'Profile Color:';
    colorLabel.style.fontWeight = '700';
    inputs.appendChild(colorLabel);

    const color = document.createElement('input');
    color.type = 'color';
    color.style.width = '100%';
    color.style.height = '40px';
    color.style.border = 'none';
    color.style.background = 'transparent';
    color.value = user.color || '#7dc6ff';
    inputs.appendChild(color);

    content.insertBefore(inputs, btnContainer);

    const saveBtn = document.createElement('button');
    saveBtn.className = 'popup-btn primary';
    saveBtn.textContent = 'Save';
    saveBtn.onclick = () => {
        const bundle = loadBundle();
        const accountIndex = bundle.accounts.findIndex(a => a.pin === user.pin);
        if (accountIndex >= 0) {
            bundle.accounts[accountIndex].color = color.value;
            saveBundle(bundle);
            user.color = color.value;
            setCurrentUser(user);
            updateHeaderProfile();
        }
        popup.remove();
    };
    btnContainer.appendChild(saveBtn);

    const logoutBtn = document.createElement('button');
    logoutBtn.className = 'popup-btn';
    logoutBtn.style.marginTop = '10px';
    logoutBtn.textContent = 'Logout';
    logoutBtn.onclick = () => {
        setCurrentUser(null);
        window.location.reload();
    };
    btnContainer.appendChild(logoutBtn);
}

let currentLogTeamFilter = null;

function updateActivityLogUI() {
  const logBox = document.getElementById('activity-log');
  const searchInput = document.getElementById('log-search');
  const teamFilters = document.getElementById('log-team-filters');
  if (!logBox) return;
  
  const bundle = loadBundle();
  logBox.innerHTML = '';
  
  if (teamFilters) {
      teamFilters.innerHTML = '';
      const teams = Array.from(new Set(bundle.activityLog.map(e => e.team).filter(t => t && t !== 'System'))).sort();
      
      const allBtn = document.createElement('button');
      allBtn.className = 'mini-pill' + (!currentLogTeamFilter ? ' active' : '');
      allBtn.textContent = 'All Teams';
      allBtn.onclick = () => { currentLogTeamFilter = null; updateActivityLogUI(); };
      teamFilters.appendChild(allBtn);
      
      teams.forEach(t => {
          const btn = document.createElement('button');
          btn.className = 'mini-pill' + (currentLogTeamFilter === t ? ' active' : '');
          btn.textContent = t;
          btn.onclick = () => { 
              currentLogTeamFilter = (currentLogTeamFilter === t) ? null : t; 
              updateActivityLogUI(); 
          };
          teamFilters.appendChild(btn);
      });
  }

  const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
  
  if (searchInput && !searchInput.dataset.listenerAdded) {
    searchInput.addEventListener('input', updateActivityLogUI);
    searchInput.dataset.listenerAdded = 'true';
  }

  bundle.activityLog.forEach(entry => {
    const datePart = entry.date || '';
    const timePart = entry.time || '';
    const teamPart = entry.team || '';
    const membersPart = entry.members || '';
    const tagPart = entry.tag || 'base';
    const displayTag = tagPart.split(' - ')[0];
    const actionPart = entry.action || '';
    
    if (currentLogTeamFilter && teamPart !== currentLogTeamFilter) return;

    if (searchTerm) {
      const combined = `${datePart} ${timePart} ${teamPart} ${membersPart} ${tagPart} ${actionPart}`.toLowerCase();
      if (!combined.includes(searchTerm)) return;
    }

    const div = document.createElement('div');
    div.className = 'activity-log-entry';
    div.style.display = 'flex';
    div.style.alignItems = 'flex-start';
    div.style.gap = '10px';
    div.style.padding = '8px 0';
    
    const timePillContainer = createLogTimestampPill(entry);
    div.appendChild(timePillContainer);
    
    const infoSpan = document.createElement('span');
    if (teamPart === 'System') {
      infoSpan.textContent = ` ${actionPart}`;
    } else {
      infoSpan.textContent = ` Team ${teamPart} (${membersPart}) [${displayTag}]: ${actionPart}`;
    }
    div.appendChild(infoSpan);
    
    logBox.appendChild(div);
  });
}

function createLogTimestampPill(entry) {
  const container = document.createElement('div');
  container.className = 'pill-cell-container';
  container.style.justifyContent = 'flex-start';
  container.style.width = 'auto';

  const pill = document.createElement('div');
  pill.className = 'mini-pill clickable-pill';
  pill.style.fontSize = '0.75rem';
  pill.textContent = `${entry.date} ${entry.time}`;
  pill.onclick = (e) => {
    e.stopPropagation();
    showEditLogTimePopup(entry);
  };
  container.appendChild(pill);

  if (entry.tag && entry.tag.includes(' - ')) {
    const parts = entry.tag.split(' - ');
    const userName = parts[1];
    const userPill = document.createElement('div');
    userPill.className = 'mini-pill log-user-pill';
    userPill.style.fontSize = '0.7rem';
    userPill.style.background = 'rgba(255,255,255,0.05)';
    userPill.style.borderColor = 'rgba(255,255,255,0.1)';
    userPill.textContent = userName;
    container.appendChild(userPill);
  } else if (entry.team === 'System' && entry.tag === 'base') {
    const currentUser = getCurrentUser();
    if (currentUser) {
      const userPill = document.createElement('div');
      userPill.className = 'mini-pill log-user-pill';
      userPill.style.fontSize = '0.7rem';
      userPill.style.background = 'rgba(255,255,255,0.05)';
      userPill.style.borderColor = 'rgba(255,255,255,0.1)';
      userPill.textContent = currentUser.handle || (currentUser.firstName + ' ' + (currentUser.lastName || '')).trim();
      container.appendChild(userPill);
    }
  }

  if (entry.originalDate && entry.originalTime) {
    const resetBtn = document.createElement('button');
    resetBtn.className = 'mini-pill update-pill';
    resetBtn.style.padding = '0 5px';
    resetBtn.style.fontSize = '0.8rem';
    resetBtn.textContent = '↺';
    resetBtn.onclick = (e) => {
      e.stopPropagation();
      entry.date = entry.originalDate;
      entry.time = entry.originalTime;
      delete entry.originalDate;
      delete entry.originalTime;
      const bundle = loadBundle();
      const idx = bundle.activityLog.findIndex(l => l.timestamp === entry.timestamp);
      if (idx > -1) {
        bundle.activityLog[idx] = entry;
        saveBundle(bundle);
        refreshCurrentPageTable();
      }
    };
    container.appendChild(resetBtn);
  }

  return container;
}

function showEditLogTimePopup(entry) {
  const popup = createPopup('Edit Timestamp');
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');

  const container = document.createElement('div');
  container.style.display = 'flex';
  container.style.gap = '15px';
  container.style.marginBottom = '20px';

  const dateInput = document.createElement('input');
  dateInput.type = 'text';
  dateInput.className = 'form-input';
  dateInput.style.borderRadius = '8px';
  dateInput.style.textAlign = 'center';
  dateInput.placeholder = 'MM-DD-YYYY';
  dateInput.value = entry.date;
  dateInput.oninput = () => {
    let val = dateInput.value.replace(/\D/g, '');
    if (val.length > 2) val = val.substring(0, 2) + '-' + val.substring(2);
    if (val.length > 5) val = val.substring(0, 5) + '-' + val.substring(5, 9);
    dateInput.value = val;
  };

  const timeInput = document.createElement('input');
  timeInput.type = 'text';
  timeInput.className = 'form-input';
  timeInput.style.borderRadius = '8px';
  timeInput.style.textAlign = 'center';
  timeInput.placeholder = 'HH:MM';
  timeInput.value = entry.time;
  timeInput.oninput = () => {
    let val = timeInput.value.replace(/\D/g, '');
    if (val.length > 2) val = val.substring(0, 2) + ':' + val.substring(2, 4);
    timeInput.value = val;
  };

  container.appendChild(dateInput);
  container.appendChild(timeInput);
  content.insertBefore(container, btnContainer);

  const saveBtn = document.createElement('button');
  saveBtn.className = 'popup-btn primary';
  saveBtn.textContent = 'Save';
  saveBtn.onclick = () => {
    const dMatch = dateInput.value.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    const tMatch = timeInput.value.match(/^(\d{2}):(\d{2})$/);
    if (dMatch && tMatch) {
      if (!entry.originalDate) {
        entry.originalDate = entry.date;
        entry.originalTime = entry.time;
      }
      entry.date = dateInput.value;
      entry.time = timeInput.value;
      
      const bundle = loadBundle();
      const idx = bundle.activityLog.findIndex(l => l.timestamp === entry.timestamp);
      if (idx > -1) {
        bundle.activityLog[idx] = entry;
        saveBundle(bundle);
        refreshCurrentPageTable();
      }
      closePopup(popup);
    } else {
      alert('Invalid format. Use MM-DD-YYYY and HH:MM');
    }
  };
  btnContainer.appendChild(saveBtn);
}

function printCurrentReport(type) {
  const bundle = loadBundle();
  const log = bundle.activityLog || [];
  let title = '';
  let filteredLogs = [];

  if (type === 'team') {
    const filterContainer = document.getElementById('team-report-filters');
    const activeTeam = filterContainer ? filterContainer.dataset.activeTeam : null;
    if (!activeTeam) { alert('Please select a team first.'); return; }
    title = `Team Activity Report: ${activeTeam}`;
    filteredLogs = log.filter(e => e.team === activeTeam);
  } else {
    if (!currentMemberReportSelection) { alert('Please select a member first.'); return; }
    title = `Member Activity Report: ${currentMemberReportSelection}`;
    filteredLogs = log.filter(e => {
      if (!e.members) return false;
      const mList = e.members.split(', ').map(m => m.replace('*', '').trim());
      return mList.includes(currentMemberReportSelection);
    });
  }

  const printWindow = window.open('', '_blank');
  if (!printWindow) { alert("Please allow popups to view the printout."); return; }

  const style = `
    @page { size: auto; margin: 10mm; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: 11pt; color: #000; background: #fff; margin: 0; padding: 0; }
    .print-container { max-width: 8.5in; margin: 0 auto; padding: 20px; }
    h1 { font-size: 18pt; border-bottom: 2px solid #000; margin: 0 0 15px 0; padding-bottom: 5px; }
    .activity-log { font-size: 10pt; line-height: 1.0; }
    .activity-log-entry { margin-bottom: 1px; }
    .activity-log-time { font-weight: bold; margin-right: 5px; }
    .no-print { text-align: center; margin-bottom: 20px; padding: 10px; background: #f0f2f5; }
    @media print { .no-print { display: none; } }
  `;

  let htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      <style>${style}</style>
    </head>
    <body>
      <div class="no-print">
        <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 999px;">Print PDF</button>
      </div>
      <div class="print-container">
        <h1>${title}</h1>
        <div class="activity-log">
  `;

  filteredLogs.forEach(entry => {
    const tagPart = entry.tag || 'base';
    const displayTag = tagPart.split(' - ')[0];
    const userName = tagPart.includes(' - ') ? tagPart.split(' - ')[1] : '';
    
    htmlContent += '<div class="activity-log-entry">';
    htmlContent += `<span class="activity-log-time">[${entry.date} ${entry.time} ${displayTag}${userName ? ' (' + userName + ')' : ''}]</span> `;
    htmlContent += `Team ${entry.team} (${entry.members}): ${entry.action}`;
    htmlContent += '</div>';
  });

  htmlContent += `
        </div>
      </div>
      <script>
        setTimeout(() => { window.print(); }, 500);
      </script>
    </body>
    </html>
  `;

  printWindow.document.write(htmlContent);
  printWindow.document.close();
}

function printAllReports(type) {
    const bundle = loadBundle();
    const log = bundle.activityLog || [];
    const printWindow = window.open('', '_blank');
    if (!printWindow) { alert("Please allow popups to view the printout."); return; }

    const style = `
      @page { size: auto; margin: 10mm; }
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: 11pt; color: #000; background: #fff; margin: 0; padding: 0; }
      .print-container { max-width: 8.5in; margin: 0 auto; padding: 20px; }
      .print-section { page-break-after: always; padding: 20px 0; }
      .print-section:last-child { page-break-after: auto; }
      h1 { font-size: 18pt; border-bottom: 2px solid #000; margin: 0 0 15px 0; padding-bottom: 5px; }
      .activity-log { font-size: 10pt; line-height: 1.0; }
      .activity-log-entry { margin-bottom: 1px; }
      .activity-log-time { font-weight: bold; margin-right: 5px; }
      .no-print { text-align: center; margin-bottom: 20px; padding: 10px; background: #f0f2f5; }
      @media print { .no-print { display: none; } }
    `;

    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>All Reports</title>
        <style>${style}</style>
      </head>
      <body>
        <div class="no-print">
          <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 999px;">Print PDF</button>
        </div>
        <div class="print-container">
    `;
    
    if (type === 'team') {
        const teams = Array.from(new Set(log.map(e => e.team).filter(Boolean))).sort();
        teams.forEach(team => {
            htmlContent += '<div class="print-section">';
            htmlContent += `<h1>Team Activity Report: ${team}</h1>`;
            htmlContent += '<div class="activity-log">';
            log.filter(e => e.team === team).forEach(entry => {
                const tagPart = entry.tag || 'base';
                const displayTag = tagPart.split(' - ')[0];
                const userName = tagPart.includes(' - ') ? tagPart.split(' - ')[1] : '';
                htmlContent += `<div class="activity-log-entry"><span class="activity-log-time">[${entry.date} ${entry.time} ${displayTag}${userName ? ' (' + userName + ')' : ''}]</span> Team ${entry.team} (${entry.members}): ${entry.action}</div>`;
            });
            htmlContent += '</div></div>';
        });
    } else {
        const membersData = bundle.pages.page3 || [];
        const members = Array.from(new Set(membersData.map(m => m[0]).filter(Boolean))).sort();
        members.forEach(member => {
            const memberLogs = log.filter(e => {
                if (!e.members) return false;
                const mList = e.members.split(', ').map(m => m.replace('*', '').trim());
                return mList.includes(member);
            });
            if (memberLogs.length > 0) {
                htmlContent += '<div class="print-section">';
                htmlContent += `<h1>Member Activity Report: ${member}</h1>`;
                htmlContent += '<div class="activity-log">';
                memberLogs.forEach(entry => {
                    const tagPart = entry.tag || 'base';
                    const displayTag = tagPart.split(' - ')[0];
                    const userName = tagPart.includes(' - ') ? tagPart.split(' - ')[1] : '';
                    htmlContent += `<div class="activity-log-entry"><span class="activity-log-time">[${entry.date} ${entry.time} ${displayTag}${userName ? ' (' + userName + ')' : ''}]</span> Team ${entry.team} (${entry.members}): ${entry.action}</div>`;
                });
                htmlContent += '</div></div>';
            }
        });
    }

    htmlContent += `
        </div>
        <script>
          setTimeout(() => { window.print(); }, 500);
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(htmlContent);
    printWindow.document.close();
}

function buildTeamReports() {
  const filterContainer = document.getElementById('team-report-filters');
  const tableBody = document.getElementById('team-reports-body');
  if (!filterContainer || !tableBody) return;

  const bundle = loadBundle();
  const logs = bundle.activityLog || [];
  const teams = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot', 'Golf', 'Command', 'Off Duty', 'Base Support'];

  filterContainer.innerHTML = '';
  tableBody.innerHTML = '';

  teams.forEach(teamName => {
    const btn = document.createElement('button');
    btn.className = 'sub-nav-btn';
    btn.style.fontSize = '0.8rem';
    btn.style.padding = '5px 10px';
    btn.textContent = teamName;
    if (filterContainer.dataset.activeTeam === teamName) btn.classList.add('active');
    btn.onclick = () => {
      filterContainer.dataset.activeTeam = teamName;
      buildTeamReports();
    };
    filterContainer.appendChild(btn);
  });

  const activeTeam = filterContainer.dataset.activeTeam;
  if (!activeTeam) return;

  const filteredLogs = logs.filter(log => log.team === activeTeam);
  
  filteredLogs.forEach(entry => {
    const tr = document.createElement('tr');
    tr.className = 'report-row';
    
    const tdTime = document.createElement('td');
    tdTime.dataset.label = 'Timestamp';
    tdTime.appendChild(createLogTimestampPill(entry));
    tr.appendChild(tdTime);
    
    const tdTeam = document.createElement('td');
    tdTeam.dataset.label = 'Team';
    const teamPill = document.createElement('div');
    teamPill.className = 'mini-pill readonly-pill';
    teamPill.textContent = entry.team || '';
    tdTeam.appendChild(teamPill);
    tr.appendChild(tdTeam);
    
    const tdMembers = document.createElement('td');
    tdMembers.dataset.label = 'Members';
    const memberContainer = document.createElement('div');
    memberContainer.className = 'pill-container';
    memberContainer.style.justifyContent = 'flex-start';
    
    if (entry.members) {
      const members = entry.members.split(', ').map(m => m.trim());
      members.sort((a, b) => (b.endsWith('*') ? -1 : 1));
      members.forEach(m => {
        const pill = document.createElement('div');
        pill.className = 'mini-pill readonly-pill';
        if (m.endsWith('*')) {
          pill.style.background = 'var(--pill-focus)';
          pill.style.borderColor = 'var(--accent)';
          pill.style.fontWeight = '700';
        }
        pill.textContent = m;
        memberContainer.appendChild(pill);
      });
    }
    tdMembers.appendChild(memberContainer);
    tr.appendChild(tdMembers);
    
    const tagPart = entry.tag || 'base';
    const displayTag = tagPart.split(' - ')[0];
    const tdActivity = document.createElement('td');
    tdActivity.dataset.label = 'Activity';
    tdActivity.textContent = `[${displayTag}] ${entry.action}`;
    tr.appendChild(tdActivity);
    
    tableBody.appendChild(tr);
  });
}

let currentMemberReportSelection = '';

function buildMemberReports() {
  const btn = document.getElementById('member-report-btn');
  const tableBody = document.getElementById('member-reports-body');
  if (!btn || !tableBody) return;

  const bundle = loadBundle();
  const logs = bundle.activityLog || [];
  
  const onSceneMembers = (bundle.pages.page3 || []).filter(row => row[6] === 'true' && row[0]);

  btn.textContent = currentMemberReportSelection ? `Member: ${currentMemberReportSelection}` : 'Select Member...';
  btn.onclick = () => {
    const popup = createPopup('Select Member');
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');
    
    const list = document.createElement('div');
    list.style.display = 'flex';
    list.style.flexWrap = 'wrap';
    list.style.gap = '10px';
    list.style.marginBottom = '20px';
    
    onSceneMembers.forEach(mRow => {
      const mBtn = document.createElement('button');
      mBtn.className = 'mini-pill';
      mBtn.textContent = mRow[0];
      mBtn.onclick = () => {
        currentMemberReportSelection = mRow[0];
        closePopup(popup);
        buildMemberReports();
      };
      list.appendChild(mBtn);
    });

    content.insertBefore(list, btnContainer);
  };

  tableBody.innerHTML = '';
  if (!currentMemberReportSelection) return;

  const filteredLogs = logs.filter(log => {
    if (!log.members) return false;
    const members = log.members.split(', ').map(m => m.replace('*', '').trim());
    return members.includes(currentMemberReportSelection);
  });

  filteredLogs.forEach(entry => {
    const tr = document.createElement('tr');
    tr.className = 'report-row';
    
    const tdTime = document.createElement('td');
    tdTime.dataset.label = 'Timestamp';
    tdTime.appendChild(createLogTimestampPill(entry));
    tr.appendChild(tdTime);
    
    const tdTeam = document.createElement('td');
    tdTeam.dataset.label = 'Team';
    const teamPill = document.createElement('div');
    teamPill.className = 'mini-pill readonly-pill';
    teamPill.textContent = entry.team || '';
    tdTeam.appendChild(teamPill);
    tr.appendChild(tdTeam);
    
    const tdMembers = document.createElement('td');
    tdMembers.dataset.label = 'Members';
    const memberContainer = document.createElement('div');
    memberContainer.className = 'pill-container';
    memberContainer.style.justifyContent = 'flex-start';
    
    if (entry.members) {
      const members = entry.members.split(', ').map(m => m.trim());
      members.sort((a, b) => (b.endsWith('*') ? -1 : 1));
      members.forEach(m => {
        const pill = document.createElement('div');
        pill.className = 'mini-pill readonly-pill';
        if (m.endsWith('*')) {
          pill.style.background = 'var(--pill-focus)';
          pill.style.borderColor = 'var(--accent)';
          pill.style.fontWeight = '700';
        }
        pill.textContent = m;
        memberContainer.appendChild(pill);
      });
    }
    tdMembers.appendChild(memberContainer);
    tr.appendChild(tdMembers);
    
    const tagPart = entry.tag || 'base';
    const displayTag = tagPart.split(' - ')[0];
    const tdActivity = document.createElement('td');
    tdActivity.dataset.label = 'Activity';
    tdActivity.textContent = `[${displayTag}] ${entry.action}`;
    tr.appendChild(tdActivity);
    
    tableBody.appendChild(tr);
  });
}

function buildSearchLogTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  const clearBtn = document.getElementById('clear-table');
  const sortToggle = document.getElementById('sort-toggle');
  const sortLabel = document.getElementById('sort-label');

  recalculateEverything();
  const data = loadData();
  const bundle = loadBundle();

  // Identify unfinished tasks from team assignments and statuses
  const unfinishedTasks = new Set();
  if (bundle.currentAssignments && bundle.teamStatuses) {
    for (const team in bundle.currentAssignments) {
      const status = bundle.teamStatuses[team] || '';
      const assignment = bundle.currentAssignments[team] || '';
      // A task is unfinished if the team status is not "at base" and they have an assignment with a task number
      if (!status.includes('at base') && assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
        const match = assignment.match(/#(\d+)/);
        if (match) {
          unfinishedTasks.add(`#${match[1]}`);
        }
      }
    }
  }

  // Recalculate PSR After for all entries (now handled by recalculateEverything above)
  saveCurrentPageData(data);

  const isAscending = sortToggle && sortToggle.checked;
  if (sortLabel) {
    sortLabel.textContent = isAscending ? 'Sort Chronologically' : 'Sort Chronologically';
  }

  // To sort properly, we map row to an object with original index for stable editing if needed
  // but since we save back the entire array, sorting here might make editing complex if index changes.
  // HOWEVER, for a log, usually sorting is fine.
  
  const sortedData = [...data].sort((a, b) => {
    // Column 1: MM-DD-YYYY, Column 2: HH:mm
    const [m1, d1, y1] = (a[1] || '').split('-').map(Number);
    const [m2, d2, y2] = (b[1] || '').split('-').map(Number);
    const [h1, min1] = (a[2] || '').split(':').map(Number);
    const [h2, min2] = (b[2] || '').split(':').map(Number);

    const date1 = new Date(y1 || 0, (m1 || 1) - 1, d1 || 1, h1 || 0, min1 || 0);
    const date2 = new Date(y2 || 0, (m2 || 1) - 1, d2 || 1, h2 || 0, min2 || 0);

    return isAscending ? date1 - date2 : date2 - date1;
  });

  const params = new URLSearchParams(window.location.search);
  if (params.get('scroll') === 'latest') {
      const absoluteLatest = [...sortedData].sort((a, b) => {
           const [m1, d1, y1] = (a[1] || '').split('-').map(Number);
           const [m2, d2, y2] = (b[1] || '').split('-').map(Number);
           const [h1, min1] = (a[2] || '').split(':').map(Number);
           const [h2, min2] = (b[2] || '').split(':').map(Number);
           const date1 = new Date(y1 || 0, (m1 || 1) - 1, d1 || 1, h1 || 0, min1 || 0);
           const date2 = new Date(y2 || 0, (m2 || 1) - 1, d2 || 1, h2 || 0, min2 || 0);
           return date2 - date1;
      })[0];
      highlightedRowIndex = sortedData.indexOf(absoluteLatest);
  }

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headers = ['Task #', 'Date', 'Time', 'Region', 'Segment', 'PSR Before', 'PSR After', 'Team', 'Sweep Width (ft)', 'Num of Sweeps', 'Delete'];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.className = 'fixed-header';
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  for (let r = 0; r < sortedData.length; r++) {
    const tr = document.createElement('tr');
    animateNewRow(tr, r);
    
    // index 7 is team
    if (sortedData[r][7]) {
        // Handle team names like "Team 1 (3)" by extracting the name
        const fullTeamStr = sortedData[r][7];
        const teamNameMatch = fullTeamStr.match(/^(.*?)\s*\(/) || [null, fullTeamStr];
        const teamName = teamNameMatch[1].trim();
        animateArrivedRow(tr, teamName);
    }
    
    // Highlight if task is unfinished
    const taskNum = sortedData[r][0];
    if (unfinishedTasks.has(taskNum)) {
      tr.classList.add('unfinished-row');
    }

    // If this is the last entry by chronological order and we just arrived, we might scroll to it
    // But since it's sorted, latest might be top or bottom.
    
    const headers = ['Task #', 'Date', 'Time', 'Region', 'Segment', 'PSR Before', 'PSR After', 'Team', 'Sweep Width (ft)', 'Num of Sweeps', 'Delete'];
    for (let c = 0; c < 10; c++) {
      const td = document.createElement('td');
      td.dataset.label = headers[c];
      const cellContainer = document.createElement('div');
      cellContainer.className = 'pill-cell-container';

      const cell = document.createElement('div');
      cell.className = 'pill-cell';
      cell.dataset.row = String(r);
      cell.dataset.col = String(c);
      const cellValue = sortedData[r][c] || '';
      
      // index 7: Team, show as team pill if it looks like "TeamName (Count)"
      if (c === 7 && cellValue.includes('(') && cellValue.includes(')')) {
          const pill = document.createElement('div');
          pill.className = 'mini-pill readonly-pill';
          pill.textContent = cellValue;
          cell.appendChild(pill);
          cell.classList.add('readonly-pill');

          const teamName = cellValue.split(' (')[0];
          cell.style.cursor = 'pointer';
          cell.onclick = () => showTeamUpdatePopup(teamName);
          
          if (unfinishedTasks.has(taskNum)) {
            // Highlight orange if par check due
            if (isParCheckDue(teamName, bundle)) {
              cell.classList.add('par-check-due');
            }
            // Show progress bar
            const status = bundle.teamStatuses[teamName] || '';
            const progress = getTaskProgressPercent(status);
            cell.appendChild(createProgressBar(progress, taskNum));
          }
      } else {
          cell.textContent = cellValue;
      }
      
      cell.spellcheck = false;

      // Highlight Team (7) and Num of Sweeps (9) when blank
      if ([7, 9].includes(c) && cellValue === '') {
        cell.classList.add('blank-highlight');
      }

      // Read-only columns: Task # (0), Region (3), Segment (4), PSR Before (5), PSR After (6), and Team if it's a pill (7)
      const isReadonly = [0, 3, 4, 5, 6].includes(c) || (c === 7 && cellValue.includes('('));
      if (isReadonly) {
        cell.classList.add('readonly-pill');
      } else {
        cell.contentEditable = 'true';
      }

      cell.addEventListener('input', () => {
        if ([7, 9].includes(c)) {
          if (cell.textContent.trim() === '') {
            cell.classList.add('blank-highlight');
          } else {
            cell.classList.remove('blank-highlight');
          }
        }
        let val = cell.textContent.replace(/[^\d]/g, '');
        if (c === 1) { // Date MM-DD-YYYY
          if (val.length > 8) val = val.slice(0, 8);
          let formatted = val;
          if (val.length > 4) {
            formatted = val.slice(0, 2) + '-' + val.slice(2, 4) + '-' + val.slice(4);
          } else if (val.length > 2) {
            formatted = val.slice(0, 2) + '-' + val.slice(2);
          }
          if (cell.textContent !== formatted) {
            const selection = window.getSelection();
            const offset = selection.focusOffset;
            cell.textContent = formatted;
            // Simple caret restoration (approximate)
            try {
               placeCaretAtEnd(cell);
            } catch(e){}
          }
        } else if (c === 2) { // Time hh:mm
          if (val.length > 4) val = val.slice(0, 4);
          let formatted = val;
          if (val.length > 2) {
            formatted = val.slice(0, 2) + ':' + val.slice(2);
          }
          if (cell.textContent !== formatted) {
            cell.textContent = formatted;
            try {
              placeCaretAtEnd(cell);
            } catch(e){}
          }
        }
      });

      cell.addEventListener('blur', () => {
        // Need to find original row index if we are editing in sorted view
        const originalRow = sortedData[r];
        const oldVal = originalRow[c];
        const newVal = cell.textContent.trim();
        originalRow[c] = newVal;
        
        // If Num of Sweeps (9), Sweep Width (8), or Team (7) was changed, recalculate PSR After (6)
        if ([7, 8, 9].includes(c)) {
          const bundle = loadBundle();
          originalRow[6] = calculatePSRAfter(originalRow, bundle);
          
          if (c === 9 && oldVal !== newVal) {
              // If sweeps filled, we might need to update nav immediately
              recalculateEverything();
          }
        }
        
        saveCurrentPageData(data);
        buildSearchLogTable(); // Re-render to show updated PSR After and handle cascading if needed
      });

      cell.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          cell.blur();
          focusCell(Math.min(r + 1, sortedData.length - 1), c);
        }
        if (event.key === 'Tab') {
          event.preventDefault();
          cell.blur();
          const nextCol = event.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, 9);
          focusCell(r, nextCol);
        }
      });

      cellContainer.appendChild(cell);

      td.appendChild(cellContainer);
      tr.appendChild(td);
    }

    // New Delete Column
    const deleteTd = document.createElement('td');
    deleteTd.dataset.label = 'Delete';
    const deleteContainer = document.createElement('div');
    deleteContainer.className = 'pill-cell-container';
    const delBtn = document.createElement('button');
    delBtn.className = 'row-delete-btn';
    delBtn.textContent = 'Delete';
    delBtn.type = 'button';
    delBtn.onclick = () => {
      confirmDeleteRow(tr, () => {
        const rowToDelete = sortedData[r];
        const indexInData = data.indexOf(rowToDelete);
        if (indexInData > -1) {
            const taskNum = rowToDelete[0]; // e.g. "#1"
            let bundle = loadBundle();
            let changed = false;
            
            if (bundle.currentAssignments && taskNum) {
                for (const team in bundle.currentAssignments) {
                    const assignment = bundle.currentAssignments[team] || '';
                    if (assignment.startsWith(taskNum + ' ')) {
                        const now = new Date();
                        const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
                        bundle.teamStatuses[team] = `at base (${timeStr})`;
                        bundle.currentAssignments[team] = 'Base';
                        bundle.teamAssignmentTimes[team] = Date.now();
                        
                        // Delete log entries that referred to that team performing that task
                        bundle.activityLog = bundle.activityLog.filter(entry => 
                            !(entry.team === team && entry.tag === taskNum)
                        );
                        changed = true;
                    }
                }
            }

            const taskNumId = taskNum.startsWith('#') ? taskNum.substring(1) : taskNum;
            if (bundle.forms && bundle.forms[taskNumId]) {
                delete bundle.forms[taskNumId];
                changed = true;
            }
            
            if (changed) {
                saveBundle(bundle);
            }

            data.splice(indexInData, 1);
            logDeletion('Search Log Entry', taskNum);
            if (data.length === 0) data.push(Array.from({ length: 10 }, () => ''));
            saveCurrentPageData(data);
            buildSearchLogTable();
        }
      });
    };
    deleteContainer.appendChild(delBtn);
    deleteTd.appendChild(deleteContainer);
    tr.appendChild(deleteTd);

    tableBody.appendChild(tr);
    
    // Check if this is the "latest" search to scroll to it
    const params = new URLSearchParams(window.location.search);
    if (params.get('scroll') === 'latest') {
        // Identify latest by date/time
        // Actually, we can just scroll to the very last one added if it's the first time we load
        // But the requirement says "scroll to the latest search".
        // If sorting is DESCENDING (default), latest is at the top.
        // If sorting is ASCENDING, latest is at the bottom.
        
        // Let's scroll to the row that matches the absolute latest timestamp in sortedData
        const absoluteLatest = [...sortedData].sort((a, b) => {
             const [m1, d1, y1] = (a[1] || '').split('-').map(Number);
             const [m2, d2, y2] = (b[1] || '').split('-').map(Number);
             const [h1, min1] = (a[2] || '').split(':').map(Number);
             const [h2, min2] = (b[2] || '').split(':').map(Number);
             const date1 = new Date(y1 || 0, (m1 || 1) - 1, d1 || 1, h1 || 0, min1 || 0);
             const date2 = new Date(y2 || 0, (m2 || 1) - 1, d2 || 1, h2 || 0, min2 || 0);
             return date2 - date1;
        })[0];
        
        if (sortedData[r] === absoluteLatest) {
            setTimeout(() => {
                tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
                // Remove param so it doesn't keep scrolling on refresh
                const newUrl = window.location.pathname;
                window.history.replaceState({}, '', newUrl);
            }, 100);
        }
    }
  }

  const existing = document.querySelector('.add-row-container');
  if (existing) existing.remove();

  if (sortToggle) {
    sortToggle.onchange = () => {
        buildSearchLogTable();
    };
  }

  if (clearBtn) {
    clearBtn.remove();
  }
  
  initCharts();
}

let selectedChartDate = (new Date()).toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\//g, '-');
let selectedChartDuration = 24;

// Helper to format hour offset
function formatHourOffset(offset) {
    const h = Math.floor(offset);
    return `${h.toString().padStart(2, '0')}`;
}

// Helper to get timestamp
function getTs(dateStr, timeStr) {
    if (!dateStr) return 0;
    const [m, d, y] = dateStr.split('-').map(Number);
    const [hh, mm] = (timeStr || '00:00').split(':').map(Number);
    return new Date(y, m - 1, d, hh, mm).getTime();
}

function calculateHourlyMetrics(targetDateStr, durationHours = 24) {
    const bundle = loadBundle();
    const searchLog = bundle.pages.page4 || [];
    const segments = bundle.pages.page2 || [];
    const regions = bundle.pages.index.rows || [];
    
    const results = [];
    const increment = durationHours / 24;

    for (let i = 0; i < 24; i++) {
        const hourOffset = i * increment;
        const [m, d, y] = targetDateStr.split('-').map(Number);
        const baseDate = new Date(y, m - 1, d, 0, 0);
        const currentTs = baseDate.getTime() + (hourOffset * 3600000);
        
        let totalPOS = 0;
        let totalPSRc = 0;
        
        const hourMetrics = [];

        segments.forEach(segRow => {
            const regionName = segRow[0];
            const segmentName = segRow[1];
            const area = parseNumeric(segRow[2]);
            const length = parseNumeric(segRow[3]);
            const timePerSweep = parseNumeric(segRow[5]);
            const psri = parseNumeric(segRow[6]);

            // Region consensus and sum of areas
            const regionRow = regions.find(r => r[0] === regionName);
            const consensus = regionRow ? parseFloat(computeConsensus(bundle.pages.index, regions.indexOf(regionRow))) || 0 : 0;
            const sumOfAreas = segments.filter(s => s[0] === regionName).reduce((sum, s) => sum + parseNumeric(s[2]), 0);
            
            const initialPOC = (sumOfAreas > 0) ? (consensus * area / sumOfAreas) : 0;

            // Find all logs for this segment up to currentTs
            const logs = searchLog.filter(log => {
                if (log[3] !== regionName || log[4] !== segmentName) return false;
                return getTs(log[1], log[2]) <= currentTs;
            }).sort((a, b) => getTs(a[1], a[2]) - getTs(b[1], b[2]));

            let currentPOC = initialPOC;
            let currentPSR = psri;
            let segTotalPOS = 0;
            
            logs.forEach(log => {
                const sweepWidth = parseNumeric(log[8]);
                const numSweeps = parseNumeric(log[9]);
                const teamInfo = log[7] || '';
                const match = teamInfo.match(/\((\d+)\)/);
                const numMembers = match ? parseInt(match[1]) : 0;
                
                if (area > 0 && length > 0 && numMembers > 0) {
                    // Searcher Spacing (ft) = (Area (ac) / 640 / Length (mi) / NumSweeps / NumMembers) * 5280
                    const spacing = (area / 640 / length / numSweeps / numMembers) * 5280;
                    
                    // Coverage = Sweep Width (ft) / (Area (ac) / 640 / Length (mi) / NumMembers * 5280)
                    // Note: user's formula for coverage didn't include numSweeps in the denominator.
                    // "Coverage = sweep width / (area / 640 / length / num of team members searching on the team that searched this)"
                    // But if we have multiple sweeps, the spacing between searchers is different.
                    // Usually Coverage = SweepWidth / EffectiveSpacing.
                    // I'll stick to user's formula precisely:
                    const spacingForCoverage = (area / 640 / length / numMembers) * 5280;
                    const coverage = sweepWidth / spacingForCoverage;
                    
                    const pod = 1 - Math.exp(-coverage);
                    const pos = currentPOC * pod;
                    
                    segTotalPOS += pos;
                    currentPOC = currentPOC - pos;
                    
                    // PSR = length / time per sweep * Sweep Width * POCa / Area
                    // User says: length / time per sweep * Sweep Width (do not try to convert these ft to miles: leave these ft as ft) * POCa / Area
                    currentPSR = (length / timePerSweep) * sweepWidth * currentPOC / area;

                    // Store these for the detailed metrics table
                    hourMetrics.push({
                        region: regionName,
                        segment: segmentName,
                        spacing: spacing,
                        coverage: coverage,
                        pod: pod,
                        pos: pos,
                        psrc: currentPSR,
                        poca: currentPOC
                    });
                }
            });

            totalPOS += segTotalPOS;
            totalPSRc += (logs.length > 0) ? currentPSR : psri;
            
            // If no logs, still record current state for POCa/PSRc
            if (logs.length === 0) {
                hourMetrics.push({
                    region: regionName,
                    segment: segmentName,
                    spacing: 0,
                    coverage: 0,
                    pod: 0,
                    pos: 0,
                    psrc: psri,
                    poca: initialPOC
                });
            }
        });

        results.push({
            hour: hourOffset,
            totalPOS,
            totalPSRc,
            segments: hourMetrics
        });
    }
    
    return results;
}

function buildMetricsTable() {
    const tableBody = document.getElementById('metrics-table-body');
    const dateTitle = document.getElementById('metrics-date-title');
    if (!tableBody || !dateTitle) return;

    dateTitle.textContent = selectedChartDate;
    tableBody.innerHTML = '';
    
    const metricsData = calculateHourlyMetrics(selectedChartDate, selectedChartDuration);
    
    metricsData.forEach(m => {
        const tr = document.createElement('tr');
        
        // Hour Column
        const tdHour = document.createElement('td');
        tdHour.setAttribute('data-label', 'Hour');
        tdHour.textContent = formatHourOffset(m.hour);
        tdHour.style.fontWeight = 'bold';
        tr.appendChild(tdHour);
        
        // Segments logic
        const segs = m.segments || [];
        // Only count segments that actually have search stats
        const activeSegs = segs.filter(s => s.spacing > 0 || s.coverage > 0);
        const count = activeSegs.length || 1;
        
        const avgSpacing = activeSegs.reduce((s, seg) => s + seg.spacing, 0) / count;
        const avgCoverage = activeSegs.reduce((s, seg) => s + seg.coverage, 0) / count;
        const avgPOD = activeSegs.reduce((s, seg) => s + seg.pod, 0) / count;
        
        // PSR and POCa are cumulative/summed across the operation
        const avgPOCa = segs.reduce((s, seg) => s + seg.poca, 0); // Sum of POCa across all segments
        
        const createTd = (val, label) => {
            const td = document.createElement('td');
            td.setAttribute('data-label', label);
            td.textContent = isFinite(val) ? val.toFixed(4) : '0.0000';
            return td;
        };

        tr.appendChild(createTd(avgSpacing, 'Searcher Spacing (ft)'));
        tr.appendChild(createTd(avgCoverage, 'Coverage'));
        tr.appendChild(createTd(avgPOD, 'POD'));
        tr.appendChild(createTd(m.totalPOS, 'POS'));
        tr.appendChild(createTd(avgPOCa, 'POC After'));
        tr.appendChild(createTd(m.totalPSRc, 'PSR'));
        
        tableBody.appendChild(tr);
    });
}

function initCharts() {
    renderCharts();
    document.querySelectorAll('.chart-duration-btn').forEach(btn => {
        btn.onclick = () => {
            selectedChartDuration = parseInt(btn.dataset.duration);
            document.querySelectorAll('.chart-duration-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderCharts();
        };
    });
}

function renderCharts() {
    const targetDateStr = selectedChartDate; // MM-DD-YYYY
    const metrics = calculateHourlyMetrics(targetDateStr, selectedChartDuration);
    
    const psrcData = metrics.map(m => m.totalPSRc);
    const posData = metrics.map(m => m.totalPOS);

    drawLineChart('psrc-chart-container', psrcData, '#7dc6ff', selectedChartDuration);
    drawLineChart('activity-chart-container', posData, '#ff8c00', selectedChartDuration);
    
    const dateBtn = document.getElementById('chart-date-btn');
    if (dateBtn) {
        dateBtn.textContent = targetDateStr;
        dateBtn.onclick = showCalendarPopup;
    }
}

function drawLineChart(containerId, data, color, durationHours = 24) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const padding = { top: 20, right: 20, bottom: 30, left: 60 };
    const width = container.clientWidth || 300;
    const height = container.clientHeight || 150;
    const chartWidth = width - padding.left - padding.right;
    const chartHeight = height - padding.top - padding.bottom;
    const max = Math.max(...data, 1);

    container.innerHTML = '';
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", "100%");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.style.overflow = 'visible';

    // Y Axis (Vertical Axis)
    const yAxis = document.createElementNS("http://www.w3.org/2000/svg", "line");
    yAxis.setAttribute("x1", padding.left);
    yAxis.setAttribute("y1", padding.top);
    yAxis.setAttribute("x2", padding.left);
    yAxis.setAttribute("y2", height - padding.bottom);
    yAxis.setAttribute("stroke", "rgba(255,255,255,0.2)");
    yAxis.setAttribute("stroke-width", "1");
    svg.appendChild(yAxis);

    // Major increments on Y Axis
    const numTicks = 5;
    for (let i = 0; i <= numTicks; i++) {
        const val = (max / numTicks) * i;
        const y = height - padding.bottom - (val / max) * (height - padding.top - padding.bottom);
        
        const tick = document.createElementNS("http://www.w3.org/2000/svg", "line");
        tick.setAttribute("x1", padding.left - 5);
        tick.setAttribute("y1", y);
        tick.setAttribute("x2", padding.left);
        tick.setAttribute("y2", y);
        tick.setAttribute("stroke", "rgba(255,255,255,0.4)");
        svg.appendChild(tick);

        const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
        text.setAttribute("x", padding.left - 10);
        text.setAttribute("y", y + 4);
        text.setAttribute("fill", "var(--muted)");
        text.setAttribute("font-size", "10");
        text.setAttribute("text-anchor", "end");
        
        // If it's the activity-chart (POS), show as percent
        if (containerId === 'activity-chart-container') {
            text.textContent = (val * 100).toFixed(0) + '%';
        } else {
            text.textContent = val.toFixed(1);
        }
        svg.appendChild(text);
    }

    // X Axis Ticks
    const numXTicks = 8;
    for (let i = 0; i <= numXTicks; i++) {
        const offset = (durationHours / numXTicks) * i;
        const x = padding.left + (i / numXTicks) * chartWidth;
        
        const tick = document.createElementNS("http://www.w3.org/2000/svg", "line");
        tick.setAttribute("x1", x);
        tick.setAttribute("y1", height - padding.bottom);
        tick.setAttribute("x2", x);
        tick.setAttribute("y2", height - padding.bottom + 5);
        tick.setAttribute("stroke", "rgba(255,255,255,0.4)");
        svg.appendChild(tick);

        const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
        text.setAttribute("x", x);
        text.setAttribute("y", height - padding.bottom + 20);
        text.setAttribute("fill", "var(--muted)");
        text.setAttribute("font-size", "9");
        text.setAttribute("text-anchor", "middle");
        text.textContent = formatHourOffset(offset);
        svg.appendChild(text);
    }

    const points = data.map((val, i) => {
        const x = padding.left + (i / (data.length - 1)) * chartWidth;
        const y = height - padding.bottom - (val / max) * chartHeight;
        return { x, y };
    });

    if (points.length > 0) {
        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        let d = `M ${points[0].x} ${points[0].y}`;

        if (points.length > 1) {
            for (let i = 0; i < points.length - 1; i++) {
                const curr = points[i];
                const next = points[i + 1];
                const cp1x = curr.x + (next.x - curr.x) / 3;
                const cp2x = curr.x + 2 * (next.x - curr.x) / 3;
                d += ` C ${cp1x} ${curr.y}, ${cp2x} ${next.y}, ${next.x} ${next.y}`;
            }
        }

        path.setAttribute("d", d);
        path.setAttribute("fill", "none");
        path.setAttribute("stroke", color);
        path.setAttribute("stroke-width", "3");
        path.setAttribute("stroke-linecap", "round");
        path.setAttribute("stroke-linejoin", "round");
        svg.appendChild(path);
    }

    // No more dots/circles on line graphs as per request
    /*
    data.forEach((val, i) => {
        ...
    });
    */

    container.appendChild(svg);
}

function drawBarChart(containerId, data, color) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const max = Math.max(...data, 1);
    container.innerHTML = '';
    
    data.forEach((val, i) => {
        const bar = document.createElement('div');
        bar.style.flex = '1';
        bar.style.background = color;
        bar.style.height = `${(val / max) * 100}%`;
        bar.style.borderRadius = '4px 4px 0 0';
        bar.style.opacity = '0.6';
        bar.style.transition = 'height 0.3s ease';
        bar.title = `Hour ${i}:00 - Members: ${val}`;
        container.appendChild(bar);
    });
}

function showCalendarPopup() {
    const popup = createPopup('Select Report Date');
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');

    const [m, d, y] = selectedChartDate.split('-').map(Number);
    let viewMonth = m - 1;
    let viewYear = y;

    const calendarWrap = document.createElement('div');
    
    const renderMonth = () => {
        calendarWrap.innerHTML = '';
        const header = document.createElement('div');
        header.className = 'calendar-header';
        
        const prevBtn = document.createElement('button');
        prevBtn.className = 'mini-pill';
        prevBtn.textContent = '<';
        prevBtn.onclick = () => { viewMonth--; if (viewMonth < 0) { viewMonth = 11; viewYear--; } renderMonth(); };
        
        const title = document.createElement('div');
        title.style.fontWeight = '700';
        title.textContent = new Date(viewYear, viewMonth).toLocaleString('en-US', { month: 'long', year: 'numeric' });
        
        const nextBtn = document.createElement('button');
        nextBtn.className = 'mini-pill';
        nextBtn.textContent = '>';
        nextBtn.onclick = () => { viewMonth++; if (viewMonth > 11) { viewMonth = 0; viewYear++; } renderMonth(); };
        
        header.appendChild(prevBtn);
        header.appendChild(title);
        header.appendChild(nextBtn);
        calendarWrap.appendChild(header);
        
        const grid = document.createElement('div');
        grid.className = 'calendar-grid';
        
        ['S','M','T','W','T','F','S'].forEach(day => {
            const dDiv = document.createElement('div');
            dDiv.style.fontWeight = '700';
            dDiv.style.textAlign = 'center';
            dDiv.style.fontSize = '0.8rem';
            dDiv.textContent = day;
            grid.appendChild(dDiv);
        });
        
        const firstDay = new Date(viewYear, viewMonth, 1).getDay();
        const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
        
        for (let i = 0; i < firstDay; i++) {
            grid.appendChild(document.createElement('div'));
        }
        
        const today = new Date();
        const todayStr = `${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}-${today.getFullYear()}`;

        for (let i = 1; i <= daysInMonth; i++) {
            const dayDiv = document.createElement('div');
            dayDiv.className = 'calendar-day';
            dayDiv.textContent = i;
            
            const dateStr = `${(viewMonth + 1).toString().padStart(2, '0')}-${i.toString().padStart(2, '0')}-${viewYear}`;
            if (dateStr === todayStr) dayDiv.classList.add('today');
            if (dateStr === selectedChartDate) dayDiv.classList.add('selected');
            
            dayDiv.onclick = () => {
                calendarWrap.querySelectorAll('.calendar-day').forEach(d => d.classList.remove('selected'));
                dayDiv.classList.add('selected');
                tempSelectedDate = dateStr;
            };
            grid.appendChild(dayDiv);
        }
        calendarWrap.appendChild(grid);
    };

    let tempSelectedDate = selectedChartDate;
    renderMonth();
    content.insertBefore(calendarWrap, btnContainer);

    const actionRow = document.createElement('div');
    actionRow.style.display = 'flex';
    actionRow.style.gap = '10px';
    actionRow.style.marginTop = '20px';

    const applyBtn = document.createElement('button');
    applyBtn.className = 'popup-btn primary';
    applyBtn.style.flex = '1';
    applyBtn.textContent = 'Apply';
    applyBtn.onclick = () => {
        selectedChartDate = tempSelectedDate;
        closePopup(popup);
        renderCharts();
    };
    
    actionRow.appendChild(applyBtn);
    btnContainer.appendChild(actionRow);
}

function buildSavedFilesTable() {
    const tbody = document.getElementById('saved-files-body');
    if (!tbody) return;

    const files = getSavedFiles();
    const fileNames = Object.keys(files).sort();
    
    if (fileNames.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align: center; color: var(--muted); padding: 20px;">No saved search files yet.</td></tr>';
        return;
    }

    tbody.innerHTML = '';
    const currentUser = getCurrentUser();
    const isAdmin = isUserAdmin(currentUser);
    const isFileManager = currentUser && (currentUser.isFileManager === true || currentUser.isFileManager === 'true');

    fileNames.forEach(name => {
        const fileInfo = files[name];
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid rgba(255,255,255,0.05)';

        // Name (Open button-like pill)
        const tdName = document.createElement('td');
        tdName.setAttribute('data-label', 'File Name');
        tdName.style.padding = '12px 15px';
        const nameBtn = document.createElement('button');
        nameBtn.className = 'mini-pill';
        nameBtn.style.fontWeight = 'bold';
        nameBtn.textContent = name;
        nameBtn.onclick = () => {
            saveBundle(fileInfo.bundle);
            window.location.reload();
        };
        tdName.appendChild(nameBtn);
        tr.appendChild(tdName);

        // Date
        const tdDate = document.createElement('td');
        tdDate.setAttribute('data-label', 'Last Modified');
        tdDate.style.padding = '12px 15px';
        tdDate.style.color = 'var(--muted)';
        tdDate.textContent = fileInfo.lastModified;
        tr.appendChild(tdDate);

        // Size
        const tdSize = document.createElement('td');
        tdSize.setAttribute('data-label', 'File Size');
        tdSize.style.padding = '12px 15px';
        tdSize.style.color = 'var(--muted)';
        const sizeInBytes = JSON.stringify(fileInfo.bundle).length;
        const sizeInKB = (sizeInBytes / 1024).toFixed(1);
        tdSize.textContent = `${sizeInKB} KB`;
        tr.appendChild(tdSize);

        // Actions
        const tdActions = document.createElement('td');
        tdActions.setAttribute('data-label', 'Actions');
        tdActions.style.padding = '12px 15px';
        tdActions.style.textAlign = 'center';
        
        const btnCont = document.createElement('div');
        btnCont.className = 'tool-actions';
        btnCont.style.justifyContent = 'center';
        btnCont.style.gap = '10px';

        const downBtn = document.createElement('button');
        downBtn.className = 'mini-pill';
        downBtn.textContent = 'Download';
        downBtn.onclick = () => {
            downloadTextFile(name, JSON.stringify(fileInfo.bundle, null, 2));
        };
        btnCont.appendChild(downBtn);

        const delBtn = document.createElement('button');
        delBtn.className = 'row-delete-btn';
        delBtn.textContent = 'Delete';
        delBtn.onclick = () => {
            if (isAdmin || isFileManager) {
                const b = loadBundle();
                const doDelete = () => {
                    deleteFileFromList(name);
                    buildSavedFilesTable();
                };
                if (b.deleteMode) {
                    doDelete();
                } else if (confirm(`Are you sure you want to delete "${name}"?`)) {
                    doDelete();
                }
            } else {
                alert('You do not have permission to delete files. Contact Super Admin or a File Manager.');
            }
        };
        btnCont.appendChild(delBtn);

        tdActions.appendChild(btnCont);
        tr.appendChild(tdActions);

        tbody.appendChild(tr);
    });
}

function buildHomePage() {
  updateFileNameDisplay();
  const fileNameInput = document.getElementById('bundle-file-name');
  const saveNameBtn = document.getElementById('save-file-name');
  const homeStatus = document.getElementById('home-status');

  const bundle = loadBundle();
  fileNameInput.value = bundle.fileName;

  buildSavedFilesTable();

  const createNewBtn = document.getElementById('create-new-search-btn');
  if (createNewBtn) {
    createNewBtn.onclick = () => {
        const popup = createPopup('Create New Search?', createNewBtn);
        const content = popup.querySelector('.popup-content');
        const btnContainer = popup.querySelector('.popup-buttons');
        
        const desc = document.createElement('p');
        desc.style.color = 'var(--muted)';
        desc.style.fontSize = '0.9rem';
        desc.style.margin = '10px 0 20px 0';
        desc.textContent = 'Create a new search file? Registered personnel will be preserved but set to off-scene.';
        content.insertBefore(desc, btnContainer);
        
        const confirmBtn = document.createElement('button');
        confirmBtn.className = 'popup-btn primary';
        confirmBtn.textContent = 'Confirm';
        confirmBtn.onclick = () => {
            const currentBundle = loadBundle();
            const newBundle = defaultBundle();
            
            // Preserve personnel but set to off-scene
            const oldPersonnel = currentBundle.pages.page3 || [];
            const preservedPersonnel = oldPersonnel.filter(r => r[0] && r[0].trim() !== '').map(r => {
                const newRow = [...r];
                newRow[1] = ''; // Clear team
                newRow[2] = ''; // Clear lead
                newRow[6] = 'false'; // Off-scene
                return newRow;
            });
            
            if (preservedPersonnel.length > 0) {
                newBundle.pages.page3 = preservedPersonnel;
            }
            
            // Preserve accounts
            newBundle.accounts = currentBundle.accounts;

            logCreation('New Search File', newBundle.fileName, newBundle);
            saveBundle(newBundle);
            window.location.reload();
        };
        
        const cancelBtn = document.createElement('button');
        cancelBtn.className = 'popup-btn';
        cancelBtn.textContent = 'Cancel';
        cancelBtn.onclick = () => closePopup(popup);
        
        btnContainer.appendChild(confirmBtn);
        btnContainer.appendChild(cancelBtn);
    };
  }

  const printBtn = document.getElementById('print-search-file-btn');
  if (printBtn) {
    printBtn.onclick = () => printSearchFile();
  }

  // Populate stats
  const statRegions = document.getElementById('stat-regions');
  const statSegments = document.getElementById('stat-segments');
  const statPersonnel = document.getElementById('stat-personnel');
  const statTasks = document.getElementById('stat-tasks');

  if (statRegions) {
    const rows = bundle.pages.index.rows || [];
    statRegions.textContent = rows.filter(r => r[0] && r[0].trim() !== '').length;
  }
  if (statSegments) {
    const rows = bundle.pages.page2 || [];
    statSegments.textContent = rows.filter(r => r[1] && r[1].trim() !== '').length;
  }
  if (statPersonnel) {
    const rows = bundle.pages.page3 || [];
    statPersonnel.textContent = rows.filter(r => r[0] && r[0].trim() !== '').length;
  }
  if (statTasks) {
    const rows = bundle.pages.page4 || [];
    statTasks.textContent = rows.filter(r => r[0] && r[0].trim() !== '').length;
  }

  const recentLogsContainer = document.getElementById('recent-logs');
  const updateRecentLogs = () => {
    if (!recentLogsContainer) return;
    const b = loadBundle();
    const logs = b.activityLog || [];
    if (logs.length > 0) {
      recentLogsContainer.innerHTML = '';
      // Take the latest 3 logs (they are at the beginning because of unshift)
      const latestThree = logs.slice(0, 3);
      latestThree.forEach(entry => {
        const div = document.createElement('div');
        div.className = 'activity-log-entry';
        div.style.marginBottom = '5px';
        
        const tagPart = entry.tag || 'base';
        const displayTag = tagPart.split(' - ')[0];
        const userName = tagPart.includes(' - ') ? tagPart.split(' - ')[1] : '';
        const timeSpan = document.createElement('span');
        timeSpan.className = 'activity-log-time';
        timeSpan.textContent = '[' + entry.time + ' ' + displayTag + (userName ? ' ' + userName : '') + ']';
        
        const text = document.createTextNode(` Team ${entry.team}: ${entry.action}`);
        
        div.appendChild(timeSpan);
        div.appendChild(text);
        recentLogsContainer.appendChild(div);
      });
    } else {
      recentLogsContainer.innerHTML = '<p>No recent activity logs.</p>';
    }
  };

  if (recentLogsContainer) {
    updateRecentLogs();
    // Refresh every 5 seconds to catch new logs from other pages/background
    setInterval(updateRecentLogs, 5000);
  }

  const importBtn = document.getElementById('import-search-btn');
  const importInput = document.getElementById('import-search-input');
  if (importBtn && importInput) {
    importBtn.onclick = () => importInput.click();
    importInput.onchange = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const importedBundle = JSON.parse(event.target.result);
          if (!importedBundle.pages || !importedBundle.fileName) {
            throw new Error('Invalid search file format. Missing pages or fileName.');
          }
          logCreation('Imported Search File', importedBundle.fileName, importedBundle);
          saveBundle(importedBundle);
          saveFileToList(importedBundle.fileName, importedBundle);
          window.location.reload();
        } catch (err) {
          alert('Error importing file: ' + err.message);
        }
      };
      reader.readAsText(file);
    };
  }

  saveNameBtn.onclick = () => {
    const currentBundle = loadBundle();
    let nextName = fileNameInput.value.trim() || DEFAULT_FILE_NAME;
    if (!nextName.toLowerCase().endsWith('.json')) nextName += '.json';
    
    // Check if we are renaming an existing file list entry
    const files = getSavedFiles();
    const oldName = currentBundle.fileName;
    
    currentBundle.fileName = nextName;
    saveBundle(currentBundle);
    
    // If not already in the list, add it
    if (!files[nextName] && !files[oldName]) {
        saveFileToList(nextName, currentBundle);
    }
    
    fileNameInput.value = nextName;
    homeStatus.textContent = `File identifier updated to ${nextName} and saved to list.`;
    updateFileNameDisplay();
    buildSavedFilesTable();
  };

  initCharts();

  const viewMetricsBtn = document.getElementById('view-metrics-btn');
  const backToDashboardBtn = document.getElementById('back-to-dashboard-btn');
  const dashboardView = document.getElementById('home-dashboard-view');
  const metricsView = document.getElementById('home-metrics-view');

  if (viewMetricsBtn && backToDashboardBtn && dashboardView && metricsView) {
      viewMetricsBtn.onclick = () => {
          dashboardView.style.display = 'none';
          metricsView.style.display = 'block';
          buildMetricsTable();
      };
      backToDashboardBtn.onclick = () => {
          dashboardView.style.display = 'grid';
          metricsView.style.display = 'none';
      };
  }
}

function applyTheme(bundle) {
  const user = getCurrentUser();
  let theme = bundle.theme || 'dark';
  
  if (user) {
    // Look up current user in bundle to get latest theme preference
    const actualUser = (bundle.accounts || []).find(a => 
      a.firstName === user.firstName && a.lastName === user.lastName && a.pin === user.pin
    );
    if (actualUser && actualUser.theme) {
      theme = actualUser.theme;
    }
  }

  if (theme === 'light') {
    document.body.classList.add('light-mode');
  } else {
    document.body.classList.remove('light-mode');
  }
}

function applyBackground(bundle) {
  if (bundle && bundle.background) {
    document.body.style.backgroundImage = `linear-gradient(var(--bg-dim-start), var(--bg-dim-end)), url('${bundle.background}')`;
  }
}

function applyTipsVisibility(bundle) {
  const showTips = bundle.showTips !== false; // Default to true
  const tips = document.querySelectorAll('.hero p');
  tips.forEach(p => {
    // Only target p elements that are direct children of hero or within hero-col-left
    // to avoid hiding actual content if any hero section uses p for something else.
    // Based on exploration, these are indeed the informational tips.
    if (p.closest('.hero')) {
      p.style.display = showTips ? '' : 'none';
    }
  });
}

function buildSettingsPage() {
  const toggle = document.getElementById('delete-mode-toggle');
  const label = document.getElementById('delete-mode-label');
  const themeToggle = document.getElementById('theme-toggle');
  const themeLabel = document.getElementById('theme-label');
  const tipsToggle = document.getElementById('tips-toggle');
  const tipsLabel = document.getElementById('tips-label');
  const status = document.getElementById('settings-status');
  const bgInput = document.getElementById('bg-image-input');
  const resetBgBtn = document.getElementById('reset-bg-btn');
  const parFreqInput = document.getElementById('par-freq-input');
  const bundle = loadBundle();

  toggle.checked = !!bundle.deleteMode;
  label.textContent = `Delete Mode is ${toggle.checked ? 'ON' : 'OFF'}`;

  if (themeToggle) {
    themeToggle.checked = bundle.theme === 'light';
    themeLabel.textContent = bundle.theme === 'light' ? 'Grey Mode' : 'Dark Mode';
    themeToggle.onchange = () => {
      const nextBundle = loadBundle();
      nextBundle.theme = themeToggle.checked ? 'light' : 'dark';
      saveBundle(nextBundle);
      applyTheme(nextBundle);
      applyBackground(nextBundle);
      themeLabel.textContent = nextBundle.theme === 'light' ? 'Grey Mode' : 'Dark Mode';
      status.textContent = 'Theme updated and saved.';
    };
  }

  if (tipsToggle) {
    tipsToggle.checked = bundle.showTips !== false;
    tipsLabel.textContent = `Tips are ${tipsToggle.checked ? 'ON' : 'OFF'}`;
    tipsToggle.onchange = () => {
      const nextBundle = loadBundle();
      nextBundle.showTips = tipsToggle.checked;
      saveBundle(nextBundle);
      applyTipsVisibility(nextBundle);
      tipsLabel.textContent = `Tips are ${tipsToggle.checked ? 'ON' : 'OFF'}`;
      status.textContent = 'Tips display preference updated.';
    };
  }

  if (parFreqInput) {
    parFreqInput.value = bundle.parCheckFrequency || 20;
    parFreqInput.onchange = () => {
      const val = parseInt(parFreqInput.value);
      if (isNaN(val) || val < 1) {
        parFreqInput.value = 20;
        return;
      }
      const nextBundle = loadBundle();
      nextBundle.parCheckFrequency = val;
      saveBundle(nextBundle);
      status.textContent = `Par check frequency updated to ${val} minutes.`;
    };
  }

  toggle.onchange = () => {
    const nextBundle = loadBundle();
    nextBundle.deleteMode = toggle.checked;
    saveBundle(nextBundle);
    label.textContent = `Delete Mode is ${toggle.checked ? 'ON' : 'OFF'}`;
    status.textContent = `Settings saved automatically.`;
  };

  if (bgInput) {
    bgInput.onchange = async () => {
      const file = bgInput.files?.[0];
      if (!file) return;

      try {
        const reader = new FileReader();
        reader.onload = (e) => {
          const nextBundle = loadBundle();
          nextBundle.background = e.target.result;
          saveBundle(nextBundle);
          applyBackground(nextBundle);
          status.textContent = 'Background updated and saved.';
        };
        reader.readAsDataURL(file);
      } catch (err) {
        status.textContent = 'Could not load the image.';
      }
      bgInput.value = '';
    };
  }

  if (resetBgBtn) {
    resetBgBtn.onclick = () => {
      const nextBundle = loadBundle();
      nextBundle.background = 'assets/us-night.jpg';
      saveBundle(nextBundle);
      applyBackground(nextBundle);
      status.textContent = 'Background reset to default.';
    };
  }

  const syncUrlInput = document.getElementById('sync-url-input');
  const saveSyncBtn = document.getElementById('save-sync-url-btn');
  const syncStatusMsg = document.getElementById('sync-status-msg');
  const proxyInput = document.getElementById('sartopo-proxy-input');
  const saveProxyBtn = document.getElementById('save-proxy-btn');

  if (proxyInput && saveProxyBtn) {
    proxyInput.value = getSartopoProxy();
    saveProxyBtn.onclick = () => {
      setSartopoProxy(proxyInput.value.trim());
      status.textContent = 'SarTopo Proxy URL saved.';
    };
  }

  if (syncUrlInput && saveSyncBtn) {
    const syncBucketInput = document.getElementById('sync-bucket-input');
    if (syncBucketInput) syncBucketInput.value = getSyncBucket();

    saveSyncBtn.onclick = async () => {
      if (syncBucketInput) {
        const bucket = syncBucketInput.value.trim();
        if (bucket) {
           localStorage.setItem(SYNC_BUCKET_STORAGE_KEY, bucket);
           syncStatusMsg.textContent = 'Sync bucket saved! Testing connection...';
           
           try {
             // Test connection by trying to get the bundle
             const bucketUrl = `${KVDB_BASE_URL}/${bucket}`;
             const resp = await fetch(`${bucketUrl}/bundle`);
             if (resp.ok) {
                syncStatusMsg.textContent = 'Sync connection successful! Data found in bucket.';
             } else if (resp.status === 404) {
                syncStatusMsg.textContent = 'Connected! This is a new bucket, data will be pushed on next change.';
             } else {
                syncStatusMsg.textContent = `Server returned status ${resp.status}. Bucket might be invalid.`;
             }
             syncWithServer();
           } catch (err) {
             syncStatusMsg.textContent = 'Could not reach sync service (kvdb.io). Check your internet connection.';
           }
        }
      }
    };
  }
}

function buildProfilePage() {
  const container = document.getElementById('profile-form-container');
  if (!container) return;

  const bundle = loadBundle();
  const profile = bundle.profile || {};

  // Redesign Profile Page Header for two columns
  const hero = document.querySelector('.hero');
  if (hero) {
    hero.innerHTML = `
      <div class="hero-columns">
        <div class="hero-col-left">
          <h1>Profile</h1>
          <p>Enter the general incident and lost person information below.</p>
        </div>
        <div class="hero-col-right" style="display: flex; align-items: flex-end; justify-content: flex-end;">
          <div style="display: flex; align-items: center; gap: 10px; background: rgba(0,0,0,0.2); padding: 10px 20px; border-radius: 999px; border: 1px solid var(--line);">
            <input type="checkbox" id="profile-completed" class="pill-checkbox" ${profile.completed ? 'checked' : ''}>
            <label for="profile-completed" style="font-weight: bold; cursor: pointer;">Profile Completed</label>
          </div>
        </div>
      </div>
    `;
    const cb = document.getElementById('profile-completed');
    if (cb) {
      cb.onchange = () => {
        const b = loadBundle();
        if (!b.profile) b.profile = {};
        b.profile.completed = cb.checked;
        saveBundle(b);
        checkParChecksAndNotify(); // Trigger header refresh
      };
    }
  }

  container.innerHTML = '';
  
  const form = document.createElement('div');
  form.className = 'task-form'; 
  
  const save = () => {
    const b = loadBundle();
    b.profile = profile;
    saveBundle(b);
  };

  const addGroup = (label, key, type = 'text') => {
    const grp = document.createElement('div');
    grp.className = 'form-group ' + (type === 'textarea' ? 'large' : 'small');
    grp.innerHTML = `<label>${label}</label>`;
    
    let input;
    if (type === 'textarea') {
      input = document.createElement('textarea');
      input.className = 'form-input';
      input.style.minHeight = '100px';
    } else {
      input = document.createElement('input');
      input.type = type;
      input.className = 'form-input';
    }
    
    input.value = profile[key] || '';
    input.oninput = () => {
      profile[key] = input.value;
      save();
    };
    
    grp.appendChild(input);
    form.appendChild(grp);
  };

  addGroup('Incident #', 'incidentNumber');
  addGroup('Lost Person Name', 'lostPersonName');
  addGroup('Age', 'lostPersonAge');
  addGroup('Gender', 'lostPersonGender');
  addGroup('Description', 'lostPersonDescription', 'textarea');
  addGroup('Clothing', 'lostPersonClothing', 'textarea');
  addGroup('Physical / Medical', 'lostPersonPhysical', 'textarea');

  container.appendChild(form);
}

let currentFormsSubpage = 'task-assignment';
let currentTaskNumber = null;

function buildFormsPage() {
  const urlParams = new URLSearchParams(window.location.search);
  const taskParam = urlParams.get('task');
  if (taskParam) {
     currentTaskNumber = taskParam;
     currentFormsSubpage = 'task-assignment';
  }

  const btnTask = document.getElementById('btn-task-assignment');
  const btnManage = document.getElementById('btn-manage-forms');
  const taskView = document.getElementById('task-assignment-view');
  const manageView = document.getElementById('manage-forms-view');
  const printContainer = document.getElementById('print-btn-container');

  if (btnTask) {
    btnTask.onclick = () => {
      currentFormsSubpage = 'task-assignment';
      buildFormsPage();
    };
  }
  if (btnManage) {
    btnManage.onclick = () => {
      currentFormsSubpage = 'manage-forms';
      buildFormsPage();
    };
  }

  // Set visibility and active classes based on state
  if (currentFormsSubpage === 'task-assignment') {
    if (btnTask) btnTask.classList.add('active');
    if (btnManage) btnManage.classList.remove('active');
    if (taskView) taskView.style.display = 'block';
    if (manageView) manageView.style.display = 'none';
  } else {
    if (btnManage) btnManage.classList.add('active');
    if (btnTask) btnTask.classList.remove('active');
    if (taskView) taskView.style.display = 'none';
    if (manageView) manageView.style.display = 'block';
  }

  if (printContainer) printContainer.innerHTML = '';

  if (printContainer) {
    const downloadAllBtn = document.createElement('button');
    downloadAllBtn.id = 'download-all-forms-btn';
    downloadAllBtn.className = 'sub-nav-btn';
    downloadAllBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 8px; vertical-align: middle;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>Download All Forms';
    downloadAllBtn.onclick = () => downloadAllForms();
    printContainer.appendChild(downloadAllBtn);
  }

  if (currentFormsSubpage === 'task-assignment') {
    if (printContainer) {
      const printBtn = document.createElement('button');
      printBtn.className = 'sub-nav-btn active';
      printBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 8px; vertical-align: middle;"><polyline points="6 9 6 2 18 2 18 9"></polyline><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>Print Form';
        printBtn.onclick = () => {
            if (currentTaskNumber) {
                printSingleTaskForm(currentTaskNumber);
            } else {
                alert("Please select a task first.");
            }
        };
      printContainer.appendChild(printBtn);
    }
    buildTaskAssignmentForm();
  } else {
    buildManageFormsTable();
  }
}

function buildTaskAssignmentForm() {
  const pillsContainer = document.getElementById('task-pills-container');
  const container = document.getElementById('interactive-form-container');
  if (!pillsContainer || !container) return;

  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  
  // Identify unfinished tasks
  const unfinishedTasks = new Set();
  if (bundle.currentAssignments && bundle.teamStatuses) {
    for (const team in bundle.currentAssignments) {
      const status = bundle.teamStatuses[team] || '';
      const assignment = bundle.currentAssignments[team] || '';
      if (!status.includes('at base') && assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
        const match = assignment.match(/#(\d+)/);
        if (match) unfinishedTasks.add(match[1]);
      }
    }
  }

  const tasks = [];
  searchLog.forEach(row => {
    if (row[0] && row[0].startsWith('#')) {
      const num = row[0].substring(1);
      const region = row[3] || '';
      const segment = row[4] || '';
      const teamWithCount = row[7] || '';
      let teamName = teamWithCount;
      if (teamWithCount.includes(' (')) {
        teamName = teamWithCount.split(' (')[0];
      }

      if (!tasks.some(t => t.num === num)) {
        tasks.push({ num, region, segment, teamName });
      }
    }
  });
  tasks.sort((a, b) => parseInt(a.num) - parseInt(b.num));

  pillsContainer.innerHTML = '';
  tasks.forEach(task => {
    const btn = document.createElement('button');
    btn.className = 'mini-pill';
    btn.textContent = `#${task.num} ${task.region} ${task.segment} ${task.teamName}`;
    
    const isUnfinished = unfinishedTasks.has(task.num);
    const form = bundle.forms?.[task.num];
    const isComplete = form && form.completed;

    if (isUnfinished) {
       btn.style.opacity = '0.5';
       btn.style.background = 'rgba(128, 128, 128, 0.2)';
    } else if (!isComplete) {
       btn.style.background = 'rgba(255, 140, 0, 0.25)';
       btn.style.borderColor = 'rgba(255, 140, 0, 0.5)';
       btn.style.color = '#ff8c00';
    } else {
       // Standard website format
    }

    if (task.num === currentTaskNumber) {
      btn.style.borderWidth = '2px';
      btn.style.boxShadow = '0 0 8px var(--accent)';
    }

    btn.onclick = () => {
      currentTaskNumber = task.num;
      buildTaskAssignmentForm();
    };
    pillsContainer.appendChild(btn);
  });

  if (!currentTaskNumber) {
    container.innerHTML = '<div style="text-align: center; color: var(--text); opacity: 0.6; margin-top: 100px;">Select a Task # above to manage its form.</div>';
    return;
  }

  if (!bundle.forms) bundle.forms = {};
  if (!bundle.forms[currentTaskNumber]) {
    const searchRow = searchLog.find(row => row[0] === '#' + currentTaskNumber);
    const dateStamp = searchRow ? searchRow[1] : '';
    const teamNameWithCount = searchRow ? searchRow[7] : '';
    
    let teamName = '';
    if (teamNameWithCount) {
       const match = teamNameWithCount.match(/^(.*)\s\(\d+\)$/);
       teamName = match ? match[1] : teamNameWithCount;
    }

    const members = getTeamMembers(teamName).map(m => ({
      name: m[0],
      leader: m[2] === m[0],
      gps: false,
      radio: false,
      medic: false
    }));
    
    bundle.forms[currentTaskNumber] = {
      incidentNumber: '',
      opPeriod: '1',
      dateTime: dateStamp,
      lostPersonName: '',
      lostPersonAge: '',
      lostPersonGender: '',
      lostPersonDescription: '',
      lostPersonClothing: '',
      lostPersonPhysical: '',
      onSceneFamily: false,
      onSceneMedia: false,
      briefedBy: '',
      radioNumber: '',
      gpsNumber: '',
      leaveBase: '',
      beginSearch: '',
      completeSearch: '',
      returnBase: '',
      teamType: '', // Now using checklist
      teamTypes: { hasty: false, grid: false, area: false, k9: false, atv: false, argo: false, drone: false, boat: false, other: false },
      otherTeamType: '',
      instructions: '',
      teamName: teamName,
      teamMembers: members,
      statusUpdates: Array.from({length: 8}, () => ({time: '', clue: '', usng: ''}))
    };
    saveBundle(bundle);
    addActivityLogEntry(teamName, 'Task Form auto-created for #' + currentTaskNumber);
  }

  renderTaskForm(container, currentTaskNumber, bundle.forms[currentTaskNumber]);
}

function renderTaskForm(container, taskNum, formData) {
  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  const profile = bundle.profile || {};
  const taskTag = '#' + taskNum;

  container.innerHTML = '';
  
  const form = document.createElement('div');
  form.className = 'task-form';
  
  const save = () => {
    const b = loadBundle();
    b.forms[taskNum] = formData;
    saveBundle(b);
  };

  // 1. Auto-fill from Profile if blank
  let changed = false;
  if (!formData.incidentNumber && profile.incidentNumber) { formData.incidentNumber = profile.incidentNumber; changed = true; }
  if (!formData.lostPersonName && profile.lostPersonName) { formData.lostPersonName = profile.lostPersonName; changed = true; }
  if (!formData.lostPersonAge && profile.lostPersonAge) { formData.lostPersonAge = profile.lostPersonAge; changed = true; }
  if (!formData.lostPersonGender && profile.lostPersonGender) { formData.lostPersonGender = profile.lostPersonGender; changed = true; }
  if (!formData.lostPersonDescription && profile.lostPersonDescription) { formData.lostPersonDescription = profile.lostPersonDescription; changed = true; }
  if (!formData.lostPersonClothing && profile.lostPersonClothing) { formData.lostPersonClothing = profile.lostPersonClothing; changed = true; }
  if (!formData.lostPersonPhysical && profile.lostPersonPhysical) { formData.lostPersonPhysical = profile.lostPersonPhysical; changed = true; }

  // 2. Auto-fill Timestamps from logs
  const findLog = (actionPart) => {
    return bundle.activityLog.find(l => 
      (l.tag === taskTag || l.tag.startsWith(taskTag + ' - ')) && 
      l.action.toLowerCase().includes(actionPart.toLowerCase())
    );
  };

  const lLeave = findLog('leaving base') || findLog('leave base');
  if (lLeave) {
    const lTime = lLeave.time;
    if (formData.leaveBase !== lTime) {
      formData.leaveBase = lTime; 
      changed = true; 
    }
  }

  const lBegin = findLog('beginning search') || findLog('begin search') || findLog('beginning assignment') || findLog('begin assignment') || findLog('started search'); 
  if (lBegin) {
    const lTime = lBegin.time;
    if (formData.beginSearch !== lTime) {
      formData.beginSearch = lTime; 
      changed = true; 
    }
  }

  const lComplete = findLog('finished search') || findLog('finish search') || findLog('finished assignment') || findLog('finish assignment'); 
  if (lComplete) {
    const lTime = lComplete.time;
    if (formData.completeSearch !== lTime) {
      formData.completeSearch = lTime; 
      changed = true; 
    }
  }

  if (!formData.briefedBy) {
    const log = bundle.activityLog.find(l => 
      (l.tag === taskTag || l.tag.startsWith(taskTag + ' - ') || (l.tag.startsWith('base') && l.action.includes(taskTag))) && 
      (l.action.toLowerCase().includes('form auto-created') || l.action.toLowerCase().includes('form marked as completed'))
    );
    if (log) {
      const parts = log.tag.split(' - ');
      if (parts.length > 1) {
        formData.briefedBy = parts[1];
        changed = true;
      }
    }
  }
  
  const finishLog = findLog('finish') || findLog('complete') || findLog('returning to base');
  let arriveLog;
  if (finishLog) {
      const finishIdx = bundle.activityLog.indexOf(finishLog);
      arriveLog = [...bundle.activityLog.slice(0, finishIdx)].reverse().find(l => l.action.toLowerCase().includes('arrived at base') && l.team === formData.teamName);
  } else {
      arriveLog = bundle.activityLog.find(l => l.action.toLowerCase().includes('arrived at base') && l.team === formData.teamName);
  }
  if (arriveLog) {
     const lTime = arriveLog.time;
     if (formData.returnBase !== lTime) {
       formData.returnBase = lTime;
       changed = true;
     }
  }

  // 3. 20 Minute Status (Par Checks)
  const parCheckLogs = [...bundle.activityLog]
    .filter(l => (l.tag === taskTag || l.tag.startsWith(taskTag + ' - ')) && (l.action.toLowerCase().includes('par check') || l.action.toLowerCase().includes('check-in')))
    .reverse();
  
  // We'll store the log objects themselves to render later
  formData.parChecksRaw = parCheckLogs;

  if (changed) save();

  let currentCard = null;

  const addSection = (title) => {
    currentCard = document.createElement('div');
    currentCard.className = 'form-card';
    form.appendChild(currentCard);

    const sec = document.createElement('div');
    sec.className = 'form-section';
    sec.innerHTML = `<h3>${title}</h3>`;
    currentCard.appendChild(sec);
  };

  const addGroup = (label, key, type = 'text', readonly = false) => {
    const grp = document.createElement('div');
    grp.className = 'form-group';
    grp.dataset.key = key;
    const lbl = document.createElement('label');
    lbl.textContent = label;
    let inp;
    if (type === 'textarea') {
      inp = document.createElement('textarea');
      inp.style.minHeight = '60px';
    } else {
      inp = document.createElement('input');
      inp.type = type;
    }
    inp.className = 'form-input';
    inp.value = formData[key] || '';
    if (readonly) inp.readOnly = true;
    inp.oninput = () => {
      formData[key] = inp.value;
      save();
    };
    grp.appendChild(lbl);
    grp.appendChild(inp);
    (currentCard || form).appendChild(grp);
  };

  addSection('Incident Information');
  addGroup('Incident #', 'incidentNumber');
  addGroup('OP Period', 'opPeriod');
  addGroup('Date/Time', 'dateTime');
  addGroup('Task #', 'taskNumDisplay', 'text', true);
  formData.taskNumDisplay = '#' + taskNum;

  addSection('Lost Person Information');
  addGroup('Name', 'lostPersonName');
  addGroup('Age', 'lostPersonAge');
  addGroup('Gender', 'lostPersonGender');
  addGroup('Description', 'lostPersonDescription');
  addGroup('Clothing', 'lostPersonClothing');
  addGroup('Physical / Medical', 'lostPersonPhysical');
  
  const cbGroup = document.createElement('div');
  cbGroup.className = 'form-checkbox-group';
  cbGroup.style.gridColumn = 'span 4';
  
  const addCB = (label, key) => {
    const wrap = document.createElement('div');
    wrap.style.display = 'flex';
    wrap.style.alignItems = 'center';
    wrap.style.gap = '10px';
    wrap.style.background = 'rgba(0,0,0,0.1)';
    wrap.style.padding = '8px 15px';
    wrap.style.borderRadius = '8px';
    wrap.style.border = '1px solid var(--line)';
    
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'pill-checkbox';
    cb.id = 'cb-' + key + '-' + taskNum;
    cb.checked = !!formData[key];
    cb.onchange = () => {
      formData[key] = cb.checked;
      save();
    };
    
    const lbl = document.createElement('label');
    lbl.setAttribute('for', cb.id);
    lbl.style.fontWeight = 'bold';
    lbl.style.cursor = 'pointer';
    lbl.style.fontSize = '0.9rem';
    lbl.textContent = label;
    
    wrap.appendChild(cb);
    wrap.appendChild(lbl);
    cbGroup.appendChild(wrap);
  };
  addCB('On Scene: Family', 'onSceneFamily');
  addCB('On Scene: Media', 'onSceneMedia');
  currentCard.appendChild(cbGroup);

  addSection('Assignment Details');
  addGroup('Briefed By', 'briefedBy');
  addGroup('Radio #', 'radioNumber');
  addGroup('GPS #', 'gpsNumber');
  
  // Team Type Checklist
  const ttSection = document.createElement('div');
  ttSection.className = 'form-group';
  ttSection.dataset.key = 'teamTypes';
  ttSection.style.gridColumn = 'span 4';
  ttSection.innerHTML = '<label>Team Type (Check all that apply)</label>';
  const ttWrap = document.createElement('div');
  ttWrap.style.display = 'flex';
  ttWrap.style.gap = '15px';
  ttWrap.style.flexWrap = 'wrap';
  ttWrap.style.marginTop = '5px';
  
  const types = ['hasty', 'grid', 'area', 'k9', 'atv', 'argo', 'drone', 'boat', 'other'];
  types.forEach(type => {
    const wrap = document.createElement('div');
    wrap.style.display = 'flex';
    wrap.style.alignItems = 'center';
    wrap.style.gap = '10px';
    wrap.style.background = 'rgba(0,0,0,0.1)';
    wrap.style.padding = '6px 12px';
    wrap.style.borderRadius = '8px';
    wrap.style.border = '1px solid var(--line)';

    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.className = 'pill-checkbox';
    chk.id = 'chk-' + type + '-' + taskNum;
    chk.checked = !!(formData.teamTypes && formData.teamTypes[type]);
    chk.onchange = () => {
      if (!formData.teamTypes) formData.teamTypes = {};
      formData.teamTypes[type] = chk.checked;
      save();
    };
    
    const lbl = document.createElement('label');
    lbl.setAttribute('for', chk.id);
    lbl.style.fontWeight = 'bold';
    lbl.style.cursor = 'pointer';
    lbl.style.fontSize = '0.85rem';
    const labels = {
      hasty: 'Hasty',
      grid: 'Grid',
      area: 'Area',
      k9: 'K9',
      atv: 'ATV',
      argo: 'Argo',
      drone: 'Drone',
      boat: 'Boat',
      other: 'Other'
    };
    lbl.textContent = labels[type] || (type.charAt(0).toUpperCase() + type.slice(1));
    
    wrap.appendChild(chk);
    wrap.appendChild(lbl);
    ttWrap.appendChild(wrap);
  });
  ttSection.appendChild(ttWrap);
  currentCard.appendChild(ttSection);
  addGroup('Other Type Description', 'otherTeamType');

  addSection('Operational Times');
  addGroup('Leave Base', 'leaveBase');
  addGroup('Begin Search', 'beginSearch');
  addGroup('Complete Search', 'completeSearch');
  addGroup('Return Base', 'returnBase');
  
  const parSec = document.createElement('div');
  parSec.className = 'form-group';
  parSec.dataset.key = 'parChecks';
  parSec.style.gridColumn = 'span 4';
  parSec.innerHTML = '<label>20 Minute Status (Par Checks)</label>';
  currentCard.appendChild(parSec);
  const parWrap = document.createElement('div');
  parWrap.style.display = 'flex';
  parWrap.style.flexDirection = 'column';
  parWrap.style.gap = '8px';
  parWrap.style.marginTop = '10px';

  (formData.parChecksRaw || []).forEach(log => {
     const row = document.createElement('div');
     row.className = 'pill-cell clickable-pill';
     row.style.display = 'flex';
     row.style.alignItems = 'center';
     row.style.gap = '10px';
     row.style.padding = '8px 15px';
     row.style.marginBottom = '5px';
     row.style.width = '100%';
     row.style.boxSizing = 'border-box';
     
     // 1. Log (Action)
     const actionText = document.createElement('span');
     actionText.textContent = log.action;
     actionText.style.fontSize = '0.9rem';
     actionText.style.flex = '1';
     row.appendChild(actionText);

     // 2. Timestamp
     const timePill = document.createElement('div');
     timePill.className = 'pill-cell readonly-pill';
     timePill.style.background = 'rgba(255,255,255,0.1)';
     timePill.textContent = log.time;
     row.appendChild(timePill);
     
     // 3. Team Members (Team Names)
     if (log.members) {
       const mContainer = document.createElement('div');
       mContainer.style.display = 'flex';
       mContainer.style.gap = '5px';
       log.members.split(', ').forEach(m => {
         const mPill = document.createElement('div');
         mPill.className = 'mini-pill';
         mPill.style.fontSize = '0.75rem';
         mPill.style.cursor = 'default';
         if (m.endsWith('*')) {
           mPill.style.background = 'var(--pill-focus)';
           mPill.style.borderColor = 'var(--accent)';
         }
         mPill.textContent = m;
         mContainer.appendChild(mPill);
       });
       row.appendChild(mContainer);
     }
     
     row.onclick = () => {
        const popup = createPopup('Edit Log Entry');
        const content = popup.querySelector('.popup-content');
        const btnContainer = popup.querySelector('.popup-buttons');
        
        const area = document.createElement('textarea');
        area.className = 'form-input';
        area.style.minHeight = '100px';
        area.style.marginBottom = '20px';
        area.value = log.action;
        content.insertBefore(area, btnContainer);
        
        const saveBtn = document.createElement('button');
        saveBtn.className = 'popup-btn primary';
        saveBtn.textContent = 'Save';
        saveBtn.onclick = () => {
           const b = loadBundle();
           const found = b.activityLog.find(l => l.id === log.id);
           if (found) {
              found.action = area.value.trim();
              saveBundle(b);
              closePopup(popup);
              renderTaskForm(container, taskNum, formData);
           } else {
              // Fallback if no ID (for old logs)
              const oldFound = b.activityLog.find(l => l.time === log.time && l.team === log.team && l.action === log.action);
              if (oldFound) {
                 oldFound.action = area.value.trim();
                 saveBundle(b);
                 closePopup(popup);
                 renderTaskForm(container, taskNum, formData);
              }
           }
        };
        btnContainer.appendChild(saveBtn);
     };
     
     parWrap.appendChild(row);
  });
  parSec.appendChild(parWrap);
  
  addSection('Personnel');
  const pContainer = document.createElement('div');
  pContainer.className = 'form-personnel-list';
  pContainer.style.gridColumn = 'span 4';
  pContainer.style.display = 'flex';
  pContainer.style.flexDirection = 'column';
  pContainer.style.gap = '8px';
  pContainer.style.marginTop = '10px';
  
  const renderPersonnelRow = (m, idx) => {
    const row = document.createElement('div');
    row.className = 'pill-cell';
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.gap = '15px';
    row.style.padding = '8px 20px';
    row.style.width = '100%';
    row.style.boxSizing = 'border-box';
    
    if (m.leader) {
      row.style.background = 'var(--pill-focus)';
      row.style.borderColor = 'var(--accent)';
    }

    const nameSpan = document.createElement('span');
    nameSpan.contentEditable = 'true';
    nameSpan.textContent = m.name || '-- Empty --';
    nameSpan.style.flex = '1';
    nameSpan.style.fontWeight = '700';
    nameSpan.style.minWidth = '120px';
    nameSpan.onblur = () => {
      const oldName = m.name;
      m.name = nameSpan.textContent.trim();
      if (m.name !== oldName) {
        save();
        renderTaskForm(container, taskNum, formData);
      }
    };
    nameSpan.onkeydown = (e) => { 
      if(e.key === 'Enter') {
        e.preventDefault();
        nameSpan.blur();
      }
    };
    row.appendChild(nameSpan);

    const rolesWrap = document.createElement('div');
    rolesWrap.style.display = 'flex';
    rolesWrap.style.gap = '15px';
    rolesWrap.style.alignItems = 'center';

    const addRoleCB = (label, key) => {
      const wrap = document.createElement('div');
      wrap.style.display = 'flex';
      wrap.style.alignItems = 'center';
      wrap.style.gap = '6px';
      
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'pill-checkbox';
      cb.id = `role-${key}-${idx}-${taskNum}`;
      cb.checked = !!m[key];
      cb.onchange = () => { 
        m[key] = cb.checked; 
        save(); 
        if (key === 'leader') renderTaskForm(container, taskNum, formData);
      };
      
      const lbl = document.createElement('label');
      lbl.setAttribute('for', cb.id);
      lbl.textContent = label;
      lbl.style.fontSize = '0.8rem';
      lbl.style.fontWeight = '600';
      lbl.style.color = 'var(--muted)';
      lbl.style.cursor = 'pointer';

      wrap.appendChild(cb);
      wrap.appendChild(lbl);
      rolesWrap.appendChild(wrap);
    };

    addRoleCB('Leader', 'leader');
    addRoleCB('GPS', 'gps');
    addRoleCB('Radio', 'radio');
    addRoleCB('Medic', 'medic');
    
    row.appendChild(rolesWrap);

    const delBtn = document.createElement('button');
    delBtn.innerHTML = '&times;';
    delBtn.style.background = 'none';
    delBtn.style.border = 'none';
    delBtn.style.color = 'inherit';
    delBtn.style.cursor = 'pointer';
    delBtn.style.fontSize = '1.2rem';
    delBtn.style.lineHeight = '1';
    delBtn.style.padding = '0 5px';
    delBtn.style.marginLeft = '10px';
    delBtn.title = 'Remove Member';
    delBtn.onclick = (e) => {
       e.stopPropagation();
       formData.teamMembers.splice(idx, 1);
       save();
       renderTaskForm(container, taskNum, formData);
    };
    row.appendChild(delBtn);

    return row;
  };

  formData.teamMembers.forEach((m, idx) => {
    if (m.name) {
      pContainer.appendChild(renderPersonnelRow(m, idx));
    }
  });
  currentCard.appendChild(pContainer);

  // Add row button for personnel
  const addRowContainer = document.createElement('div');
  addRowContainer.style.gridColumn = 'span 4';
  addRowContainer.style.textAlign = 'right';
  addRowContainer.style.marginTop = '10px';
  const addPBtn = document.createElement('button');
  addPBtn.className = 'mini-pill';
  addPBtn.style.padding = '5px 15px';
  addPBtn.textContent = '+ Add Member';
  addPBtn.onclick = () => {
    // Show popup with all members not on this team
    const popup = createPopup('Add Team Member');
    const content = popup.querySelector('.popup-content');
    const btnContainer = popup.querySelector('.popup-buttons');
    
    const roster = loadBundle().pages.page3 || [];
    const currentMemberNames = formData.teamMembers.map(m => m.name);
    const availableMembers = roster.filter(m => m[0] && !currentMemberNames.includes(m[0]));
    
    const list = document.createElement('div');
    list.style.display = 'flex';
    list.style.flexWrap = 'wrap';
    list.style.gap = '10px';
    list.style.maxHeight = '300px';
    list.style.overflowY = 'auto';
    list.style.marginBottom = '20px';
    
    availableMembers.forEach(m => {
       const mBtn = document.createElement('button');
       mBtn.className = 'mini-pill';
       mBtn.textContent = m[0];
       mBtn.onclick = () => {
         formData.teamMembers.push({
           name: m[0],
           leader: m[0] === m[2],
           gps: m[3] === 'true',
           radio: m[4] === 'true',
           medic: m[5] === 'true'
         });
         save();
         popup.remove();
         renderTaskForm(container, taskNum, formData);
       };
       list.appendChild(mBtn);
    });
    
    if (availableMembers.length === 0) {
      list.innerHTML = '<p style="opacity: 0.6;">No more members available.</p>';
    }
    
    content.insertBefore(list, btnContainer);
  };
  addRowContainer.appendChild(addPBtn);
  currentCard.appendChild(addRowContainer);

  addSection('Assignment Instructions');
  const instGrp = document.createElement('div');
  instGrp.className = 'form-group';
  instGrp.dataset.key = 'instructions';
  instGrp.style.gridColumn = 'span 4';
  const instArea = document.createElement('textarea');
  instArea.className = 'form-input';
  instArea.style.minHeight = '150px';
  instArea.value = formData.instructions || '';
  instArea.oninput = () => { formData.instructions = instArea.value; save(); };
  instGrp.appendChild(instArea);
  currentCard.appendChild(instGrp);

  addSection('Form Completion');
  
  const compWrap = document.createElement('div');
  compWrap.className = 'form-group';
  compWrap.style.gridColumn = 'span 4';
  compWrap.style.display = 'flex';
  compWrap.style.alignItems = 'center';
  compWrap.style.gap = '12px';
  compWrap.style.background = 'rgba(125, 198, 255, 0.1)';
  compWrap.style.padding = '10px 20px';
  compWrap.style.borderRadius = '8px';
  compWrap.style.border = '1px solid var(--accent)';
  compWrap.style.width = 'fit-content';
  
  const compCheck = document.createElement('input');
  compCheck.type = 'checkbox';
  compCheck.className = 'pill-checkbox';
  compCheck.id = 'form-completed-check';
  compCheck.checked = !!formData.completed;
  
  const compLabel = document.createElement('label');
  compLabel.setAttribute('for', 'form-completed-check');
  compLabel.style.cursor = 'pointer';
  compLabel.style.fontWeight = '700';
  compLabel.textContent = formData.completedBy ? `Form Completed by ${formData.completedBy}` : 'Mark Form as Completed';
  
  compCheck.onchange = () => {
    if (compCheck.checked) {
      const user = getCurrentUser();
      const userName = user ? getAccountName(user) : 'Unknown User';

      formData.completed = true;
      formData.completedBy = userName;
      save();
      
      addActivityLogEntry(formData.teamName || 'N/A', `Form #${taskNum} marked as completed by ${userName}`);
      
      renderTaskForm(container, taskNum, formData);
      checkParChecksAndNotify(); // Refresh header if needed
      
      // Refresh Task Assignment pills immediately if on that page
      if (typeof buildTaskAssignmentForm === 'function') {
        buildTaskAssignmentForm();
      }
    } else {
      formData.completed = false;
      formData.completedBy = '';
      save();
      compLabel.textContent = 'Mark Form as Completed';
      checkParChecksAndNotify();
      
      // Refresh Task Assignment pills immediately if on that page
      if (typeof buildTaskAssignmentForm === 'function') {
        buildTaskAssignmentForm();
      }
    }
  };
  
  compLabel.onclick = () => compCheck.click();
  
  compWrap.appendChild(compCheck);
  compWrap.appendChild(compLabel);
  currentCard.appendChild(compWrap);

  container.appendChild(form);
}

const TASK_FORM_PRINT_STYLES = `
    @page { size: auto; margin: 10mm; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: 11pt; color: #000; background: #fff; margin: 0; padding: 0; }
    .print-container { max-width: 8.5in; margin: 0 auto; }
    .print-section { page-break-after: always; padding: 20px 0; }
    .print-section:last-child { page-break-after: auto; }
    
    h1 { font-size: 20pt; border-bottom: 2px solid #000; margin: 0 0 15px 0; padding-bottom: 5px; }
    h2 { font-size: 14pt; border-bottom: 1px solid #333; margin: 20px 0 10px 0; padding-bottom: 2px; }
    
    .task-form { border: 2px solid #000; padding: 15px; margin-bottom: 20px; position: relative; }
    .form-header { display: flex; justify-content: space-between; border-bottom: 2px solid #000; margin-bottom: 15px; padding-bottom: 5px; }
    .form-section { margin-bottom: 12px; }
    .form-section-title { background: #eee !important; -webkit-print-color-adjust: exact; font-weight: bold; padding: 3px 8px; margin-bottom: 8px; font-size: 11pt; border: 1px solid #ccc; }
    .form-row { display: flex; gap: 15px; margin-bottom: 8px; }
    .form-field { flex: 1; border-bottom: 1px solid #999; min-height: 1.2em; padding-bottom: 2px; }
    .field-label { font-size: 8pt; color: #444; font-weight: bold; display: block; margin-bottom: 1px; text-transform: uppercase; }
    .field-value { font-size: 11pt; }

    .par-check-item { border-bottom: 1px dotted #ccc; padding: 4px 0; display: flex; align-items: baseline; }
    .par-check-time { font-weight: bold; width: 60px; font-size: 9pt; color: #444; }
    .par-check-action { flex: 1; font-size: 10pt; }
    
    @media screen {
        body { background: #f0f2f5; padding: 40px 0; }
        .print-container { background: #fff; padding: 40px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        .no-print { display: block; text-align: center; margin-bottom: 20px; }
    }
    @media print {
        .no-print { display: none; }
    }
`;

function getTaskFormPrintHTML(num, f) {
    const members = (f.teamMembers || []).map(m => `${m.name} ${m.leader ? '(L)' : ''}`).join(', ');

    // Team Types string
    const activeTypes = [];
    if (f.teamTypes) {
        Object.entries(f.teamTypes).forEach(([type, active]) => {
            if (active) {
                activeTypes.push(type === 'other' ? (f.otherTeamType || 'Other') : type.toUpperCase());
            }
        });
    }
    const teamTypeStr = activeTypes.length > 0 ? activeTypes.join(', ') : '';

    return `
                <div class="task-form">
                    <div class="form-header">
                        <span style="font-weight: bold; font-size: 16pt;">Task Assignment Form</span>
                        <span style="font-weight: bold; font-size: 16pt;">Task # ${num}</span>
                    </div>
                    
                    <div class="form-section">
                        <div class="form-section-title">1. INCIDENT OVERVIEW</div>
                        <div class="form-row">
                            <div class="form-field"><span class="field-label">Incident Name</span><div class="field-value">${f.incidentNumber || ''}</div></div>
                            <div class="form-field" style="flex:0.3;"><span class="field-label">Op Period</span><div class="field-value">${f.opPeriod || ''}</div></div>
                            <div class="form-field" style="flex:0.5;"><span class="field-label">Date/Time</span><div class="field-value">${f.dateTime || ''}</div></div>
                        </div>
                    </div>

                    <div class="form-section">
                        <div class="form-section-title">2. SUBJECT INFORMATION</div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Subject Name</span><div class="field-value">${f.lostPersonName || ''}</div></div>
                             <div class="form-field" style="flex:0.2;"><span class="field-label">Age</span><div class="field-value">${f.lostPersonAge || ''}</div></div>
                             <div class="form-field" style="flex:0.2;"><span class="field-label">Gender</span><div class="field-value">${f.lostPersonGender || ''}</div></div>
                        </div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Description / Clothing</span><div class="field-value">${f.lostPersonDescription || ''} ${f.lostPersonClothing || ''}</div></div>
                        </div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Physical / Medical Information</span><div class="field-value">${f.lostPersonPhysical || ''}</div></div>
                             <div class="form-field" style="flex:0.4;"><span class="field-label">On Scene</span><div class="field-value">${[f.onSceneFamily ? 'Family' : '', f.onSceneMedia ? 'Media' : ''].filter(Boolean).join(', ') || 'None'}</div></div>
                        </div>
                    </div>

                    <div class="form-section">
                        <div class="form-section-title">3. ASSIGNMENT DETAILS</div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Region/Segment</span><div class="field-value">${f.segment || ''}</div></div>
                             <div class="form-field"><span class="field-label">Team ID</span><div class="field-value">${f.teamName || ''}</div></div>
                             <div class="form-field" style="flex:0.6;"><span class="field-label">Team Type</span><div class="field-value">${teamTypeStr}</div></div>
                        </div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Task Objective</span><div class="field-value">${f.instructions || ''}</div></div>
                        </div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Briefed By</span><div class="field-value">${f.briefedBy || ''}</div></div>
                             <div class="form-field" style="flex:0.3;"><span class="field-label">Radio #</span><div class="field-value">${f.radioNumber || ''}</div></div>
                             <div class="form-field" style="flex:0.3;"><span class="field-label">GPS #</span><div class="field-value">${f.gpsNumber || ''}</div></div>
                        </div>
                    </div>
                    
                    <div class="form-section">
                        <div class="form-section-title">4. PERSONNEL</div>
                        <div class="form-field"><span class="field-label">Team Members</span><div class="field-value">${members}</div></div>
                    </div>

                    <div class="form-section">
                        <div class="form-section-title">5. TIMESTAMPS</div>
                        <div class="form-row">
                             <div class="form-field"><span class="field-label">Leave Base</span><div class="field-value">${f.leaveBase || ''}</div></div>
                             <div class="form-field"><span class="field-label">Begin Search</span><div class="field-value">${f.beginSearch || ''}</div></div>
                             <div class="form-field"><span class="field-label">Finish Search</span><div class="field-value">${f.completeSearch || ''}</div></div>
                             <div class="form-field"><span class="field-label">Return Base</span><div class="field-value">${f.returnBase || ''}</div></div>
                        </div>
                    </div>

                    <div class="form-section">
                        <div class="form-section-title">6. PAR CHECKS (20 MINUTE STATUS)</div>
                        <div id="par-checks-${num}">
                            ${(f.parChecksRaw || []).length > 0 ? 
                                f.parChecksRaw.map(l => '<div class="par-check-item"><div class="par-check-time">' + l.time + '</div><div class="par-check-action">' + l.action + '</div></div>').join('') : 
                                '<div style="font-style: italic; color: #888; padding: 5px;">No par checks recorded for this task.</div>'
                            }
                        </div>
                    </div>
                </div>
    `;
}

function printSingleTaskForm(taskNum) {
    const bundle = loadBundle();
    const f = bundle.forms ? bundle.forms[taskNum] : null;
    if (!f) {
        alert("Form data not found.");
        return;
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        alert("Please allow popups to view the printout.");
        return;
    }

    printWindow.document.write(`
<!DOCTYPE html>
<html>
<head>
    <title>Task Assignment Form - Task #${taskNum}</title>
    <style>${TASK_FORM_PRINT_STYLES}</style>
</head>
<body>
    <div class="no-print">
        <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 999px;">Print PDF</button>
        <p style="font-size: 12px; color: #666;">Note: Use "Save as PDF" in the print dialog for a digital copy.</p>
    </div>
    <div class="print-container">
        <div class="print-section">
            ${getTaskFormPrintHTML(taskNum, f)}
        </div>
    </div>
    <script>
        setTimeout(() => { window.print(); }, 500);
    </script>
</body>
</html>
    `);
    printWindow.document.close();
}

function downloadAllForms() {
  const bundle = loadBundle();
  const forms = bundle.forms || {};
  const taskNums = Object.keys(forms).sort((a,b) => parseInt(a) - parseInt(b));
  
  if (taskNums.length === 0) {
    alert("No forms found to download.");
    return;
  }

  const printWindow = window.open('', '_blank');
  if (!printWindow) {
    alert("Please allow popups to view the printout.");
    return;
  }

  let formsHTML = '';
  taskNums.forEach(num => {
    formsHTML += `<div class="print-section">${getTaskFormPrintHTML(num, forms[num])}</div>`;
  });

  printWindow.document.write(`
<!DOCTYPE html>
<html>
<head>
    <title>All Task Assignment Forms</title>
    <style>${TASK_FORM_PRINT_STYLES}</style>
</head>
<body>
    <div class="no-print">
        <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 999px;">Print PDF</button>
        <p style="font-size: 12px; color: #666;">Note: Use "Save as PDF" in the print dialog for a digital copy.</p>
    </div>
    <div class="print-container">
        ${formsHTML}
    </div>
    <script>
        setTimeout(() => { window.print(); }, 500);
    </script>
</body>
</html>
  `);
  printWindow.document.close();
}

function printSearchFile() {
    recalculateEverything();
    const bundle = loadBundle();
    const fileName = (bundle.fileName || "Search_File").replace('.json', '');
    
    // Calculate metrics for charts using current chart settings
    const targetDateStr = selectedChartDate;
    const durationHours = selectedChartDuration;
    const metrics = calculateHourlyMetrics(targetDateStr, durationHours);
    const psrcData = metrics.map(m => m.totalPSRc);
    const posData = metrics.map(m => m.totalPOS);

    const forms = bundle.forms || {};
    const taskFormsHTML = Object.keys(forms)
        .sort((a, b) => parseInt(a) - parseInt(b))
        .map(num => `<div class="print-section">${getTaskFormPrintHTML(num, forms[num])}</div>`)
        .join('');

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        alert("Please allow popups to view the printout.");
        return;
    }

    printWindow.document.write(`
<!DOCTYPE html>
<html>
<head>
    <title>Search File Printout - ${fileName}</title>
    <style>
        @page { size: auto; margin: 10mm; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: 11pt; color: #000; background: #fff; margin: 0; padding: 0; }
        .print-container { max-width: 8.5in; margin: 0 auto; }
        .print-section { page-break-after: always; padding: 20px 0; }
        .print-section:last-child { page-break-after: auto; }
        
        h1 { font-size: 20pt; border-bottom: 2px solid #000; margin: 0 0 15px 0; padding-bottom: 5px; }
        h2 { font-size: 14pt; border-bottom: 1px solid #333; margin: 20px 0 10px 0; padding-bottom: 2px; }
        
        .search-log-table { width: 100%; border-collapse: collapse; font-size: 10pt; margin-bottom: 20px; }
        .search-log-table th, .search-log-table td { border: 1px solid #000; padding: 4px 6px; text-align: left; }
        .search-log-table th { background: #eee !important; -webkit-print-color-adjust: exact; }
        
        .activity-log { font-size: 10pt; line-height: 1.0; }
        .activity-log-entry { margin-bottom: 1px; }
        .activity-log-time { font-weight: bold; margin-right: 5px; }

        .charts-container { display: flex; gap: 20px; margin-bottom: 20px; height: 180px; }
        .chart-item { flex: 1; display: flex; flex-direction: column; border: 1px solid #ddd; padding: 10px; border-radius: 8px; }
        .chart-label { font-size: 9pt; font-weight: bold; margin-bottom: 5px; color: #555; text-align: center; }
        .chart-svg-wrap { flex: 1; min-height: 140px; }

        .task-form { border: 2px solid #000; padding: 15px; margin-bottom: 20px; position: relative; }
        .form-header { display: flex; justify-content: space-between; border-bottom: 2px solid #000; margin-bottom: 15px; padding-bottom: 5px; }
        .form-section { margin-bottom: 12px; }
        .form-section-title { background: #eee !important; -webkit-print-color-adjust: exact; font-weight: bold; padding: 3px 8px; margin-bottom: 8px; font-size: 11pt; border: 1px solid #ccc; }
        .form-row { display: flex; gap: 15px; margin-bottom: 8px; }
        .form-field { flex: 1; border-bottom: 1px solid #999; min-height: 1.2em; padding-bottom: 2px; }
        .field-label { font-size: 8pt; color: #444; font-weight: bold; display: block; margin-bottom: 1px; text-transform: uppercase; }
        .field-value { font-size: 11pt; }

        .par-check-item { border-bottom: 1px dotted #ccc; padding: 4px 0; display: flex; align-items: baseline; }
        .par-check-time { font-weight: bold; width: 60px; font-size: 9pt; color: #444; }
        .par-check-action { flex: 1; font-size: 10pt; }
        
        @media screen {
            body { background: #f0f2f5; padding: 40px 0; }
            .print-container { background: #fff; padding: 40px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
            .no-print { display: block; text-align: center; margin-bottom: 20px; }
        }
        @media print {
            .no-print { display: none; }
        }
    </style>
</head>
<body>
    <div class="no-print">
        <button onclick="window.print()" style="padding: 10px 20px; font-size: 16px; cursor: pointer; background: #007bff; color: white; border: none; border-radius: 999px;">Print PDF</button>
        <p style="font-size: 12px; color: #666;">Note: Use "Save as PDF" in the print dialog for a digital copy.</p>
    </div>

    <div class="print-container">
        <!-- Search Log & Charts -->
        <div class="print-section">
            <h1>Search Log: ${fileName}</h1>
            
            <div class="charts-container">
                <div class="chart-item">
                    <div class="chart-label">PSRc Sum Chart (${targetDateStr})</div>
                    <div id="psrc-chart" class="chart-svg-wrap"></div>
                </div>
                <div class="chart-item">
                    <div class="chart-label">POS Sum Chart (${targetDateStr})</div>
                    <div id="pos-chart" class="chart-svg-wrap"></div>
                </div>
            </div>

            <table class="search-log-table">
                <thead>
                    <tr>
                        <th>Task #</th><th>Date</th><th>Time</th><th>Region</th><th>Segment</th>
                        <th>PSR Before</th><th>PSR After</th><th>Team</th><th>Width</th><th>Sweeps</th>
                    </tr>
                </thead>
                <tbody id="sl-body"></tbody>
            </table>
        </div>

        <!-- Activity Log -->
        <div class="print-section">
            <h1>Activity Log</h1>
            <div id="al-body" class="activity-log"></div>
        </div>

        <!-- Task Forms -->
        <div id="forms-body">
            ${taskFormsHTML}
        </div>
    </div>

    <script>
        const bundle = ${JSON.stringify(bundle)};
        const psrcData = ${JSON.stringify(psrcData)};
        const posData = ${JSON.stringify(posData)};
        const durationHours = ${durationHours};

        // Render Search Log
        const slBody = document.getElementById('sl-body');
        (bundle.pages.page4 || []).forEach(row => {
            const tr = document.createElement('tr');
            for(let i=0; i<10; i++) {
                const td = document.createElement('td');
                td.textContent = row[i] || '';
                tr.appendChild(td);
            }
            slBody.appendChild(tr);
        });

        // Render Activity Log
        const alBody = document.getElementById('al-body');
        (bundle.activityLog || []).forEach(entry => {
            const tagPart = entry.tag || 'base';
            const displayTag = tagPart.split(' - ')[0];
            const userName = tagPart.includes(' - ') ? tagPart.split(' - ')[1] : '';
            const div = document.createElement('div');
            div.className = 'activity-log-entry';
            div.innerHTML = \`<span class="activity-log-time">[\${entry.date} \${entry.time} \${displayTag}\${userName ? ' (' + userName + ')' : ''}]</span> Team \${entry.team} (\${entry.members}): \${entry.action}\`;
            alBody.appendChild(div);
        });

        // Charts
        function formatHourOffset(offset) {
            const hrs = Math.floor(offset);
            return \`\${hrs.toString().padStart(2, '0')}\`;
        }

        function drawLineChart(containerId, data, color, isPOS) {
            const container = document.getElementById(containerId);
            const width = container.clientWidth || 350;
            const height = container.clientHeight || 140;
            const padding = { top: 10, right: 10, bottom: 30, left: 45 };
            const chartWidth = width - padding.left - padding.right;
            const chartHeight = height - padding.top - padding.bottom;
            const max = Math.max(...data, isPOS ? 0.01 : 1);

            const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
            svg.setAttribute("width", "100%");
            svg.setAttribute("height", "100%");
            svg.setAttribute("viewBox", \`0 0 \${width} \${height}\`);
            
            // Y-axis ticks
            for (let i = 0; i <= 4; i++) {
                const val = (max / 4) * i;
                const y = height - padding.bottom - (val / max) * chartHeight;
                const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
                text.setAttribute("x", padding.left - 5);
                text.setAttribute("y", y + 3);
                text.setAttribute("text-anchor", "end");
                text.setAttribute("font-size", "8");
                text.setAttribute("fill", "#666");
                text.textContent = isPOS ? (val * 100).toFixed(0) + '%' : val.toFixed(1);
                svg.appendChild(text);
                
                const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
                line.setAttribute("x1", padding.left); line.setAttribute("y1", y);
                line.setAttribute("x2", width - padding.right); line.setAttribute("y2", y);
                line.setAttribute("stroke", "#eee");
                svg.appendChild(line);
            }

            // X-axis ticks
            const numXTicks = 8;
            for (let i = 0; i <= numXTicks; i++) {
                const x = padding.left + (i / numXTicks) * chartWidth;
                const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
                text.setAttribute("x", x);
                text.setAttribute("y", height - 10);
                text.setAttribute("text-anchor", "middle");
                text.setAttribute("font-size", "8");
                text.setAttribute("fill", "#666");
                text.textContent = formatHourOffset((durationHours / numXTicks) * i);
                svg.appendChild(text);
            }

            // Path
            if (data.length > 1) {
                const points = data.map((val, i) => ({
                    x: padding.left + (i / (data.length - 1)) * chartWidth,
                    y: height - padding.bottom - (val / max) * chartHeight
                }));
                const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
                let d = \`M \${points[0].x} \${points[0].y}\`;
                for (let i = 0; i < points.length - 1; i++) {
                    const curr = points[i]; const next = points[i + 1];
                    const cp1x = curr.x + (next.x - curr.x) / 3;
                    const cp2x = curr.x + 2 * (next.x - curr.x) / 3;
                    d += \` C \${cp1x} \${curr.y}, \${cp2x} \${next.y}, \${next.x} \${next.y}\`;
                }
                path.setAttribute("d", d); path.setAttribute("fill", "none");
                path.setAttribute("stroke", color); path.setAttribute("stroke-width", "2");
                svg.appendChild(path);
            }
            container.appendChild(svg);
        }
        drawLineChart('psrc-chart', psrcData, '#007bff', false);
        drawLineChart('pos-chart', posData, '#fd7e14', true);

        // Auto-print
        setTimeout(() => {
            window.print();
        }, 1000);
    </script>
</body>
</html>
    `);
    printWindow.document.close();
}

function buildManageFormsTable() {
  const tableHead = document.getElementById('table-head');
  const tableBody = document.getElementById('table-body');
  if (!tableHead || !tableBody) return;

  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  const forms = bundle.forms || {};

  const unfinishedTasks = new Set();
  if (bundle.currentAssignments && bundle.teamStatuses) {
    for (const team in bundle.currentAssignments) {
      const status = bundle.teamStatuses[team] || '';
      const assignment = bundle.currentAssignments[team] || '';
      if (!status.includes('at base') && assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
        const match = assignment.match(/#(\d+)/);
        if (match) unfinishedTasks.add(match[1]);
      }
    }
  }
  
  const taskMap = new Map();
  searchLog.forEach(row => {
    if (row[0] && row[0].startsWith('#')) {
      const num = row[0].substring(1);
      taskMap.set(num, {
        num: num,
        timestamp: row[1] + ' ' + (row[2] || ''),
        region: row[3],
        segment: row[4]
      });
    }
  });

  // Also include forms that might not be in search log (if any manual additions)
  Object.keys(forms).forEach(num => {
    if (!taskMap.has(num)) {
       taskMap.set(num, {
         num: num,
         timestamp: forms[num].dateTime || 'Manual',
         region: 'Manual',
         segment: 'Manual'
       });
    }
  });

  let tasks = Array.from(taskMap.values()).sort((a, b) => parseInt(a.num) - parseInt(b.num));

  if (highlightedRowIndex === -2 && window.lastAddedTaskNum) {
      highlightedRowIndex = tasks.findIndex(t => t.num === window.lastAddedTaskNum);
      window.lastAddedTaskNum = null;
  }

  // Requirement: "should have only one empty row" if no tasks
  if (tasks.length === 0) {
     tasks.push({ num: '?', timestamp: '', region: 'N/A', segment: 'N/A', isEmpty: true });
  }

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  const headers = ['Task #', 'Region/Segment', 'Timestamp', 'Edit', 'Download'];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    th.className = 'fixed-header';
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  tasks.forEach((task, idx) => {
    const tr = document.createElement('tr');
    animateNewRow(tr, idx);

    const tdTask = document.createElement('td');
    tdTask.setAttribute('data-label', 'Task #');
    const taskCell = document.createElement('div');
    taskCell.className = 'pill-cell readonly-pill';
    taskCell.textContent = task.isEmpty ? '-' : '#' + task.num;
    tdTask.appendChild(taskCell);
    tr.appendChild(tdTask);

    const tdInfo = document.createElement('td');
    tdInfo.setAttribute('data-label', 'Region/Segment');
    const infoCell = document.createElement('div');
    infoCell.className = 'pill-cell readonly-pill';
    infoCell.textContent = task.isEmpty ? '-' : `${task.region} - ${task.segment}`;
    tdInfo.appendChild(infoCell);
    tr.appendChild(tdInfo);

    const tdTime = document.createElement('td');
    tdTime.setAttribute('data-label', 'Timestamp');
    const timeCell = document.createElement('div');
    timeCell.className = 'pill-cell readonly-pill';
    timeCell.textContent = task.timestamp || '-';
    tdTime.appendChild(timeCell);
    tr.appendChild(tdTime);

    const tdEdit = document.createElement('td');
    tdEdit.setAttribute('data-label', 'Edit');
    const editBtn = document.createElement('button');
    editBtn.className = 'row-delete-btn update-pill';
    editBtn.textContent = 'Edit';
    if (task.isEmpty) editBtn.disabled = true;

    // Apply color highlighting
    const isUnfinished = unfinishedTasks.has(task.num);
    const form = forms[task.num];
    const isComplete = form && form.completed;
    if (isUnfinished) {
       editBtn.style.opacity = '0.5';
       editBtn.style.background = 'rgba(128, 128, 128, 0.2)';
    } else if (isComplete) {
       editBtn.style.background = 'rgba(76, 175, 80, 0.25)';
       editBtn.style.borderColor = 'rgba(76, 175, 80, 0.5)';
       editBtn.style.color = '#4caf50';
    } else if (!isComplete) {
       editBtn.style.background = 'rgba(255, 140, 0, 0.25)';
       editBtn.style.borderColor = 'rgba(255, 140, 0, 0.5)';
       editBtn.style.color = '#ff8c00';
    }

    editBtn.onclick = () => {
       currentTaskNumber = task.num;
       currentFormsSubpage = 'task-assignment';
       buildFormsPage();
    };
    tdEdit.appendChild(editBtn);
    tr.appendChild(tdEdit);

    const tdDown = document.createElement('td');
    tdDown.setAttribute('data-label', 'Download');
    const downBtn = document.createElement('button');
    downBtn.className = 'row-delete-btn update-pill';
    downBtn.textContent = 'Download';
    if (task.isEmpty) downBtn.disabled = true;

    if (isUnfinished) {
       downBtn.style.opacity = '0.5';
       downBtn.style.background = 'rgba(128, 128, 128, 0.2)';
    } else if (isComplete) {
       downBtn.style.background = 'rgba(76, 175, 80, 0.25)';
       downBtn.style.borderColor = 'rgba(76, 175, 80, 0.5)';
       downBtn.style.color = '#4caf50';
    } else if (!isComplete) {
       downBtn.style.background = 'rgba(255, 140, 0, 0.25)';
       downBtn.style.borderColor = 'rgba(255, 140, 0, 0.5)';
       downBtn.style.color = '#ff8c00';
    }

    downBtn.onclick = () => {
       const oldTitle = document.title;
       document.title = `SAR_TASK_ASSIGNMENT_FORM_Task_${task.num}`;
       currentTaskNumber = task.num;
       currentFormsSubpage = 'task-assignment';
       buildFormsPage();
       setTimeout(() => {
           window.print();
           document.title = oldTitle;
       }, 500);
    };
    tdDown.appendChild(downBtn);
    tr.appendChild(tdDown);

    tableBody.appendChild(tr);
  });

  // Requirement: "be able to add rows"
  const addRowContainer = document.createElement('div');
  addRowContainer.className = 'add-row-container';
  const addRowBtn = document.createElement('button');
  addRowBtn.className = 'add-row-btn';
  addRowBtn.textContent = '+ Create New Task Form Manually';
  addRowBtn.onclick = () => {
    const nextNum = getNextTaskNumber();
    const b = loadBundle();
    if (!b.forms) b.forms = {};
    const now = new Date();
    const dateStr = `${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}-${now.getFullYear()} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    
    b.forms[nextNum] = {
      incidentNumber: '',
      opPeriod: '1',
      dateTime: dateStr,
      lostPersonName: '',
      lostPersonAge: '',
      lostPersonGender: '',
      lostPersonDescription: '',
      lostPersonClothing: '',
      lostPersonPhysical: '',
      onSceneFamily: false,
      onSceneMedia: false,
      briefedBy: '',
      radioNumber: '',
      gpsNumber: '',
      leaveBase: '',
      beginSearch: '',
      completeSearch: '',
      returnBase: '',
      teamType: 'Manual Entry',
      instructions: '',
      teamMembers: Array.from({length: 8}, () => ({name: '', leader: false, gps: false, radio: false, medic: false})),
      statusUpdates: Array.from({length: 8}, () => ({time: '', clue: '', usng: ''}))
    };
    logCreation('Task Form', '#' + nextNum, b);
    saveBundle(b);
    highlightedRowIndex = -2;
    window.lastAddedTaskNum = nextNum;
    buildManageFormsTable();
  };
  addRowContainer.appendChild(addRowBtn);
  const existing = document.querySelector('.add-row-container');
  if (existing) existing.remove();
  tableBody.parentElement.after(addRowContainer);
}

function isTaskUnfinished(taskWithHash) {
  const bundle = loadBundle();
  if (!bundle.currentAssignments || !bundle.teamStatuses) return false;
  for (const team in bundle.currentAssignments) {
    const status = bundle.teamStatuses[team] || '';
    const assignment = bundle.currentAssignments[team] || '';
    if (!status.includes('at base') && assignment !== 'Base' && assignment !== 'None' && assignment !== '') {
      const match = assignment.match(/#(\d+)/);
      if (match && '#' + match[1] === taskWithHash) return true;
    }
  }
  return false;
}

function getLogSweepsDue() {
  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  const due = [];
  searchLog.forEach(row => {
    const taskNum = row[0];
    const numSweeps = row[9];
    if ((!numSweeps || numSweeps.toString().trim() === '') && !isTaskUnfinished(taskNum)) {
      due.push({
        taskNum: taskNum,
        region: row[3],
        segment: row[4],
        fullRow: row
      });
    }
  });
  return due;
}

function showLogSweepsPopup(taskNumWithHash) {
  const bundle = loadBundle();
  const searchLog = bundle.pages.page4 || [];
  const row = searchLog.find(r => r[0] === taskNumWithHash);
  if (!row) return;

  const popup = createPopup('Log Sweeps');
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');
  
  const info = document.createElement('div');
  info.style.marginBottom = '20px';
  info.style.display = 'flex';
  info.style.flexDirection = 'column';
  info.style.gap = '8px';
  
  // Create a display similar to the search log row
  const rowData = [
    { label: 'Task #', val: row[0] },
    { label: 'Date', val: row[1] },
    { label: 'Time', val: row[2] },
    { label: 'Region', val: row[3] },
    { label: 'Segment', val: row[4] },
    { label: 'Team', val: row[7] }
  ];
  
  rowData.forEach(d => {
    const item = document.createElement('div');
    item.style.display = 'flex';
    item.style.justifyContent = 'space-between';
    item.innerHTML = `<span style="color:var(--muted);">${d.label}:</span> <span>${d.val}</span>`;
    info.appendChild(item);
  });
  
  const inputGrp = document.createElement('div');
  inputGrp.style.marginTop = '15px';
  inputGrp.innerHTML = `<label style="display:block; margin-bottom: 5px; font-weight: bold;">Number of Sweeps:</label>`;
  const sweepInput = document.createElement('input');
  sweepInput.type = 'number';
  sweepInput.className = 'pill-input';
  sweepInput.style.width = '100%';
  sweepInput.placeholder = 'Enter sweep count';
  inputGrp.appendChild(sweepInput);
  info.appendChild(inputGrp);
  
  content.insertBefore(info, btnContainer);
  
  const submitBtn = document.createElement('button');
  submitBtn.className = 'popup-btn primary';
  submitBtn.textContent = 'Submit';
  submitBtn.onclick = () => {
    const val = sweepInput.value.trim();
    if (!val) {
      alert("Please enter a sweep count.");
      return;
    }
    
    const b = loadBundle();
    const log = b.pages.page4 || [];
    const target = log.find(r => r[0] === taskNumWithHash);
    if (target) {
      target[9] = val; // Num of Sweeps is at index 9
      saveBundle(b);
      
      // Update local sortedData if we are on the segments page to avoid waiting for full rebuild if needed
      // Actually, refreshCurrentPageTable will handle it, but let's make sure it's smooth.
      
      recalculateEverything(); // Trigger cascading PSR recalculation
      closePopup(popup);
      
      // If we are on segments page, we can try to find the row and animate it out before refresh
      if (isSegmentsPage()) {
          const row = Array.from(document.querySelectorAll('#table-body tr')).find(tr => {
              const cells = tr.querySelectorAll('.pill-cell');
              return cells.length >= 2 && cells[0].textContent === target[3] && cells[1].textContent === target[4];
          });
          if (row) {
              row.style.transition = 'all 0.4s ease';
              row.classList.remove('log-sweeps-due');
              const btn = row.querySelector('.row-search-btn');
              if (btn) {
                  btn.style.transition = 'all 0.4s ease';
                  btn.classList.remove('log-sweeps-active');
                  btn.textContent = 'search';
              }
              // Wait for transition before full refresh to keep it smooth
              setTimeout(() => {
                  refreshCurrentPageTable();
              }, 400);
          } else {
              refreshCurrentPageTable();
          }
      } else {
          refreshCurrentPageTable();
      }
    }
  };
  btnContainer.appendChild(submitBtn);
}

function checkParChecksAndNotify(skipTableRefresh = false) {
  const bundle = loadBundle();
  const now = Date.now();
  const freqMs = (bundle.parCheckFrequency || 20) * 60 * 1000;
  
  if (!window._notifiedParChecks) window._notifiedParChecks = {};

  const data = bundle.pages.page3 || [];
  const teamsMap = new Map();
  const baseTeamNames = ['Base Support', 'Off Duty', 'Command'];
  
  data.forEach(row => {
    if (row[1] && row[6] === 'true' && !baseTeamNames.includes(row[1])) {
      teamsMap.set(row[1], true);
    }
  });

  let totalDue = 0;

  teamsMap.forEach((_, teamName) => {
    const status = bundle.teamStatuses[teamName] || '';
    if (status.startsWith('at base')) return; // Skip teams at base

    const lastPar = bundle.parChecks?.[teamName];
    const leaveTime = bundle.teamLeaveTimes?.[teamName];
    const assignTime = bundle.teamAssignmentTimes?.[teamName];
    
    let startTime = 0;
    if (lastPar) startTime = Math.max(startTime, lastPar.lastTime);
    if (leaveTime) startTime = Math.max(startTime, leaveTime);
    if (assignTime) startTime = Math.max(startTime, assignTime);

    if (startTime > 0 && (now - startTime) >= freqMs) {
      totalDue++;
      const lastCheckTime = lastPar 
        ? new Date(lastPar.lastTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        : (new Date(startTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
      
      const lastNote = lastPar?.lastNote || 'No previous par check note.';
      
      const notifyKey = teamName + '_' + (lastPar?.lastTime || startTime);
      if (!window._notifiedParChecks[notifyKey]) {
        if (typeof Notification !== 'undefined' && Notification.permission === "granted") {
          try {
            const membersData = getTeamMembers(teamName);
            const leadRow = membersData.find(m => m[0] === m[2]);
            const leadName = leadRow ? leadRow[0] : 'Unknown Lead';
            const otherMembers = membersData.filter(m => m[0] !== leadName).map(m => m[0]);
            
            let membersListStr = leadName;
            if (otherMembers.length > 0) {
              membersListStr += ` and ${otherMembers.join(', ')}`;
            }
            
            // Get last log entry for this team's task assignment
            let latestLogEntry = 'No log entry found.';
            if (bundle.activityLog) {
              const currentAssignment = bundle.currentAssignments[teamName] || '';
              const match = currentAssignment.match(/#(\d+)/);
              const currentTag = match ? match[0] : 'base';
              
              const teamLogs = bundle.activityLog.filter(l => l.team === teamName && (l.tag === currentTag || l.tag.startsWith(currentTag + ' - ')));
              if (teamLogs.length > 0) {
                latestLogEntry = teamLogs[0].action;
              }
            }

            const title = `Par Check Due`;
            const body = `${teamName}, ${membersListStr} Par Check is due; latest update was ${latestLogEntry}`;
            
            new Notification(title, { body: body });
            
            // Also show custom toast
            showToastNotification(title, body, () => navigateToPage('page3.html'), 'par-check-due');
            
            window._notifiedParChecks[notifyKey] = true;
          } catch (e) {}
        }
      }
    }
  });

  const parCheckDot = document.getElementById('par-check-dot');
  if (parCheckDot) {
    if (totalDue > 0) {
      parCheckDot.classList.add('active');
    } else {
      parCheckDot.classList.remove('active');
    }
  }

  const navPersonnel = document.getElementById('nav-personnel');
  if (navPersonnel) {
    if (totalDue > 0) {
      navPersonnel.title = `Par Checks Due - ${totalDue}`;
      navPersonnel.classList.add('par-check-due');
    } else {
      navPersonnel.title = 'Personnel';
      navPersonnel.classList.remove('par-check-due');
    }
  }

  const navSearchLog = document.getElementById('nav-search-log');
  if (navSearchLog) {
    const logSweepsDue = getLogSweepsDue();
    if (logSweepsDue.length > 0) {
      navSearchLog.title = 'Log Sweeps';
      navSearchLog.classList.add('log-sweeps-due');
    } else {
      navSearchLog.title = 'Search Log';
      navSearchLog.classList.remove('log-sweeps-due');
    }
  }

  const navProfile = document.getElementById('nav-profile');
  if (navProfile) {
    if (!bundle.profile || !bundle.profile.completed) {
      navProfile.classList.add('profile-incomplete');
    } else {
      navProfile.classList.remove('profile-incomplete');
    }
  }

  const downloadAllBtn = document.getElementById('download-all-forms-btn');
  if (downloadAllBtn) {
    const forms = bundle.forms || {};
    const taskNums = Object.keys(forms);
    let allCompleted = taskNums.length > 0;
    taskNums.forEach(num => {
      if (!forms[num].completed) allCompleted = false;
    });

    if (taskNums.length === 0) {
      downloadAllBtn.innerHTML = 'Download All Forms';
      downloadAllBtn.className = 'sub-nav-btn';
      downloadAllBtn.style.backgroundColor = '';
      downloadAllBtn.style.color = '';
      downloadAllBtn.style.borderColor = '';
    } else if (allCompleted) {
      downloadAllBtn.innerHTML = '✓ Download All Forms';
      downloadAllBtn.className = 'sub-nav-btn active';
      downloadAllBtn.style.background = 'rgba(46, 204, 113, 0.15)';
      downloadAllBtn.style.color = '#2ecc71';
      downloadAllBtn.style.borderColor = 'rgba(46, 204, 113, 0.45)';
    } else {
      downloadAllBtn.innerHTML = '⚠️ Download All Forms';
      downloadAllBtn.className = 'sub-nav-btn active';
      downloadAllBtn.style.background = 'rgba(255, 165, 0, 0.15)';
      downloadAllBtn.style.color = '#ffa500';
      downloadAllBtn.style.borderColor = 'rgba(255, 165, 0, 0.45)';
    }
  }

  updateNotifications();

  if (!skipTableRefresh && isPersonnelPage()) {
    refreshCurrentPageTable();
  }
}

function showToastNotification(title, text, action, extraClass = '') {
    const toast = document.createElement('div');
    toast.className = 'notif-toast ' + extraClass;
    
    // Position toasts further down if there are others
    const existingToasts = document.querySelectorAll('.notif-toast');
    let offset = 20;
    existingToasts.forEach(t => {
        offset += t.offsetHeight + 10;
    });
    toast.style.top = offset + 'px';

    const toastKey = title + text;

    toast.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color:var(--accent); flex-shrink:0;"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>
        <div style="flex-grow:1; margin-right:10px;">
            <div style="font-weight:700; font-size:0.9rem;">${title}</div>
            <div style="font-size:0.8rem; opacity:0.8;">${text}</div>
        </div>
        <button class="toast-close-btn" title="Dismiss" style="background:none; border:none; color:inherit; cursor:pointer; padding:5px; margin:-5px; opacity:0.6;">✕</button>
    `;

    const closeBtn = toast.querySelector('.toast-close-btn');
    closeBtn.onclick = (e) => {
        e.stopPropagation();
        const bundle = loadBundle();
        if (!bundle.dismissedNotifications.includes(toastKey)) {
            bundle.dismissedNotifications.push(toastKey);
            saveBundle(bundle);
        }
        toast.remove();
        repositionToasts();
    };

    toast.onclick = () => {
        action();
        toast.remove();
        // Reposition remaining toasts
        repositionToasts();
    };
    document.body.appendChild(toast);
    setTimeout(() => {
        if (toast.parentElement) {
            toast.style.animation = 'notifSlideIn 0.4s reverse forwards';
            setTimeout(() => {
                toast.remove();
                repositionToasts();
            }, 400);
        }
    }, 5000);
}

function repositionToasts() {
    const toasts = document.querySelectorAll('.notif-toast');
    let offset = 20;
    toasts.forEach(t => {
        t.style.top = offset + 'px';
        offset += t.offsetHeight + 10;
    });
}

function promptMemberOffScene(member) {
  showTimePrompt('Mark Off Scene', (date, time) => {
    const bundle = loadBundle();
    const data = bundle.pages.page3 || [];
    const memberName = member[0];
    const originalRow = data.find(row => row[0] === memberName);
    
    if (originalRow) {
      originalRow[6] = 'false'; // On Scene = false
      originalRow[1] = '';      // Clear Team
      originalRow[2] = '';      // Clear Team Lead
      
      saveBundle(bundle);
      addActivityLogEntry('Personnel', `${memberName} is now Off Scene at ${date} ${time}`, null, memberName);
      refreshCurrentPageTable();
    }
  });
}

function updateNotifications() {
  const list = document.getElementById('notif-list');
  const dot = document.getElementById('notif-dot');
  if (!list) return;
  
  list.innerHTML = '';
  let count = 0;
  const bundle = loadBundle();

  if (!window._shownToasts) window._shownToasts = new Set();

  const add = (title, text, action, extraClass = '') => {
    count++;
    const item = document.createElement('div');
    item.className = 'notification-item ' + extraClass;
    item.innerHTML = `
      <div style="display:flex; gap:10px; align-items:flex-start; width:100%;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>
        <div style="flex-grow:1;">
          <div class="notification-item-title">${title}</div>
          <div class="notification-item-text">${text}</div>
        </div>
      </div>
    `;
    item.onclick = () => {
      action();
      document.getElementById('notif-sidebar').classList.remove('open');
    };
    list.appendChild(item);

    const toastKey = title + text;
    if (!window._shownToasts.has(toastKey) && !bundle.dismissedNotifications.includes(toastKey)) {
        window._shownToasts.add(toastKey);
        showToastNotification(title, text, action, extraClass);
    }
  };

  const sweeps = getLogSweepsDue();
  sweeps.forEach(d => {
    add('Log Sweeps', `Task ${d.taskNum} (${d.region}/${d.segment}) needs sweep count.`, () => {
      navigateToPage('page4.html');
    }, 'log-sweeps-due');
  });

  const teamsMap = new Map();
  const data = bundle.pages.page3 || [];
  const baseTeamNames = ['Base Support', 'Off Duty', 'Command'];
  data.forEach(row => {
    if (row[1] && row[6] === 'true' && !baseTeamNames.includes(row[1])) {
       if (!teamsMap.has(row[1])) teamsMap.set(row[1], true);
    }
  });
  teamsMap.forEach((_, teamName) => {
    if (isParCheckDue(teamName, bundle)) {
      add('Par Check Due', `Team ${teamName} is due for a par check.`, () => {
        navigateToPage('page3.html');
      }, 'par-check-due');
    }
  });

  const searchLog = bundle.pages.page4 || [];
  const forms = bundle.forms || {};
  searchLog.forEach(row => {
    if (row[0] && row[0].startsWith('#')) {
      const num = row[0].substring(1);
      if (!isTaskUnfinished(row[0]) && (!forms[num] || !forms[num].completed)) {
        add('Fill Form', `Task #${num} (${row[3]}/${row[4]}) form needs completion.`, () => {
          navigateToPage(`page5.html?task=${num}`);
        });
      }
    }
  });

  if (!bundle.profile || !bundle.profile.completed) {
    add('Fill Incident Profile', 'Incident profile is not marked as completed.', () => {
      navigateToPage('page6.html');
    }, 'profile-incomplete');
  }

  if (dot) {
    if (count > 0) {
      dot.classList.add('active');
    } else {
      dot.classList.remove('active');
    }
  }
}

const PAGE_ORDER = [
  'home.html',
  'index.html',
  'page2.html',
  'page3.html',
  'page4.html',
  'page5.html',
  'page6.html',
  'page7.html',
  'page8.html',
  'page9.html',
  'settings.html'
];

function navigateToPage(targetUrl) {
    window.location.href = targetUrl;
}

function initPageTransitions() {
    document.addEventListener('click', e => {
        const a = e.target.closest('a');
        if (a && a.href && a.href.includes(window.location.origin) && !a.hasAttribute('download') && a.target !== '_blank') {
            const targetUrl = a.getAttribute('href');
            if (targetUrl !== '#' && !targetUrl.startsWith('javascript:')) {
                e.preventDefault();
                navigateToPage(targetUrl);
            }
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    initPageTransitions();
    const bundle = loadBundle();
    applyTheme(bundle);
    applyBackground(bundle);
    applyTipsVisibility(bundle);
    updateFileNameDisplay();
    updateHeaderProfile();

    const currentUser = getCurrentUser();
    if (!currentUser) {
        showLoginPopup();
    } else {
        checkAccess();
    }

  const bell = document.getElementById('notif-bell');
  const sidebar = document.getElementById('notif-sidebar');
  const closeNotif = document.getElementById('close-notif');
  if (bell && sidebar) {
    bell.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg><div class="notification-dot" id="notif-dot"></div>';
    bell.onclick = () => sidebar.classList.toggle('open');
  }
  if (closeNotif && sidebar) {
    closeNotif.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';
    closeNotif.onclick = () => sidebar.classList.remove('open');
  }

  if (typeof Notification !== 'undefined' && Notification.permission === "default") {
    Notification.requestPermission();
  }

  // Remove Gummy Bear from permanent personnel if present
  const globalPersonnel = getPermanentPersonnel();
  if (globalPersonnel['Gummy Bear']) {
    delete globalPersonnel['Gummy Bear'];
    setPermanentPersonnel(globalPersonnel);
    // If we are on page3, we might need to refresh data to remove it from current view
    if (isPersonnelPage()) {
        const bundle = loadBundle();
        if (bundle.pages.page3) {
            bundle.pages.page3 = Array.isArray(bundle.pages.page3) ? bundle.pages.page3.filter(row => row[0] !== 'Gummy Bear') : bundle.pages.page3;
            saveBundle(bundle);
            buildPersonnelTable();
        }
    }
  }
  
  checkParChecksAndNotify();
  setInterval(checkParChecksAndNotify, 60000);

  // Auto-save every 5 minutes to the file list on the home page
  setInterval(() => {
    const bundle = loadBundle();
    if (bundle.fileName) {
      saveFileToList(bundle.fileName, bundle);
      saveBundle(bundle);
      const statusEl = document.getElementById('save-status') || document.getElementById('home-status');
      if (statusEl) {
        const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        statusEl.textContent = `Auto-saved at ${timestamp}`;
      }
      if (isHomePage()) {
        buildSavedFilesTable();
      }
    }
  }, 300000);

  if (isHomePage()) {
    buildHomePage();
    return;
  }

  if (isSettingsPage()) {
    buildSettingsPage();
    return;
  }

  if (isRegionsPage()) {
    buildRegionsTable();
  } else if (isSegmentsPage()) {
    buildSegmentsTable();
  } else if (isPersonnelPage()) {
    buildPersonnelTable();
    initBaseTeamsAccordion();
    initSearchTeamsAccordion();
  } else if (isSearchLogPage()) {
    buildSearchLogTable();
  } else if (isFormsPage()) {
    buildFormsPage();
  } else if (isProfilePage()) {
    buildProfilePage();
  } else if (isPage8()) {
    buildUserAccountPage();
  } else if (isPage9()) {
    buildUserManagementPage();
  } else if (isMapsPage()) {
    buildMapsPage();
  } else if (isUploadsPage()) {
    buildUploadsPage();
  } else {
    buildStandardTable();
  }
});

function initBaseTeamsAccordion() {
  const accordionHeader = document.getElementById('base-teams-accordion-header');
  const accordionContainer = document.getElementById('base-teams-container-header');
  if (!accordionHeader || !accordionContainer) return;

  // Restore state from localStorage
  const isCollapsed = localStorage.getItem('baseTeamsCollapsed') === 'true';
  if (isCollapsed) {
    accordionContainer.classList.add('collapsed');
  }

  accordionHeader.onclick = () => {
    accordionContainer.classList.toggle('collapsed');
    localStorage.setItem('baseTeamsCollapsed', accordionContainer.classList.contains('collapsed'));
  };
}

function initSearchTeamsAccordion() {
  const accordionHeader = document.getElementById('search-teams-accordion-header');
  const accordionContainer = document.getElementById('search-teams-container');
  if (!accordionHeader || !accordionContainer) return;

  // Restore state from localStorage
  const isCollapsed = localStorage.getItem('searchTeamsCollapsed') === 'true';
  if (isCollapsed) {
    accordionContainer.classList.add('collapsed');
  }

  accordionHeader.onclick = () => {
    accordionContainer.classList.toggle('collapsed');
    localStorage.setItem('searchTeamsCollapsed', accordionContainer.classList.contains('collapsed'));
  };
}

function buildUploadsPage() {
  const tableBody = document.getElementById('uploads-table-body');
  const uploadInput = document.getElementById('file-upload-input');
  const saveStatus = document.getElementById('save-status');
  if (!tableBody || !uploadInput) return;

  let bundle = loadBundle();
  if (!bundle.uploads) bundle.uploads = [];

  const render = () => {
    tableBody.innerHTML = '';
    bundle.uploads.forEach((file, index) => {
      const tr = document.createElement('tr');
      animateNewRow(tr, index);
      
      const tdName = document.createElement('td');
      tdName.dataset.label = 'File Name';
      const namePill = document.createElement('div');
      namePill.className = 'pill-cell readonly-pill';
      namePill.textContent = file.name;
      tdName.appendChild(namePill);
      tr.appendChild(tdName);
      
      const tdType = document.createElement('td');
      tdType.dataset.label = 'Type';
      const typePill = document.createElement('div');
      typePill.className = 'pill-cell readonly-pill';
      typePill.textContent = file.type || 'unknown';
      tdType.appendChild(typePill);
      tr.appendChild(tdType);
      
      const tdSize = document.createElement('td');
      tdSize.dataset.label = 'Size';
      const sizePill = document.createElement('div');
      sizePill.className = 'pill-cell readonly-pill';
      sizePill.textContent = (file.size / 1024).toFixed(2) + ' KB';
      tdSize.appendChild(sizePill);
      tr.appendChild(tdSize);
      
      const tdActions = document.createElement('td');
      tdActions.dataset.label = 'Actions';
      const btnContainer = document.createElement('div');
      btnContainer.className = 'tool-actions';
      btnContainer.style.justifyContent = 'center';
      
      const downBtn = document.createElement('button');
      downBtn.className = 'row-delete-btn update-pill';
      downBtn.textContent = 'Download';
      downBtn.onclick = () => {
        const link = document.createElement('a');
        link.href = file.content;
        link.download = file.name;
        link.click();
      };
      
      const delBtn = document.createElement('button');
      delBtn.className = 'row-delete-btn update-pill';
      delBtn.style.color = 'var(--accent)';
      delBtn.textContent = 'Delete';
      delBtn.onclick = () => {
        const fileName = bundle.uploads[index].name;
        bundle.uploads.splice(index, 1);
        logDeletion('Upload', fileName);
        saveBundle(bundle);
        render();
      };
      
      btnContainer.appendChild(downBtn);
      btnContainer.appendChild(delBtn);
      tdActions.appendChild(btnContainer);
      tr.appendChild(tdActions);
      
      tableBody.appendChild(tr);
    });
  };

  uploadInput.onchange = (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    
    if (saveStatus) saveStatus.textContent = 'Uploading...';
    
    let processedCount = 0;
    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = (re) => {
        bundle.uploads.push({
          name: file.name,
          type: file.type,
          size: file.size,
          content: re.target.result
        });
        logCreation('Upload', file.name, bundle);
        processedCount++;
        if (processedCount === files.length) {
          saveBundle(bundle);
          if (saveStatus) saveStatus.textContent = 'Saved.';
          render();
        }
      };
      reader.onerror = () => {
         processedCount++;
         if (saveStatus) saveStatus.textContent = 'Some uploads failed.';
         if (processedCount === files.length) {
            saveBundle(bundle);
            render();
         }
      };
      reader.readAsDataURL(file);
    });
  };

  render();
}

const R_MILES = 3958.8;

const haversine = (p1, p2) => {
    const lon1 = p1[0], lat1 = p1[1];
    const lon2 = p2[0], lat2 = p2[1];
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon/2) * Math.sin(dLon/2);
    return 2 * R_MILES * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
};

const polygonArea = (rings) => {
    let totalArea = 0;
    for (const ring of rings) {
        if (ring.length < 3) continue;
        let area = 0;
        const latRef = ring[0][1];
        const k = Math.cos(latRef * Math.PI / 180);
        for (let i = 0; i < ring.length - 1; i++) {
            const p1 = ring[i];
            const p2 = ring[i+1];
            const x1 = p1[0] * k * 69.172; // miles per degree lon
            const y1 = p1[1] * 69.172; // miles per degree lat
            const x2 = p2[0] * k * 69.172;
            const y2 = p2[1] * 69.172;
            area += (x1 * y2 - x2 * y1);
        }
        totalArea += Math.abs(area) / 2;
    }
    return totalArea * 640; // in acres
};

const getBoundingBoxLength = (coords) => {
    let minLon = 180, maxLon = -180, minLat = 90, maxLat = -90;
    const process = (c) => {
        if (typeof c[0] === 'number') {
            if (c[0] < minLon) minLon = c[0];
            if (c[0] > maxLon) maxLon = c[0];
            if (c[1] < minLat) minLat = c[1];
            if (c[1] > maxLat) maxLat = c[1];
        } else {
            c.forEach(process);
        }
    };
    process(coords);
    if (minLon > maxLon) return 0;
    const midLat = (minLat + maxLat) / 2;
    const width = haversine([minLon, midLat], [maxLon, midLat]);
    const height = haversine([0, minLat], [0, maxLat]);
    return Math.max(width, height);
};

const calculateGeometry = (item) => {
    const geom = item.geometry || item; // Item might be the geometry itself
    if (!geom || !geom.type || !geom.coordinates) return { area: 0, length: 0 };

    let area = 0;
    let length = 0;

    if (geom.type === 'LineString') {
        for (let i = 0; i < geom.coordinates.length - 1; i++) {
            length += haversine(geom.coordinates[i], geom.coordinates[i+1]);
        }
    } else if (geom.type === 'MultiLineString') {
        for (const line of geom.coordinates) {
            for (let i = 0; i < line.length - 1; i++) {
                length += haversine(line[i], line[i+1]);
            }
        }
    } else if (geom.type === 'Polygon') {
        area = polygonArea(geom.coordinates);
        length = getBoundingBoxLength(geom.coordinates);
    } else if (geom.type === 'MultiPolygon') {
        for (const poly of geom.coordinates) {
            area += polygonArea(poly);
        }
        length = getBoundingBoxLength(geom.coordinates);
    } else if (geom.type === 'GeometryCollection') {
        for (const g of (geom.geometries || [])) {
            const res = calculateGeometry({ geometry: g });
            area += res.area;
            length = Math.max(length, res.length);
        }
    }

    return { area, length };
};

function showImportSegmentsPopup() {
  const bundle = loadBundle();
  const uploads = bundle.uploads || [];
  const jsonFiles = uploads.filter(f => f.name.toLowerCase().endsWith('.json'));

  if (jsonFiles.length === 0) {
    alert("No JSON files found in uploads. Please upload JSON files on the Uploads page first.");
    return;
  }

  const popup = createPopup('Import Segments from JSON', null);
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');

  content.style.display = 'flex';
  content.style.flexDirection = 'column';
  content.style.maxHeight = '90vh';

  const bodyContainer = document.createElement('div');
  bodyContainer.style.flex = '1';
  bodyContainer.style.display = 'flex';
  bodyContainer.style.flexDirection = 'column';
  bodyContainer.style.overflow = 'hidden';
  content.insertBefore(bodyContainer, btnContainer);

  function renderFileSelection() {
    bodyContainer.innerHTML = '<p style="margin-bottom: 15px; opacity: 0.8; flex-shrink: 0;">Select one or more JSON files to import as segments:</p>';
    const list = document.createElement('div');
    list.style.display = 'flex';
    list.style.flexDirection = 'column';
    list.style.gap = '10px';
    list.style.flex = '1';
    list.style.overflowY = 'auto';
    list.style.paddingRight = '5px';

    jsonFiles.forEach((file, index) => {
      const label = document.createElement('label');
      label.className = 'pill-cell-container clickable-pill';
      label.style.display = 'flex';
      label.style.alignItems = 'center';
      label.style.gap = '15px';
      label.style.padding = '12px 20px';
      label.style.background = 'rgba(255,255,255,0.03)';
      label.style.border = '1px solid var(--pill-border)';
      label.style.borderRadius = '999px';
      label.style.cursor = 'pointer';

      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.className = 'pill-checkbox';
      chk.dataset.index = index;
      
      const info = document.createElement('div');
      info.style.flex = '1';
      info.innerHTML = `<div style="font-weight:700;">${file.name}</div><div style="font-size:0.8rem; opacity:0.6;">${(file.size / 1024).toFixed(2)} KB</div>`;
      
      label.appendChild(chk);
      label.appendChild(info);
      list.appendChild(label);
    });

    bodyContainer.appendChild(list);

    btnContainer.innerHTML = '';
    btnContainer.style.display = 'flex';
    btnContainer.style.flexDirection = 'row';
    btnContainer.style.justifyContent = 'flex-end';
    btnContainer.style.marginTop = '20px';
    btnContainer.style.flexShrink = '0';
    btnContainer.style.gap = '12px';

    const cancelBtn = document.createElement('button');
    cancelBtn.className = 'popup-btn';
    cancelBtn.style.flex = '1';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.onclick = () => closePopup(popup);
    btnContainer.appendChild(cancelBtn);

    const nextBtn = document.createElement('button');
    nextBtn.className = 'popup-btn primary';
    nextBtn.style.flex = '1';
    nextBtn.textContent = 'Preview Import';
    nextBtn.onclick = () => {
      const selectedIndices = Array.from(bodyContainer.querySelectorAll('.pill-checkbox:checked')).map(cb => parseInt(cb.dataset.index));
      if (selectedIndices.length === 0) {
        alert("Please select at least one file.");
        return;
      }
      processFiles(selectedIndices.map(i => jsonFiles[i]));
    };
    btnContainer.appendChild(nextBtn);
  }

  function processFiles(files) {
    const segmentsToImport = [];

    const findProp = (item, keys) => {
      const p = item.properties || item.attributes || item.fields || {};
      for (const k of keys) {
        const found = Object.keys(p).find(x => x.toLowerCase() === k.toLowerCase());
        if (found && p[found] !== undefined && p[found] !== null && p[found] !== '') return { value: p[found], key: found };
        const foundTop = Object.keys(item).find(x => x.toLowerCase() === k.toLowerCase());
        if (foundTop && item[foundTop] !== undefined && item[foundTop] !== null && item[foundTop] !== '') return { value: item[foundTop], key: foundTop };
      }
      return null;
    };

    const calculateGeometryLocal = (item) => {
        return calculateGeometry(item);
    };

    const areaKeys = ['acres', 'area', 'size', 'sqmi', 'sq_mi', 'shape_area', 'st_area', 'hectares', 'ha', 'gis_acres', 'acres_total', 'sqft', 'sq_ft', 'sqkm', 'sq_km', 'total_acres'];
    const lengthKeys = ['length', 'len', 'leng', 'shape_leng', 'distance', 'mi', 'miles', 'shape_length', 'st_length', 'dist', 'width', 'height', 'ft', 'feet', 'km', 'meters', 'm', 'shape_len'];
    const nameKeys = ['name', 'segment', 'label', 'id', 'title', 'id_number', 'unit_id', 'objectid', 'fid'];

    files.forEach(file => {
      let jsonText = '';
      try {
        if (file.content.startsWith('data:')) {
          const base64 = file.content.split(',')[1];
          const binString = atob(base64);
          const bytes = Uint8Array.from(binString, c => c.charCodeAt(0));
          jsonText = new TextDecoder().decode(bytes);
        } else {
          jsonText = file.content;
        }
      } catch (e) {
        console.error("Failed to decode file", file.name, e);
        return;
      }

      let data;
      try {
        data = JSON.parse(jsonText);
      } catch (e) {
        console.error("Failed to parse JSON", file.name, e);
        return;
      }

      let items = [];
      if (Array.isArray(data)) {
        items = data;
      } else if (data.features && Array.isArray(data.features)) {
        items = data.features;
      } else {
        items = [data];
      }

      items.forEach((item, idx) => {
        let name = file.name.replace(/\.json$/i, '');
        if (items.length > 1) {
            const innerNameRes = findProp(item, ['name', 'segment', 'label', 'title']);
            if (innerNameRes) {
                name = String(innerNameRes.value);
            } else {
                name += ` - ${idx + 1}`;
            }
        }
        
        const areaRes = findProp(item, areaKeys);
        const lengthRes = findProp(item, lengthKeys);
        
        let areaDefaultUnit = 'ac';
        if (areaRes) {
          const lk = areaRes.key.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (lk.includes('sqmi')) areaDefaultUnit = 'sqmi';
          else if (lk.includes('ha') || lk.includes('hectare')) areaDefaultUnit = 'ha';
          else if (lk.includes('sqkm')) areaDefaultUnit = 'sqkm';
          else if (lk.includes('sqm') || lk.includes('starea') || lk.includes('shapearea')) areaDefaultUnit = 'sqm';
          else if (lk.includes('sqft')) areaDefaultUnit = 'sqft';
        }

        let lengthDefaultUnit = 'mi';
        if (lengthRes) {
          const lk = lengthRes.key.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (lk.includes('km')) lengthDefaultUnit = 'km';
          else if (lk.includes('ft') || lk.includes('feet')) lengthDefaultUnit = 'ft';
          else if (lk.includes('mi') && !lk.includes('sqmi')) lengthDefaultUnit = 'mi';
          else if (lk.includes('m') && !lk.includes('mi')) lengthDefaultUnit = 'm';
          else if (lk.includes('stlength') || lk.includes('shapelength') || lk.includes('shapeleng') || lk.includes('perimeter')) lengthDefaultUnit = 'm';
        }
        
        let area = areaRes ? parseWithUnits(String(areaRes.value), areaDefaultUnit) : 0;
        let lengthVal = lengthRes ? parseWithUnits(String(lengthRes.value), lengthDefaultUnit) : 0;
        
        // If still 0, try geometry calculation
        if (area === 0 || lengthVal === 0) {
            const geomRes = calculateGeometryLocal(item);
            if (area === 0 && geomRes.area > 0) area = geomRes.area;
            if (lengthVal === 0 && geomRes.length > 0) lengthVal = geomRes.length;
        }

        // Fallback for length if not found but width/height exist
        if (lengthVal === 0) {
          const w = findProp(item, ['width']);
          const h = findProp(item, ['height']);
          if (w || h) {
             const wv = w ? parseWithUnits(String(w.value), 'mi') : 0;
             const hv = h ? parseWithUnits(String(h.value), 'mi') : 0;
             lengthVal = Math.max(wv, hv);
          }
        }
        
        segmentsToImport.push({
          region: '',
          segment: name,
          area: area > 0 ? area.toFixed(2) : '',
          length: lengthVal > 0 ? lengthVal.toFixed(2) : '',
          sweep: 20
        });
      });
    });

    renderPreviewTable(segmentsToImport);
  }

  function parseWithUnits(valStr, defaultUnit) {
    if (!valStr) return 0;
    // Clean string for parsing numeric value: handle commas as thousand separators
    const cleanStr = String(valStr).replace(/,/g, '');
    const val = parseFloat(cleanStr);
    if (isNaN(val)) return 0;

    const lower = String(valStr).toLowerCase().replace(/[^a-z0-9]/g, '');
    
    // String contains units - they override defaultUnit
    if (lower.includes('sqkm') || lower.includes('squarekilometer')) return val * 247.105;
    if (lower.includes('sqm') || lower.includes('m2') || lower.includes('squaremeter')) return val * 0.000247105;
    if (lower.includes('ha') || lower.includes('hectare')) return val * 2.47105;
    if (lower.includes('sqft') || lower.includes('ft2') || lower.includes('squarefeet')) return val / 43560;
    if (lower.includes('sqmi') || lower.includes('squaremiles')) return val * 640;
    if (lower.includes('acres') || (lower.includes('ac') && !lower.includes('active'))) {
        if (!lower.includes('sqm') && !lower.includes('ha') && !lower.includes('sqft') && !lower.includes('sqmi')) {
            return val;
        }
    }

    if (lower.includes('km') || lower.includes('kilometers')) return val * 0.621371;
    if (lower.includes('ft') || lower.includes('feet')) return val / 5280;
    if (lower.includes('meters') || (lower.includes('m') && !lower.includes('mi') && !lower.includes('km'))) return val * 0.000621371;

    // No units in string, use defaultUnit
    if (defaultUnit === 'sqkm') return val * 247.105;
    if (defaultUnit === 'sqm') return val * 0.000247105;
    if (defaultUnit === 'ha') return val * 2.47105;
    if (defaultUnit === 'sqft') return val / 43560;
    if (defaultUnit === 'sqmi') return val * 640;
    if (defaultUnit === 'km') return val * 0.621371;
    if (defaultUnit === 'm') return val * 0.000621371;
    if (defaultUnit === 'ft') return val / 5280;
    
    return val;
  }

  function renderPreviewTable(segments) {
    bodyContainer.innerHTML = `
        <p style="margin-bottom: 15px; opacity: 0.8; flex-shrink: 0;">Review segments to be imported:</p>
        <div style="margin-bottom: 10px; display: flex; align-items: center; gap: 10px;">
            <input type="checkbox" id="check-all-preview" checked style="width: 18px; height: 18px; cursor: pointer;">
            <label for="check-all-preview" style="cursor: pointer; font-weight: bold;">Check / Uncheck All</label>
        </div>
    `;
    content.style.width = '90vw';
    content.style.maxWidth = '1100px';

    const tableWrap = document.createElement('div');
    tableWrap.style.overflowX = 'auto';
    tableWrap.style.flex = '1';
    tableWrap.style.overflowY = 'auto';
    tableWrap.style.background = 'rgba(0,0,0,0.1)';
    tableWrap.style.borderRadius = '12px';
    
    const table = document.createElement('table');
    table.className = 'grid-table';
    table.style.width = '100%';
    
    const thead = document.createElement('thead');
    const headers = ['', 'Region', 'Segment', 'Area (acres)', 'Length (mi)', 'Sweep (ft)', 'Time per Sweep (hr)', 'PSRi', 'PSRc'];
    const htr = document.createElement('tr');
    headers.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      th.style.padding = '12px';
      htr.appendChild(th);
    });
    thead.appendChild(htr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    segments.forEach((seg, idx) => {
      const tr = document.createElement('tr');
      const lengthVal = parseFloat(seg.length) || 0;
      const timeVal = lengthVal / 0.5;
      
      const tdCheck = document.createElement('td');
      tdCheck.style.textAlign = 'center';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.className = 'preview-checkbox';
      chk.dataset.index = idx;
      chk.checked = true;
      chk.style.width = '18px';
      chk.style.height = '18px';
      chk.style.cursor = 'pointer';
      tdCheck.appendChild(chk);
      tr.appendChild(tdCheck);

      const rowData = [
        '',
        seg.segment,
        seg.area ? seg.area + ' ac' : '',
        seg.length ? seg.length + ' mi' : '',
        seg.sweep + ' ft',
        timeVal > 0 ? timeVal.toFixed(2) + ' hr' : '',
        '',
        ''
      ];

      rowData.forEach((val) => {
        const td = document.createElement('td');
        const pill = document.createElement('div');
        pill.className = 'pill-cell readonly-pill';
        pill.style.padding = '8px 12px';
        pill.textContent = val;
        td.appendChild(pill);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tableWrap.appendChild(table);
    bodyContainer.appendChild(tableWrap);

    const checkAll = bodyContainer.querySelector('#check-all-preview');
    const checkboxes = bodyContainer.querySelectorAll('.preview-checkbox');
    checkAll.onchange = () => {
        checkboxes.forEach(cb => cb.checked = checkAll.checked);
    };

    btnContainer.innerHTML = '';
    btnContainer.style.display = 'flex';
    btnContainer.style.flexDirection = 'row';
    btnContainer.style.justifyContent = 'flex-end';
    btnContainer.style.marginTop = '20px';
    btnContainer.style.flexShrink = '0';
    btnContainer.style.gap = '12px';

    const backBtn = document.createElement('button');
    backBtn.className = 'popup-btn';
    backBtn.style.flex = '1';
    backBtn.textContent = 'Back';
    backBtn.onclick = () => renderFileSelection();
    btnContainer.appendChild(backBtn);

    const submitBtn = document.createElement('button');
    submitBtn.className = 'popup-btn primary';
    submitBtn.style.flex = '1';
    submitBtn.textContent = 'Submit Import';
    submitBtn.onclick = () => {
      const selected = Array.from(checkboxes)
          .filter(cb => cb.checked)
          .map(cb => segments[parseInt(cb.dataset.index)]);
      
      if (selected.length === 0) {
          alert("No segments selected.");
          return;
      }
      importSegmentsAction(selected);
      closePopup(popup);
    };
    btnContainer.appendChild(submitBtn);
  }

  function importSegmentsAction(segments) {
    const b = loadBundle();
    if (!b.pages.page2) b.pages.page2 = defaultSegmentsData();
    
    // Ensure it's the correct format (headers + rows)
    if (Array.isArray(b.pages.page2)) {
        const rows = b.pages.page2;
        b.pages.page2 = defaultSegmentsData();
        b.pages.page2.rows = rows;
    }

    const importedNames = [];
    segments.forEach(seg => {
      const lengthVal = parseFloat(seg.length) || 0;
      const timeVal = lengthVal / 0.5;
      const newRow = [
        '', // Region
        seg.segment,
        seg.area ? seg.area + ' ac' : '',
        seg.length ? seg.length + ' mi' : '',
        seg.sweep + ' ft',
        timeVal > 0 ? timeVal.toFixed(2) + ' hr' : '',
        '', // PSRi
        '', // PSRc
        ''  // manual override
      ];
      b.pages.page2.rows.push(newRow);
      newlyImportedSegments.add(`|${seg.segment}`);
      importedNames.push(seg.segment);
    });

    if (importedNames.length > 0) {
      addActivityLogEntry('System', 'Imported segments: ' + importedNames.join(', '), b);
    }

    saveBundle(b);
    recalculateEverything();
    buildSegmentsTable();

    setTimeout(() => {
      newlyImportedSegments.clear();
      buildSegmentsTable();
    }, 7000);
  }

  renderFileSelection();
}

function showSarTopoShapesPopup(features) {
  const popup = createPopup('Import Shapes from SarTopo', null);
  const content = popup.querySelector('.popup-content');
  const btnContainer = popup.querySelector('.popup-buttons');

  content.style.display = 'flex';
  content.style.flexDirection = 'column';
  content.style.maxHeight = '90vh';
  content.style.width = '90vw';
  content.style.maxWidth = '1100px';

  const bodyContainer = document.createElement('div');
  bodyContainer.style.flex = '1';
  bodyContainer.style.display = 'flex';
  bodyContainer.style.flexDirection = 'column';
  bodyContainer.style.overflow = 'hidden';
  content.insertBefore(bodyContainer, btnContainer);

  const segmentsToPreview = features.map(f => {
      const props = f.properties || {};
      const geom = f.geometry || {};
      
      let name = props.label || props.title || props.name || 'Unnamed';
      
      // Improved naming for Assignments
      if (props.class === 'Assignment' || props.type === 'Assignment' || props.assignment) {
          const a = (typeof props.assignment === 'object') ? props.assignment : props;
          const num = a.number || '';
          const label = a.label || props.title || props.label || '';
          if (num && label) {
              name = `${num} ${label}`;
          } else if (num) {
              name = num;
          } else if (label) {
              name = label;
          }
      }

      const res = calculateGeometry(f);
      return {
          name: name,
          area: res.area > 0 ? res.area.toFixed(2) : '',
          length: res.length > 0 ? res.length.toFixed(2) : '',
          type: geom.type,
          feature: f
      };
  });

  bodyContainer.innerHTML = `
    <p style="margin-bottom: 15px; opacity: 0.8; flex-shrink: 0;">Select the shapes you want to import as segments:</p>
    <div style="margin-bottom: 10px; display: flex; align-items: center; gap: 10px;">
        <input type="checkbox" id="check-all-shapes" checked style="width: 18px; height: 18px; cursor: pointer;">
        <label for="check-all-shapes" style="cursor: pointer; font-weight: bold;">Check / Uncheck All</label>
    </div>
  `;

  const tableWrap = document.createElement('div');
  tableWrap.style.overflowX = 'auto';
  tableWrap.style.flex = '1';
  tableWrap.style.overflowY = 'auto';
  tableWrap.style.background = 'rgba(0,0,0,0.1)';
  tableWrap.style.borderRadius = '12px';
  
  const table = document.createElement('table');
  table.className = 'grid-table';
  table.style.width = '100%';
  
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
        <th style="width: 40px; text-align: center; padding: 12px;"></th>
        <th style="padding: 12px;">Name</th>
        <th style="padding: 12px;">Type</th>
        <th style="padding: 12px;">Area (acres)</th>
        <th style="padding: 12px;">Length (mi)</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  segmentsToPreview.forEach((seg, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center;"><input type="checkbox" class="shape-checkbox" data-index="${idx}" checked style="width: 18px; height: 18px; cursor: pointer;"></td>
        <td><div class="pill-cell readonly-pill" style="padding: 8px 12px;">${seg.name}</div></td>
        <td><div class="pill-cell readonly-pill" style="padding: 8px 12px;">${seg.type}</div></td>
        <td><div class="pill-cell readonly-pill" style="padding: 8px 12px;">${seg.area ? seg.area + ' ac' : ''}</div></td>
        <td><div class="pill-cell readonly-pill" style="padding: 8px 12px;">${seg.length ? seg.length + ' mi' : ''}</div></td>
      `;
      tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  bodyContainer.appendChild(tableWrap);

  const checkAll = bodyContainer.querySelector('#check-all-shapes');
  const checkboxes = bodyContainer.querySelectorAll('.shape-checkbox');
  checkAll.onchange = () => {
      checkboxes.forEach(cb => cb.checked = checkAll.checked);
  };

  btnContainer.innerHTML = '';
  btnContainer.style.display = 'flex';
  btnContainer.style.gap = '12px';
  btnContainer.style.marginTop = '20px';

  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'popup-btn';
  cancelBtn.style.flex = '1';
  cancelBtn.textContent = 'Cancel';
  cancelBtn.onclick = () => closePopup(popup);
  btnContainer.appendChild(cancelBtn);

  const submitBtn = document.createElement('button');
  submitBtn.className = 'popup-btn primary';
  submitBtn.style.flex = '1';
  submitBtn.textContent = 'Submit Import';
  submitBtn.onclick = () => {
      const selected = Array.from(checkboxes)
          .filter(cb => cb.checked)
          .map(cb => segmentsToPreview[parseInt(cb.dataset.index)]);
      
      if (selected.length === 0) {
          alert("No shapes selected.");
          return;
      }
      
      importSarTopoSegments(selected);
      closePopup(popup);
  };
  btnContainer.appendChild(submitBtn);
}

function importSarTopoSegments(selected) {
    const b = loadBundle();
    if (!b.pages.page2) b.pages.page2 = defaultSegmentsData();
    
    // Ensure it's the correct format (headers + rows)
    if (Array.isArray(b.pages.page2)) {
        const rows = b.pages.page2;
        b.pages.page2 = defaultSegmentsData();
        b.pages.page2.rows = rows;
    }

    const importedNames = [];
    selected.forEach(seg => {
        const lengthVal = parseFloat(seg.length) || 0;
        const timeVal = lengthVal / 0.5;
        const newRow = [
            '', // Region
            seg.name,
            seg.area ? seg.area + ' ac' : '',
            seg.length ? seg.length + ' mi' : '',
            '20 ft', // Default sweep
            timeVal > 0 ? timeVal.toFixed(2) + ' hr' : '',
            '', // PSRi
            '', // PSRc
            ''  // manual override
        ];
        b.pages.page2.rows.push(newRow);
        newlyImportedSegments.add(`|${seg.name}`);
        importedNames.push(seg.name);
    });

    if (importedNames.length > 0) {
        addActivityLogEntry('System', 'Imported SarTopo shapes as segments: ' + importedNames.join(', '), b);
    }

    saveBundle(b);
    recalculateEverything();
    if (isSegmentsPage()) buildSegmentsTable();

    setTimeout(() => {
        newlyImportedSegments.clear();
        if (isSegmentsPage()) buildSegmentsTable();
    }, 7000);
}

function buildUserAccountPage() {
    const container = document.getElementById('user-account-container');
    const tabContainer = document.getElementById('user-tabs');
    if (!container) return;

    const bundle = loadBundle();
    const urlParams = new URLSearchParams(window.location.search);
    const userPin = urlParams.get('userPin');
    const tab = urlParams.get('tab') || 'account';
    const currentUser = getCurrentUser();

    if (tabContainer) {
        if (isUserAdmin(currentUser)) {
            tabContainer.style.display = 'flex';
            const tabAccount = document.getElementById('tab-account');
            const tabManage = document.getElementById('tab-manage');
            
            tabAccount.classList.toggle('active', tab === 'account');
            tabManage.classList.toggle('active', tab === 'manage');

            tabAccount.onclick = () => {
                const newUrl = window.location.pathname + '?tab=account';
                window.history.replaceState(null, '', newUrl);
                buildUserAccountPage();
            };
            tabManage.onclick = () => {
                const newUrl = window.location.pathname + '?tab=manage';
                window.history.replaceState(null, '', newUrl);
                buildUserAccountPage();
            };
        } else {
            tabContainer.style.display = 'none';
        }
    }

    if (tab === 'manage' && isUserAdmin(currentUser)) {
        renderUserManagement(container, bundle);
        return;
    }
    
    let userToEdit;
    if (userPin && isUserAdmin(currentUser)) {
        userToEdit = (bundle.accounts || []).find(a => a.pin === userPin);
    } else {
        userToEdit = (bundle.accounts || []).find(a => a.pin === (currentUser ? currentUser.pin : ''));
    }

    if (!userToEdit) {
        container.innerHTML = `
            <div style="text-align: center; padding: 40px; color: white;">
                <p style="font-size: 1.2rem; margin-bottom: 20px;">User not found.</p>
                <button id="switch-user-btn" class="mini-pill" style="padding: 12px 20px; font-size: 1rem; background: rgba(235, 87, 87, 0.1); border-color: rgba(235, 87, 87, 0.4);">Switch User</button>
            </div>
        `;
        const switchBtn = document.getElementById('switch-user-btn');
        if (switchBtn) {
            switchBtn.onclick = () => {
                sessionStorage.removeItem('sar-current-user');
                window.location.href = 'home.html';
            };
        }
        return;
    }

    container.innerHTML = `
        <div class="profile-form" style="max-width: 800px; margin: 0 auto; background: rgba(0,0,0,0.2); padding: 30px; border-radius: 32px; border: 1px solid rgba(255,255,255,0.1);">
            <div class="form-group small">
                <label style="display: block; margin-bottom: 8px; color: var(--text); font-weight: bold;">First Name</label>
                <input type="text" id="user-first-name" class="pill-input" value="${userToEdit.firstName || ''}" style="width: 100%; box-sizing: border-box;">
            </div>
            <div class="form-group small">
                <label style="display: block; margin-bottom: 8px; color: var(--text); font-weight: bold;">Last Name</label>
                <input type="text" id="user-last-name" class="pill-input" value="${userToEdit.lastName || ''}" style="width: 100%; box-sizing: border-box;">
            </div>
            <div class="form-group small">
                <label style="display: block; margin-bottom: 8px; color: var(--text); font-weight: bold;">User PIN</label>
                <input type="text" id="user-pin" class="pill-input" value="${userToEdit.pin || ''}" style="width: 100%; box-sizing: border-box;">
            </div>
            <div class="form-group small">
                <label style="display: block; margin-bottom: 8px; color: var(--text); font-weight: bold;">Handle</label>
                <input type="text" id="user-handle" class="pill-input" value="${userToEdit.handle || ''}" style="width: 100%; box-sizing: border-box;">
            </div>
            <div class="form-group large">
                <label style="display: block; margin-bottom: 8px; color: var(--text); font-weight: bold;">Theme Preference</label>
                <div style="display: flex; gap: 10px;">
                    <button id="theme-dark-btn" class="mini-pill ${userToEdit.theme !== 'light' ? 'active' : ''}" style="flex: 1;">Dark Mode</button>
                    <button id="theme-light-btn" class="mini-pill ${userToEdit.theme === 'light' ? 'active' : ''}" style="flex: 1;">Grey Mode</button>
                </div>
            </div>
            <div class="form-group large">
                <label style="display: block; margin-bottom: 12px; color: var(--text); font-weight: bold;">Highlight Color</label>
                <div id="color-selection" style="display: flex; flex-wrap: wrap; gap: 12px; justify-content: center;">
                    <!-- Color buttons will be here -->
                </div>
            </div>
            <div class="tool-actions" style="margin-top: 20px; justify-content: center; grid-column: span 6;">
                <button id="save-user-btn" class="update-pill" style="padding: 12px 40px; font-size: 1rem;">Update Account</button>
                <button id="switch-user-btn" class="mini-pill" style="padding: 12px 20px; font-size: 1rem; margin-left: 10px; background: rgba(235, 87, 87, 0.1); border-color: rgba(235, 87, 87, 0.4);">Switch User</button>
            </div>
        </div>
    `;


    const colors = ['none', 'orange', 'yellow', 'red', 'blue', 'green', 'purple', 'brown', 'black', 'white', 'grey', 'maroon'];
    const colorContainer = document.getElementById('color-selection');
    colors.forEach(c => {
        const btn = document.createElement('button');
        btn.className = 'pill-cell-btn';
        btn.style.width = '36px';
        btn.style.height = '36px';
        btn.style.borderRadius = '50%';
        btn.style.padding = '0';
        btn.style.cursor = 'pointer';
        btn.style.transition = 'transform 0.2s ease';
        
        const updateUI = () => {
            // Unselect all
            Array.from(colorContainer.children).forEach(b => {
                b.style.border = '1px solid rgba(255,255,255,0.3)';
                b.style.transform = 'scale(1)';
            });
            // Select this one
            btn.style.border = '3px solid white';
            btn.style.transform = 'scale(1.2)';
        };

        if (userToEdit.color === c) {
            btn.style.border = '3px solid white';
            btn.style.transform = 'scale(1.2)';
        } else {
            btn.style.border = '1px solid rgba(255,255,255,0.3)';
        }
        
        if (c === 'none') {
            btn.style.background = 'transparent';
            btn.innerHTML = '<span style="color: white; font-size: 20px;">×</span>';
        } else {
            btn.style.background = HIGHLIGHT_COLORS[c];
        }

        btn.onclick = () => {
            userToEdit.color = c;
            updateUI();
        };
        colorContainer.appendChild(btn);
    });

    const darkBtn = document.getElementById('theme-dark-btn');
    const lightBtn = document.getElementById('theme-light-btn');
    if (darkBtn && lightBtn) {
        darkBtn.onclick = () => {
            userToEdit.theme = 'dark';
            darkBtn.classList.add('active');
            lightBtn.classList.remove('active');
            applyTheme(bundle);
        };
        lightBtn.onclick = () => {
            userToEdit.theme = 'light';
            lightBtn.classList.add('active');
            darkBtn.classList.remove('active');
            applyTheme(bundle);
        };
    }

    const switchBtn = document.getElementById('switch-user-btn');
    if (switchBtn) {
        switchBtn.onclick = () => {
            sessionStorage.removeItem('sar-current-user');
            window.location.href = 'home.html';
        };
    }

    document.getElementById('save-user-btn').onclick = () => {
        const newFirstName = document.getElementById('user-first-name').value;
        const newLastName = document.getElementById('user-last-name').value;
        const newPin = document.getElementById('user-pin').value;
        const newHandle = document.getElementById('user-handle').value;
        
        if (!newFirstName || !newPin) {
            alert('First Name and PIN are required.');
            return;
        }

        const oldPin = userToEdit.pin;
        const oldHandle = userToEdit.handle;
        const oldFullName = (userToEdit.firstName + ' ' + (userToEdit.lastName || '')).trim();

        const idx = (bundle.accounts || []).findIndex(a => a.pin === oldPin);

        userToEdit.firstName = newFirstName;
        userToEdit.lastName = newLastName;
        userToEdit.pin = newPin;
        userToEdit.handle = newHandle;
        
        // Sync name change to Personnel list (page3)
        if (bundle.pages && bundle.pages.page3) {
            const newName = newHandle || (newFirstName + ' ' + (newLastName || '')).trim();
            bundle.pages.page3.forEach(row => {
                const rowName = (row[0] || '').trim();
                const rowPin = (row[8] || '').trim();
                if ((rowPin && rowPin === oldPin) || rowName === oldHandle || rowName === oldFullName) {
                    row[0] = newName;
                    row[8] = newPin;
                }
            });
        }
        
        if (idx >= 0) {
            bundle.accounts[idx] = userToEdit;
        } else {
            bundle.accounts.push(userToEdit);
        }
        
        saveBundle(bundle);
        
        // If we edited our own account, update the session
        if (oldPin === currentUser.pin) {
            setCurrentUser(userToEdit);
        }

        // If PIN changed and we are Admin editing another user, update the URL to prevent duplicate saves
        if (userPin && userPin !== newPin) {
            const newUrl = window.location.pathname + '?userPin=' + newPin;
            window.history.replaceState(null, '', newUrl);
        }
        
        const status = document.getElementById('save-status');
        if (status) {
            status.textContent = 'Account updated successfully!';
            status.style.color = '#4caf50';
            setTimeout(() => status.textContent = 'Ready.', 3000);
        }
        updateHeaderProfile();
    };
}

function renderUserManagement(container, bundle) {
    container.innerHTML = `
        <div class="table-card" style="max-width: 1200px; margin: 0 auto; background: rgba(0,0,0,0.2); border-radius: 15px; border: 1px solid rgba(255,255,255,0.1); overflow-x: auto;">
            <div style="padding: 20px; background: rgba(0,0,0,0.2); border-bottom: 1px solid rgba(255,255,255,0.1); color: var(--muted); font-size: 0.9rem; text-align: center;">
                User accounts are automatically synchronized with the Personnel list. To add or remove users, please use the Personnel page.
            </div>
            <table class="grid-table" style="width: 100%; border-collapse: collapse; min-width: 600px;">
                <thead>
                    <tr style="background: rgba(255,255,255,0.05);">
                        <th style="padding: 15px; text-align: left; color: var(--accent); width: 60%;">User Name</th>
                        <th style="padding: 15px; text-align: center; color: var(--accent); width: 15%;">File Manager</th>
                        <th style="padding: 15px; text-align: center; color: var(--accent); width: 25%;">Actions</th>
                    </tr>
                </thead>
                <tbody id="user-management-body"></tbody>
            </table>
        </div>
        <div style="text-align: center; margin-top: 30px;">
            <button id="switch-user-btn-mgmt" class="mini-pill" style="padding: 12px 20px; font-size: 1rem; background: rgba(235, 87, 87, 0.1); border-color: rgba(235, 87, 87, 0.4);">Switch User</button>
        </div>
    `;

    const switchBtnMgmt = document.getElementById('switch-user-btn-mgmt');
    if (switchBtnMgmt) {
        switchBtnMgmt.onclick = () => {
            sessionStorage.removeItem('sar-current-user');
            window.location.href = 'home.html';
        };
    }

    const tbody = document.getElementById('user-management-body');
    (bundle.accounts || []).forEach(acc => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid rgba(255,255,255,0.05)';
        
        const tdName = document.createElement('td');
        tdName.style.padding = '12px 15px';
        const nameBtn = document.createElement('button');
        nameBtn.className = 'mini-pill';
        nameBtn.style.fontWeight = 'bold';
        nameBtn.style.whiteSpace = 'normal';
        nameBtn.style.textAlign = 'left';
        nameBtn.style.wordBreak = 'break-word';
        nameBtn.textContent = `${getAccountName(acc)} (@${acc.handle || 'no-handle'})`;
        if (acc.color && acc.color !== 'none') {
            nameBtn.classList.add(`profile-highlight-${acc.color}`);
        }
        nameBtn.onclick = () => {
            const newUrl = window.location.pathname + '?tab=account&userPin=' + acc.pin;
            window.history.replaceState(null, '', newUrl);
            buildUserAccountPage();
        };
        tdName.appendChild(nameBtn);
        tr.appendChild(tdName);
        
        const tdFileManager = document.createElement('td');
        tdFileManager.style.padding = '12px 15px';
        tdFileManager.style.textAlign = 'center';
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.className = 'pill-checkbox';
        chk.checked = !!acc.isFileManager;
        chk.onchange = () => {
            acc.isFileManager = chk.checked;
            saveBundle(bundle);
        };
        tdFileManager.appendChild(chk);
        tr.appendChild(tdFileManager);
        
        const tdActions = document.createElement('td');
        tdActions.style.padding = '12px 15px';
        tdActions.style.textAlign = 'center';
        
        const span = document.createElement('span');
        span.textContent = 'Managed via Personnel';
        span.style.color = 'var(--muted)';
        span.style.fontSize = '0.8rem';
        span.style.display = 'block';
        span.style.lineHeight = '1.2';
        tdActions.appendChild(span);

        tr.appendChild(tdActions);
        
        tbody.appendChild(tr);
    });
}

function buildUserManagementPage() {
    navigateToPage('page8.html?tab=manage');
}
function buildMapsPage() {
  const container = document.querySelector('main');
  if (!container) return;

  let bundle = loadBundle();
  if (!bundle.maps) bundle.maps = [];

  container.innerHTML = `
    <section class="hero">
      <h1>Maps Management</h1>
      <p>Manage your SarTopo/CalTopo maps here. Add a Map ID to embed and fetch shapes. Polygons are imported as segments, lines are not imported.</p>
      <div style="background: rgba(64, 192, 87, 0.1); border-left: 4px solid #40c057; padding: 15px; margin-top: 15px; border-radius: 4px;">
        <p style="margin: 0; font-size: 0.95rem;"><strong>Tip:</strong> We use 'Full Site' mode by default to provide full controls and avoid login issues (403 Errors).</p>
        <p style="margin: 5px 0 0; font-size: 0.9rem; color: var(--muted);">If you still see a login prompt, click the <strong>'Login'</strong> button to open a secure login window. Once logged in, click <strong>'Refresh'</strong> here.</p>
      </div>
    </section>

    <section class="table-card">
      <div class="table-tools">
        <div style="display: flex; gap: 10px; align-items: center; flex-wrap: wrap; width: 100%;">
          <input id="map-id-input" class="pill-input" type="text" placeholder="Map ID (e.g. 0A1B2)" style="flex: 1; min-width: 200px;">
          <input id="map-name-input" class="pill-input" type="text" placeholder="Map Name (Optional)" style="flex: 1; min-width: 200px;">
          <button id="add-map-btn" class="clear-btn">Add Map</button>
        </div>
      </div>

      <div id="maps-list" style="margin-top: 20px; display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 15px;">
        <!-- Map cards will be injected here -->
      </div>
    </section>

    <section id="map-view-section" class="table-card" style="display: none; padding: 0; overflow: hidden; height: 75vh; position: relative; margin-top: 20px; border-radius: 16px;">
      <div class="table-tools" style="padding: 15px; background: var(--header-bg); border-bottom: 1px solid var(--line); display: flex; justify-content: space-between; align-items: center;">
        <h2 id="current-map-title" style="margin: 0; font-size: 1.2rem;">Map View</h2>
        <div class="tool-actions">
          <button id="toggle-map-mode-btn" class="clear-btn" title="Toggle between Embed and Full Site mode (may fix session issues)">Mode: Full Site</button>
          <button id="pop-out-map-btn" class="clear-btn" title="Open in a separate window to bypass all restrictions">Pop-out</button>
          <button id="refresh-map-btn" class="clear-btn">Refresh</button>
          <button id="login-map-btn" class="clear-btn" title="Open a login popup">Login</button>
          <button id="open-new-tab-btn" class="clear-btn">Open Tab</button>
          <button id="fetch-shapes-btn" class="clear-btn">Fetch Shapes</button>
          <button id="clear-map-btn" class="clear-btn" style="color: #ff6b6b;">Clear Map</button>
        </div>
      </div>
      <iframe id="map-iframe" style="width: 100%; height: calc(100% - 62px); border: none;" allow="storage-access; geolocation; clipboard-read; clipboard-write" referrerpolicy="strict-origin-when-cross-origin"></iframe>
    </section>
  `;

  const mapIdInput = document.getElementById('map-id-input');
  const mapNameInput = document.getElementById('map-name-input');
  const addMapBtn = document.getElementById('add-map-btn');
  const mapsList = document.getElementById('maps-list');
  const mapViewSection = document.getElementById('map-view-section');
  const mapIframe = document.getElementById('map-iframe');
  const currentMapTitle = document.getElementById('current-map-title');
  const fetchShapesBtn = document.getElementById('fetch-shapes-btn');
  const openNewTabBtn = document.getElementById('open-new-tab-btn');
  const toggleMapModeBtn = document.getElementById('toggle-map-mode-btn');
  const popOutMapBtn = document.getElementById('pop-out-map-btn');
  const loginMapBtn = document.getElementById('login-map-btn');
  const refreshMapBtn = document.getElementById('refresh-map-btn');
  const clearMapBtn = document.getElementById('clear-map-btn');

  let activeMapId = null;
  let activeMapDomain = 'sartopo.com';
  let isFullMode = true;

  const renderMaps = (skipScroll = false) => {
    if (!bundle.maps || bundle.maps.length === 0) {
      mapsList.innerHTML = '<p style="grid-column: 1/-1; text-align: center; color: var(--muted); padding: 40px;">No map added yet. Enter a Map ID above.</p>';
      mapsList.style.display = 'grid';
      mapViewSection.style.display = 'none';
      mapIframe.src = '';
      activeMapId = null;
      return;
    }
    const map = bundle.maps[0];
    mapsList.innerHTML = '';
    mapsList.style.display = 'none';
    
    viewMap(map.id, map.name, map.domain, skipScroll);
  };

  const viewMap = (id, name, domain, skipScroll = false) => {
    activeMapId = id;
    activeMapDomain = domain || 'sartopo.com';
    currentMapTitle.textContent = name || id;
    const suffix = isFullMode ? '' : '/embed';
    mapIframe.src = `https://${activeMapDomain}/m/${id}${suffix}`;
    mapViewSection.style.display = 'block';
    if (!skipScroll) {
      setTimeout(() => mapViewSection.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
    }
  };

  const popOutMap = (id, domain) => {
    const activeDomain = domain || 'sartopo.com';
    const url = `https://${activeDomain}/m/${id}`;
    window.open(url, `map_${id}`, 'width=1100,height=850,menubar=no,toolbar=no,location=no,status=no');
  };

  addMapBtn.onclick = () => {
    const rawInput = mapIdInput.value.trim();
    const name = mapNameInput.value.trim();
    if (!rawInput) {
      alert('Please enter a Map ID or URL');
      return;
    }
    
    let id = rawInput;
    let domain = 'sartopo.com';
    
    if (rawInput.includes('caltopo.com')) {
      domain = 'caltopo.com';
    } else if (rawInput.includes('sartopo.com')) {
      domain = 'sartopo.com';
    }
    
    // Extract ID from URL like https://sartopo.com/m/ABCDE or https://sartopo.com/m/ABCDE/embed
    const urlMatch = rawInput.match(/\/m\/([A-Za-z0-9-]+)/i);
    if (urlMatch) {
      id = urlMatch[1];
    } else if (rawInput.includes('/')) {
        // Basic cleanup for other URL formats
        const parts = rawInput.split('/');
        id = parts[parts.length - 1] || parts[parts.length - 2];
    }
    
    if (bundle.maps && bundle.maps.length > 0) {
        bundle.maps = [{ id, name, domain }];
    } else {
        bundle.maps = [{ id, name, domain }];
    }
    saveBundle(bundle);
    mapIdInput.value = '';
    mapNameInput.value = '';
    renderMaps();
  };

  clearMapBtn.onclick = () => {
    if (confirm('Are you sure you want to clear this map from the project?')) {
      bundle.maps = [];
      saveBundle(bundle);
      renderMaps();
    }
  };

  toggleMapModeBtn.onclick = () => {
    isFullMode = !isFullMode;
    toggleMapModeBtn.textContent = `Mode: ${isFullMode ? 'Full Site' : 'Embed'}`;
    if (activeMapId) {
      viewMap(activeMapId, currentMapTitle.textContent, activeMapDomain);
    }
  };

  refreshMapBtn.onclick = () => {
    if (activeMapId) {
      const currentSrc = mapIframe.src;
      mapIframe.src = '';
      setTimeout(() => { mapIframe.src = currentSrc; }, 10);
    }
  };

  openNewTabBtn.onclick = () => {
    if (activeMapId) {
      window.open(`https://${activeMapDomain}/m/${activeMapId}`, '_blank');
    }
  };

  popOutMapBtn.onclick = () => {
    if (activeMapId) {
      popOutMap(activeMapId, activeMapDomain);
    }
  };

  loginMapBtn.onclick = () => {
    if (activeMapId) {
      window.open(`https://${activeMapDomain}/login`, 'sartopo_login', 'width=500,height=600');
    }
  };

    fetchShapesBtn.onclick = async () => {
    if (!activeMapId) return;
    fetchShapesBtn.disabled = true;
    fetchShapesBtn.textContent = 'Fetching...';
    
    try {
      let data = null;
      let usedProxy = false;
      const proxyUrl = getSartopoProxy();

      if (proxyUrl) {
        try {
          const proxyResp = await fetch(`${proxyUrl}?mapId=${activeMapId}&domain=${activeMapDomain}`);
          if (proxyResp.ok) {
            data = await proxyResp.json();
            usedProxy = true;
          } else {
            const errData = await proxyResp.json().catch(() => ({}));
            console.warn('Proxy fetch failed, falling back to direct fetch', errData);
          }
        } catch (proxyErr) {
          console.warn('Proxy unreachable, falling back to direct fetch', proxyErr);
        }
      }

      if (!data) {
        // Direct fetch (may fail due to CORS if SarTopo doesn't explicitly allow our origin)
        const response = await fetch(`https://${activeMapDomain}/api/v1/map/${activeMapId}/features`, {
          credentials: 'include'
        });
        
        if (!response.ok) {
          if (response.status === 403 || response.status === 401 || response.status === 404) {
             let details = '';
             if (response.status === 404) details = ' (Map not found or requires login)';
             throw new Error(`Login Required or Access Denied${details}. This is usually a CORS issue. Please run the middleman.py proxy and configure it in Settings.`);
          }
          throw new Error(`Failed to fetch from ${activeMapDomain} (Status: ${response.status})`);
        }
        data = await response.json();
      }
      
      const features = (data.features || []).filter(f => {
        if (!f.geometry) return false;
        const props = f.properties || {};
        const isAssignment = props.class === 'Assignment' || props.type === 'Assignment' || !!props.assignment;
        const isPolygon = f.geometry.type === 'Polygon' || f.geometry.type === 'MultiPolygon';
        // We import all assignments (lines or polygons) and any other polygons
        return isAssignment || isPolygon;
      });
      
      if (features.length === 0) {
          alert('No compatible shapes (Assignments or Polygons) found on this map.');
      } else {
          showSarTopoShapesPopup(features);
      }
    } catch (err) {
      console.error(err);
      alert(`Could not fetch shapes: ${err.message}\n\nThis is usually caused by browser security (CORS) when the map is private.\n\nTo resolve:\n1. Recommended: In SarTopo, go to "Map Settings" and set "Access" to "Anyone with the link can view".\n2. If it must stay private, ensure you are logged in AND click the EYE ICON (or lock icon) in the address bar to "Allow cookies" for this site.\n3. Make sure you are using HTTPS if SarTopo is HTTPS.`);
    } finally {
      fetchShapesBtn.disabled = false;
      fetchShapesBtn.textContent = 'Fetch Shapes';
    }
  };

  renderMaps(true);
}

let isSyncing = false;

async function syncWithServer() {
    if (isSyncing) return;
    const bucket = getSyncBucket();
    const bucketUrl = `${KVDB_BASE_URL}/${bucket}`;
    
    isSyncing = true;
    try {
        const currentUser = getCurrentUser();
        const deviceId = getDeviceId();
        
        // 1. Check active user status
        if (currentUser && currentUser.pin) {
            const resp = await fetch(`${bucketUrl}/user-${currentUser.pin}`);
            if (resp.ok) {
                const remoteDeviceId = await resp.text();
                if (remoteDeviceId && remoteDeviceId !== deviceId) {
                    // Different device is active!
                    alert("This user has been selected on another device. Switching user...");
                    sessionStorage.removeItem('sar-current-user');
                    window.location.href = 'home.html';
                    return;
                }
            }
        }
        
        // 2. Sync bundle
        const resp = await fetch(`${bucketUrl}/bundle`);
        if (resp.ok) {
            const serverBundle = await resp.json();
            
            if (serverBundle) {
                const localBundle = loadBundle();
                // Simple check: if server bundle is different, update local
                if (JSON.stringify(serverBundle) !== JSON.stringify(localBundle)) {
                    localStorage.setItem(BUNDLE_STORAGE_KEY, JSON.stringify(serverBundle));
                    
                    // Update file list if needed
                    const files = getSavedFiles();
                    if (serverBundle.fileName) {
                        files[serverBundle.fileName] = {
                            bundle: serverBundle,
                            lastModified: new Date().toLocaleString()
                        };
                        localStorage.setItem(FILE_LIST_STORAGE_KEY, JSON.stringify(files));
                    }
                    
                    // Refresh UI if on a page that shows data
                    if (typeof recalculateEverything === 'function') recalculateEverything();
                    
                    const pk = pageKey();
                    if (pk === 'index') {
                        if (typeof buildRegionsTable === 'function') buildRegionsTable();
                    } else if (pk === 'page2') {
                        if (typeof buildSegmentsTable === 'function') buildSegmentsTable();
                    } else if (pk === 'page3') {
                        if (typeof buildPersonnelTable === 'function') buildPersonnelTable();
                    } else if (pk === 'page4') {
                        if (typeof buildSearchLogTable === 'function') buildSearchLogTable();
                    } else if (pk === 'home') {
                        if (typeof buildHomePage === 'function') buildHomePage();
                    }
                }
            }
        }
    } catch (err) {
        console.error("Sync failed:", err);
    } finally {
        isSyncing = false;
    }
}

async function pushBundleToServer(bundle) {
    const bucket = getSyncBucket();
    const bucketUrl = `${KVDB_BASE_URL}/${bucket}`;
    
    try {
        await fetch(`${bucketUrl}/bundle`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(bundle)
        });
    } catch (err) {
        console.error("Push to sync service failed:", err);
    }
}

async function notifyActiveUser(user) {
    const bucket = getSyncBucket();
    const bucketUrl = `${KVDB_BASE_URL}/${bucket}`;
    if (!user || !user.pin) return;
    
    try {
        await fetch(`${bucketUrl}/user-${user.pin}`, {
            method: 'PUT',
            body: getDeviceId()
        });
    } catch (err) {
        console.error("Notify active user failed:", err);
    }
}

// Start sync loop
setInterval(syncWithServer, 5000);
// Initial sync
setTimeout(syncWithServer, 1000);
