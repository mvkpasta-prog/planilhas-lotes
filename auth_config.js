const CLIENT_ID = '454403213456-du5uh77kn63d2b04n9b0brvt87e4pc0i.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly';

let tokenClient;
let gapiInited = false;
let gisInited = false;
let accessToken = null;

// Callback to notify when auth is ready or changed
let onAuthSuccess = null;

function gapiLoaded() {
  gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
  // We don't need an API Key if we are only using OAuth for private sheets, 
  // but usually discovery docs are needed.
  await gapi.client.init({
    discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  });
  gapiInited = true;
  checkAuth();
}

function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: (resp) => {
      if (resp.error !== undefined) {
        throw (resp);
      }
      accessToken = resp.access_token;
      localStorage.setItem('google_access_token', accessToken);
      const expiry = Date.now() + (resp.expires_in * 1000);
      localStorage.setItem('google_token_expiry', expiry);
      
      updateUIAsSignedIn();
      if (onAuthSuccess) onAuthSuccess();
    },
  });
  gisInited = true;
  checkAuth();
}

function checkAuth() {
  if (gapiInited && gisInited) {
    // Check if we have a valid token in storage
    const storedToken = localStorage.getItem('google_access_token');
    const expiry = localStorage.getItem('google_token_expiry');
    
    if (storedToken && expiry && Date.now() < parseInt(expiry)) {
      accessToken = storedToken;
      gapi.client.setToken({access_token: accessToken});
      updateUIAsSignedIn();
      if (onAuthSuccess) onAuthSuccess();
    } else {
      updateUIAsSignedOut();
    }
  }
}

function handleAuthClick() {
  tokenClient.requestAccessToken({prompt: 'consent'});
}

function handleSignoutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revoke(token.access_token);
    gapi.client.setToken('');
    accessToken = null;
    localStorage.removeItem('google_access_token');
    localStorage.removeItem('google_token_expiry');
    updateUIAsSignedOut();
  }
}

function updateUIAsSignedIn() {
  const loginBtn = document.getElementById('authorize_button');
  const logoutBtn = document.getElementById('signout_button');
  const content = document.getElementById('content-wrapper'); // We'll wrap main content
  
  if(loginBtn) loginBtn.style.display = 'none';
  if(logoutBtn) logoutBtn.style.display = 'block';
  if(content) content.style.display = 'block';
}

function updateUIAsSignedOut() {
  const loginBtn = document.getElementById('authorize_button');
  const logoutBtn = document.getElementById('signout_button');
  const content = document.getElementById('content-wrapper');
  
  if(loginBtn) loginBtn.style.display = 'block';
  if(logoutBtn) logoutBtn.style.display = 'none';
  // Optionally hide content if you want to force login
  // if(content) content.style.display = 'none'; 
}

// Helper to process Sheets API values to match existing code structure
function processSheetValues(values) {
  if (!values || values.length === 0) return {arrays:[], headersOrig:[], headersNorm:[], objects:[]};
  
  const headersOrig = values[0].map(h => h || '');
  const headersNorm = headersOrig.map(h => normalizeHeader(h));
  
  const arrays = values.slice(1);
  const objects = arrays.map(row => {
    const obj = {};
    headersNorm.forEach((key, idx) => {
      obj[key] = row[idx] !== undefined ? row[idx] : '';
    });
    return obj;
  });
  
  return {arrays, headersOrig, headersNorm, objects};
}

// Map GIDs to Sheet Names (since API v4 uses names for ranges usually, or we can use the spreadsheets.get to find names)
// We will fetch metadata once.
let sheetIdToTitle = {};

async function ensureSheetTitles(spreadsheetId) {
  if (Object.keys(sheetIdToTitle).length > 0) return;
  
  try {
    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId
    });
    
    response.result.sheets.forEach(sheet => {
      sheetIdToTitle[sheet.properties.sheetId] = sheet.properties.title;
    });
  } catch (err) {
    console.error('Error fetching spreadsheet metadata:', err);
    throw err;
  }
}

async function fetchSheetByGidV4(spreadsheetId, gid) {
  if (!accessToken) throw new Error('Not authenticated');
  
  await ensureSheetTitles(spreadsheetId);
  const sheetTitle = sheetIdToTitle[gid];
  
  if (!sheetTitle) throw new Error(`Sheet with GID ${gid} not found`);
  
  const response = await gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheetId,
    range: `'${sheetTitle}'!A:Z`, // Fetch all columns
  });
  
  return processSheetValues(response.result.values);
}
