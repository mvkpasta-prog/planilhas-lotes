// Configurações de OAuth
const CLIENT_ID = '454403213456-du5uh77kn63d2b04n9b0brvt87e4pc0i.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly';

let tokenClient;
let gapiInited = false;
let gisInited = false;
let accessToken = null;

// Callback global para páginas usarem quando o auth estiver pronto
let onAuthSuccess = null;

// -------------------------
// Inicialização GAPI (Sheets API v4)
// -------------------------
function gapiLoaded() {
  gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
  // Não usamos API Key porque o acesso é via OAuth a planilhas privadas.
  await gapi.client.init({
    discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  });
  gapiInited = true;
  checkAuth();
}

// -------------------------
// Inicialização Google Identity Services
// -------------------------
function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: (resp) => {
      if (resp.error !== undefined) {
        throw resp;
      }

      accessToken = resp.access_token;
      localStorage.setItem('google_access_token', accessToken);

      const expiry = Date.now() + (resp.expires_in * 1000);
      localStorage.setItem('google_token_expiry', expiry);

      // Marcamos que esse usuário já autorizou o app pelo menos uma vez
      localStorage.setItem('google_has_auth', 'true');

      gapi.client.setToken({ access_token: accessToken });

      updateUIAsSignedIn();
      if (onAuthSuccess) onAuthSuccess();
    },
  });

  gisInited = true;
  checkAuth();
}

// -------------------------
// Lógica de autenticação + renovação do token
// -------------------------
function checkAuth() {
  if (!(gapiInited && gisInited)) return;

  const storedToken = localStorage.getItem('google_access_token');
  const expiryStr   = localStorage.getItem('google_token_expiry');
  const hasAuth     = localStorage.getItem('google_has_auth') === 'true';

  const now    = Date.now();
  const expiry = expiryStr ? parseInt(expiryStr, 10) : 0;

  // Margem de 1 minuto para não usar token no limite
  const stillValid = storedToken && expiry && now < (expiry - 60_000);

  if (stillValid) {
    // Token ainda válido: reaproveita
    accessToken = storedToken;
    gapi.client.setToken({ access_token: accessToken });

    updateUIAsSignedIn();
    if (onAuthSuccess) onAuthSuccess();
  } else if (hasAuth) {
    // Já autorizou no passado: tenta renovar em silêncio (sem pop-up)
    tokenClient.requestAccessToken({ prompt: '' });
  } else {
    // Nunca autorizou: mostra estado deslogado
    updateUIAsSignedOut();
  }
}

// Clique no botão de Login (primeira vez ou quando quiser trocar de conta)
function handleAuthClick() {
  tokenClient.requestAccessToken({ prompt: 'consent' });
}

// Clique no botão de Sair
function handleSignoutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revoke(token.access_token);
    gapi.client.setToken('');
  }

  accessToken = null;
  localStorage.removeItem('google_access_token');
  localStorage.removeItem('google_token_expiry');
  localStorage.removeItem('google_has_auth'); // sai de tudo mesmo

  updateUIAsSignedOut();
}

// -------------------------
// Atualização da UI (botões / conteúdo)
// -------------------------
function updateUIAsSignedIn() {
  const loginBtn  = document.getElementById('authorize_button');
  const logoutBtn = document.getElementById('signout_button');
  const content   = document.getElementById('content-wrapper'); // opcional

  if (loginBtn)  loginBtn.style.display  = 'none';
  if (logoutBtn) logoutBtn.style.display = 'block';
  if (content)   content.style.display   = 'block';
}

function updateUIAsSignedOut() {
  const loginBtn  = document.getElementById('authorize_button');
  const logoutBtn = document.getElementById('signout_button');
  const content   = document.getElementById('content-wrapper');

  if (loginBtn)  loginBtn.style.display  = 'block';
  if (logoutBtn) logoutBtn.style.display = 'none';
  // Se quiser esconder tudo quando não estiver logado, descomente:
  // if (content)   content.style.display   = 'none';
}

// -------------------------
// Helpers para Sheets API v4
// -------------------------
function processSheetValues(values) {
  if (!values || values.length === 0) {
    return { arrays: [], headersOrig: [], headersNorm: [], objects: [] };
  }

  const headersOrig = values[0].map(h => h || '');
  // Usa normalizeHeader das páginas HTML (já existente nelas)
  const headersNorm = headersOrig.map(h => normalizeHeader(h));

  const arrays = values.slice(1);
  const objects = arrays.map(row => {
    const obj = {};
    headersNorm.forEach((key, idx) => {
      obj[key] = row[idx] !== undefined ? row[idx] : '';
    });
    return obj;
  });

  return { arrays, headersOrig, headersNorm, objects };
}

// Cache para mapear GID -> título da aba
let sheetIdToTitle = {};

async function ensureSheetTitles(spreadsheetId) {
  if (Object.keys(sheetIdToTitle).length > 0) return;

  try {
    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    response.result.sheets.forEach(sheet => {
      sheetIdToTitle[sheet.properties.sheetId] = sheet.properties.title;
    });
  } catch (err) {
    console.error('Error fetching spreadsheet metadata:', err);
    throw err;
  }
}

// Função usada nas páginas: fetchSheetByGidV4(SHEET_ID, gid)
async function fetchSheetByGidV4(spreadsheetId, gid) {
  if (!accessToken) throw new Error('Not authenticated');

  await ensureSheetTitles(spreadsheetId);
  const sheetTitle = sheetIdToTitle[gid];

  if (!sheetTitle) {
    throw new Error(`Sheet with GID ${gid} not found`);
  }

  const response = await gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheetId,
    range: `'${sheetTitle}'!A:Z`, // pega da coluna A até Z
  });

  return processSheetValues(response.result.values);
}
