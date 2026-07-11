// Booking Management Web App
// Single-file Express server + embedded frontend
// Backend: Google Sheets via service account
// Deploy on Railway

const express = require('express');
const https = require('https');
const crypto = require('crypto');
const app = express();

app.use(express.json());

// ===== CONFIG =====
const SHEET_ID = process.env.SHEET_ID;
const SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

const MAIN_SHEET = 'Booking';
const CUSTOMER_SHEET = 'CustomerName';
const STORE_SHEET = 'ListOfStore';
const SUPPLIER_SHEET = 'Supplier';
const TRX_SHEET = 'BookingTransaction';

// ===== NATIVE GOOGLE AUTH (bypasses gaxios/undici "Premature close" bug) =====
let tokenCache = { token: null, expiresAt: 0 };

function httpsRequest(opts, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(opts, (res) => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString('utf8') }));
      res.on('error', reject);
    });
    req.on('error', reject);
    req.setTimeout(30000, () => req.destroy(new Error('Request timeout after 30s')));
    if (body) req.write(body);
    req.end();
  });
}

async function getAccessToken() {
  if (tokenCache.token && tokenCache.expiresAt > Date.now() + 60000) return tokenCache.token;
  if (!SERVICE_ACCOUNT_EMAIL || !PRIVATE_KEY) throw new Error('Missing GOOGLE_CLIENT_EMAIL or GOOGLE_PRIVATE_KEY');

  const now = Math.floor(Date.now() / 1000);
  const b64url = o => Buffer.from(JSON.stringify(o)).toString('base64url');
  const header = { alg: 'RS256', typ: 'JWT' };
  const claims = {
    iss: SERVICE_ACCOUNT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600, iat: now
  };
  const signingInput = b64url(header) + '.' + b64url(claims);
  const signature = crypto.createSign('RSA-SHA256').update(signingInput).sign(PRIVATE_KEY).toString('base64url');
  const jwt = signingInput + '.' + signature;

  const postBody = 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + encodeURIComponent(jwt);
  const res = await httpsRequest({
    hostname: 'oauth2.googleapis.com', port: 443, path: '/token', method: 'POST', family: 4,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postBody), 'Accept': 'application/json' }
  }, postBody);

  let data;
  try { data = JSON.parse(res.body); } catch (e) { throw new Error('OAuth response not JSON (HTTP ' + res.status + '): ' + res.body.slice(0, 200)); }
  if (res.status !== 200 || !data.access_token) throw new Error('OAuth failed (HTTP ' + res.status + '): ' + (data.error_description || data.error || res.body));

  tokenCache.token = data.access_token;
  tokenCache.expiresAt = Date.now() + (data.expires_in * 1000);
  console.log('[Auth] Access token acquired via native OAuth (expires in ' + data.expires_in + 's)');
  return tokenCache.token;
}

// ===== NATIVE GOOGLE SHEETS API HELPERS =====
const SHEETS_BASE = '/v4/spreadsheets/' + SHEET_ID;

async function sheetsAPI(path, method, body) {
  const token = await getAccessToken();
  const opts = {
    hostname: 'sheets.googleapis.com', port: 443, method: method || 'GET', family: 4,
    path: SHEETS_BASE + path,
    headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }
  };
  let payload;
  if (body) {
    payload = JSON.stringify(body);
    opts.headers['Content-Type'] = 'application/json';
    opts.headers['Content-Length'] = Buffer.byteLength(payload);
  }
  const res = await httpsRequest(opts, payload);
  let data;
  try { data = JSON.parse(res.body); } catch (e) { throw new Error('Sheets API not JSON (HTTP ' + res.status + '): ' + res.body.slice(0, 300)); }
  if (res.status < 200 || res.status >= 300) throw new Error('Sheets API error (HTTP ' + res.status + '): ' + (data.error?.message || res.body.slice(0, 300)));
  return data;
}

function encRange(range) {
  return encodeURIComponent(range).replace(/%21/g, '!').replace(/%3A/g, ':');
}

async function sheetsGetValues(range) {
  return sheetsAPI('/values/' + encRange(range));
}

async function sheetsBatchGetValues(ranges) {
  const qs = ranges.map(r => 'ranges=' + encRange(r)).join('&');
  return sheetsAPI('/values:batchGet?' + qs);
}

async function sheetsAppendValues(range, values) {
  const qs = 'valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS';
  return sheetsAPI('/values/' + encRange(range) + ':append?' + qs, 'POST', { values });
}

async function sheetsUpdateValues(range, values) {
  const qs = 'valueInputOption=USER_ENTERED';
  return sheetsAPI('/values/' + encRange(range) + '?' + qs, 'PUT', { values });
}

async function sheetsGetMeta() {
  return sheetsAPI('');
}

async function sheetsBatchUpdate(requests) {
  return sheetsAPI(':batchUpdate', 'POST', { requests });
}

// ===== CACHE FOR LOOKUPS (refreshed every 60s) =====
let cache = { data: null, ts: 0 };
const CACHE_MS = 60 * 1000;

async function getLookups() {
  if (cache.data && Date.now() - cache.ts < CACHE_MS) return cache.data;

  const res = await sheetsBatchGetValues([
    `${CUSTOMER_SHEET}!A:E`,
    `${STORE_SHEET}!A:D`,
    `${SUPPLIER_SHEET}!A:C`,
  ]);

  const custRows = res.valueRanges[0].values || [];
  const storeRows = res.valueRanges[1].values || [];
  const supRows = res.valueRanges[2].values || [];

  // CustomerName: A=No, B=FullName, C=StoreID, D=StoreName(105-VAL), E=Region
  const customers = custRows.slice(1)
    .filter(r => r[1])
    .map(r => ({
      no: r[0] || '',
      name: r[1] || '',
      storeId: String(r[2] || ''),
      storeName: r[3] || '',
      region: r[4] || '',
    }));

  // ListOfStore: A=Region, B=Area, C=StoreID, D=StoreName(clean)
  const stores = storeRows.slice(1)
    .filter(r => r[3])
    .map(r => ({
      region: r[0] || '',
      area: r[1] || '',
      storeId: String(r[2] || ''),
      storeName: r[3] || '',
    }));

  // Supplier: A=Code, B=Name, C=Category(Dept)
  const suppliers = [...new Set(supRows.slice(1).map(r => r[1]).filter(Boolean))].sort();
  const depts = [...new Set(supRows.slice(1).map(r => r[2]).filter(Boolean))].sort();

  cache.data = { customers, stores, suppliers, depts };
  cache.ts = Date.now();
  return cache.data;
}

function invalidateCache() { cache.ts = 0; }

// ===== HELPERS =====
function todayMDY() {
  const d = new Date();
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}
function trim(v) { return typeof v === 'string' ? v.trim() : v; }
function upper(v) { return typeof v === 'string' ? v.toUpperCase() : v; }

// ===== READ BOOKINGS =====
async function readBookings() {
  const res = await sheetsGetValues(`${MAIN_SHEET}!A:N`);
  const rows = res.values || [];
  if (rows.length <= 1) return [];

  return rows.slice(1).map((r, i) => ({
    rowIndex: i + 2,
    Date_Booked: r[0] || '',
    Customer_No: r[1] || '',
    Customer_Name: r[2] || '',
    Region: r[3] || '',
    Area: r[4] || '',
    Store_Delivery: r[5] || '',
    Dept: r[6] || '',
    Supplier: r[7] || '',
    Booking_No: r[8] || '',
    Deals: r[9] || '',
    Total_Booked_Amount: r[10] || '',
    Remarks: r[11] || '',
    Logged_By: r[12] || '',
    Last_Edited_By: r[13] || '',
  }));
}

// ===== READ USERS (cached 60s) =====
let userCache = { data: null, ts: 0 };
async function readUsers() {
  if (userCache.data && Date.now() - userCache.ts < CACHE_MS) return userCache.data;
  const res = await sheetsGetValues('Users!A:E');
  const rows = res.values || [];
  const users = rows.slice(1).filter(r => r[0]).map(r => ({
    username: String(r[0]).trim().toLowerCase(),
    password: String(r[1] || ''),
    fullName: r[2] || r[0],
    role: (r[3] || 'user').toLowerCase(),
    storeAccess: String(r[4] || '').trim(),  // Store name, Area name, "All", or empty
  }));
  userCache.data = users;
  userCache.ts = Date.now();
  return users;
}

// Resolve a user's storeAccess into a list of allowed store names.
// Returns null = unrestricted (all stores). Returns [] or [names] = limited.
async function resolveStoreAccess(user) {
  if (!user) return null;
  if (user.role === 'admin') return null;  // admins see everything
  const access = (user.storeAccess || '').trim();
  if (!access || access.toLowerCase() === 'all') return null;

  const lookups = await getLookups();
  const lc = access.toLowerCase();

  // Try store name match first
  const storeMatch = lookups.stores.find(s => s.storeName.toLowerCase() === lc);
  if (storeMatch) return [storeMatch.storeName];

  // Fall back to area match -> all stores in that area
  const areaStores = lookups.stores.filter(s => s.area.toLowerCase() === lc).map(s => s.storeName);
  if (areaStores.length) return areaStores;

  // Value doesn't match anything -> block all access (safer than allowing everything)
  return [];
}

// ===== SIMPLE SESSION (in-memory token store) =====
// Tokens stored server-side. Cleared on app restart (users will need to log in again).
const sessions = new Map(); // token -> { username, fullName, role, ts }
const SESSION_DAYS = 30;

function makeToken() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36) + Math.random().toString(36).slice(2);
}

function authUser(req) {
  const token = req.headers['x-auth-token'];
  if (!token) return null;
  const s = sessions.get(token);
  if (!s) return null;
  if (Date.now() - s.ts > SESSION_DAYS * 86400000) {
    sessions.delete(token);
    return null;
  }
  return s;
}

function requireAuth(req, res, next) {
  const u = authUser(req);
  if (!u) return res.status(401).json({ error: 'Not logged in' });
  req.user = u;
  next();
}

// ===== AUTH ROUTES =====
app.post('/api/login', async (req, res) => {
  try {
    const username = String(req.body.username || '').trim().toLowerCase();
    const password = String(req.body.password || '');
    if (!username || !password) return res.status(400).json({ error: 'Username and password required' });

    const users = await readUsers();
    const u = users.find(x => x.username === username && x.password === password);
    if (!u) return res.status(401).json({ error: 'Invalid username or password' });

    const token = makeToken();
    sessions.set(token, { username: u.username, fullName: u.fullName, role: u.role, storeAccess: u.storeAccess, ts: Date.now() });
    res.json({ token, username: u.username, fullName: u.fullName, role: u.role, storeAccess: u.storeAccess });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/logout', (req, res) => {
  const token = req.headers['x-auth-token'];
  if (token) sessions.delete(token);
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  const u = authUser(req);
  if (!u) return res.status(401).json({ error: 'Not logged in' });
  res.json({ username: u.username, fullName: u.fullName, role: u.role, storeAccess: u.storeAccess });
});

// ===== API ROUTES =====
app.get('/api/lookups', requireAuth, async (req, res) => {
  try { res.json(await getLookups()); }
  catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.get('/api/bookings', requireAuth, async (req, res) => {
  try {
    const all = await readBookings();
    const allowed = await resolveStoreAccess(req.user);
    if (allowed === null) return res.json(all);
    const set = new Set(allowed.map(s => s.toLowerCase()));
    res.json(all.filter(b => set.has(String(b.Store_Delivery).toLowerCase())));
  }
  catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/bookings', requireAuth, async (req, res) => {
  try {
    const b = req.body;
    Object.keys(b).forEach(k => { b[k] = trim(b[k]); });
    ['Customer_Name','Region','Area','Store_Delivery','Dept','Supplier'].forEach(f => {
      if (b[f]) b[f] = upper(b[f]);
    });

    const required = ['Customer_Name','Customer_No','Region','Area','Store_Delivery','Dept','Supplier','Deals','Total_Booked_Amount'];
    for (const f of required) {
      if (!b[f] && b[f] !== 0) return res.status(400).json({ error: `${f} is required` });
    }
    if (isNaN(parseFloat(b.Total_Booked_Amount))) {
      return res.status(400).json({ error: 'Total Booked Amount must be a number' });
    }

    // Enforce store access
    const allowed = await resolveStoreAccess(req.user);
    if (allowed !== null) {
      const set = new Set(allowed.map(s => s.toLowerCase()));
      if (!set.has(String(b.Store_Delivery).toLowerCase())) {
        return res.status(403).json({ error: `You don't have access to store: ${b.Store_Delivery}` });
      }
    }

    // Duplicate check only if Booking_No is provided
    if (b.Booking_No) {
      const existing = await readBookings();
      if (existing.some(r => String(r.Booking_No).toLowerCase() === String(b.Booking_No).toLowerCase())) {
        return res.status(409).json({ error: `Booking# ${b.Booking_No} already exists` });
      }
    }

    const row = [
      todayMDY(),
      b.Customer_No,
      b.Customer_Name,
      b.Region,
      b.Area,
      b.Store_Delivery,
      b.Dept,
      b.Supplier,
      b.Booking_No || '',
      b.Deals,
      parseFloat(b.Total_Booked_Amount),
      b.Remarks || '',
      req.user.username,  // Logged_By
      '',                  // Last_Edited_By empty on create
    ];

    await sheetsAppendValues(`${MAIN_SHEET}!A:N`, [row]);

    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.put('/api/bookings/:rowIndex', requireAuth, async (req, res) => {
  try {
    const rowIndex = parseInt(req.params.rowIndex);
    const b = req.body;
    Object.keys(b).forEach(k => { b[k] = trim(b[k]); });

    const existing = await readBookings();
    const target = existing.find(r => r.rowIndex === rowIndex);
    if (!target) return res.status(404).json({ error: 'Booking not found' });

    // Ownership check: non-admins can only edit their own
    if (req.user.role !== 'admin' && target.Logged_By.toLowerCase() !== req.user.username) {
      return res.status(403).json({ error: 'You can only edit your own bookings' });
    }

    // Enforce store access on the edited store too
    const allowed = await resolveStoreAccess(req.user);
    if (allowed !== null) {
      const set = new Set(allowed.map(s => s.toLowerCase()));
      if (!set.has(String(b.Store_Delivery).toLowerCase())) {
        return res.status(403).json({ error: `You don't have access to store: ${b.Store_Delivery}` });
      }
    }

    if (b.Booking_No && existing.some(r => r.rowIndex !== rowIndex && String(r.Booking_No).toLowerCase() === String(b.Booking_No).toLowerCase())) {
      return res.status(409).json({ error: `Booking# ${b.Booking_No} already exists` });
    }

    const row = [
      b.Date_Booked || todayMDY(),
      b.Customer_No,
      b.Customer_Name,
      b.Region,
      b.Area,
      b.Store_Delivery,
      b.Dept,
      b.Supplier,
      b.Booking_No || '',
      b.Deals,
      parseFloat(b.Total_Booked_Amount),
      b.Remarks || '',
      target.Logged_By,         // preserve original creator
      req.user.username,         // Last_Edited_By
    ];

    await sheetsUpdateValues(`${MAIN_SHEET}!A${rowIndex}:N${rowIndex}`, [row]);
    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.delete('/api/bookings/:rowIndex', requireAuth, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Only admins can delete bookings' });
    }
    const rowIndex = parseInt(req.params.rowIndex);
    const meta = await sheetsGetMeta();
    const ms = meta.sheets.find(s => s.properties.title === MAIN_SHEET);
    if (!ms) throw new Error('Main sheet not found');

    await sheetsBatchUpdate([{
      deleteDimension: {
        range: {
          sheetId: ms.properties.sheetId,
          dimension: 'ROWS',
          startIndex: rowIndex - 1,
          endIndex: rowIndex,
        },
      },
    }]);

    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ===== TRANSACTIONS =====
async function readTransactions() {
  const res = await sheetsGetValues(`${TRX_SHEET}!A:P`);
  const rows = res.values || [];
  if (rows.length <= 1) return [];

  return rows.slice(1).map((r, i) => ({
    rowIndex: i + 2,
    Date_Transacted: r[0] || '',
    Customer_No: r[1] || '',
    Name: r[2] || '',
    Region: r[3] || '',
    Area: r[4] || '',
    Store_Delivery: r[5] || '',
    Dept: r[6] || '',
    Supplier: r[7] || '',
    Booking_No: r[8] || '',
    Deals: r[9] || '',
    Transacted_Amount_Gross: r[10] || '',
    Transacted_Discount_Net: r[11] || '',
    Transacted_Amount_Net: r[12] || '',
    TRX_Number: r[13] || '',
    Logged_By: r[14] || '',
    Last_Edited_By: r[15] || '',
  }));
}

app.get('/api/transactions', requireAuth, async (req, res) => {
  try {
    const all = await readTransactions();
    const allowed = await resolveStoreAccess(req.user);
    if (allowed === null) return res.json(all);
    const set = new Set(allowed.map(s => s.toLowerCase()));
    res.json(all.filter(t => set.has(String(t.Store_Delivery).toLowerCase())));
  }
  catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// Look up a Booking# from the Booking sheet to auto-fill the transaction form
app.get('/api/booking-lookup/:bookingNo', requireAuth, async (req, res) => {
  try {
    const target = String(req.params.bookingNo).trim().toLowerCase();
    if (!target) return res.status(400).json({ error: 'Booking# required' });
    const all = await readBookings();
    const b = all.find(x => String(x.Booking_No).trim().toLowerCase() === target);
    if (!b) return res.status(404).json({ error: `Booking# ${req.params.bookingNo} not found in Booking sheet` });

    // Enforce store access
    const allowed = await resolveStoreAccess(req.user);
    if (allowed !== null) {
      const set = new Set(allowed.map(s => s.toLowerCase()));
      if (!set.has(String(b.Store_Delivery).toLowerCase())) {
        return res.status(403).json({ error: `Booking belongs to ${b.Store_Delivery} — outside your store access` });
      }
    }
    res.json(b);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/transactions', requireAuth, async (req, res) => {
  try {
    const b = req.body;
    Object.keys(b).forEach(k => { b[k] = trim(b[k]); });

    const required = ['Date_Transacted','Customer_No','Name','Region','Area','Store_Delivery','Dept','Supplier','Booking_No','Transacted_Amount_Gross','Transacted_Amount_Net','TRX_Number'];
    for (const f of required) {
      if (!b[f] && b[f] !== 0) return res.status(400).json({ error: `${f} is required` });
    }
    if (isNaN(parseFloat(b.Transacted_Amount_Gross))) return res.status(400).json({ error: 'Gross amount must be a number' });
    if (b.Transacted_Discount_Net && isNaN(parseFloat(b.Transacted_Discount_Net))) return res.status(400).json({ error: 'Discount must be a number' });

    // Enforce store access
    const allowedT = await resolveStoreAccess(req.user);
    if (allowedT !== null) {
      const setT = new Set(allowedT.map(s => s.toLowerCase()));
      if (!setT.has(String(b.Store_Delivery).toLowerCase())) {
        return res.status(403).json({ error: `You don't have access to store: ${b.Store_Delivery}` });
      }
    }

    const gross = parseFloat(b.Transacted_Amount_Gross) || 0;
    const disc = parseFloat(b.Transacted_Discount_Net) || 0;
    const net = gross - disc;

    const row = [
      b.Date_Transacted,
      b.Customer_No,
      b.Name,
      b.Region,
      b.Area,
      b.Store_Delivery,
      b.Dept,
      b.Supplier,
      b.Booking_No,
      b.Deals || '',
      gross,
      disc,
      net,
      b.TRX_Number,
      req.user.username,  // Logged_By
      '',                  // Last_Edited_By
    ];

    await sheetsAppendValues(`${TRX_SHEET}!A:P`, [row]);

    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.put('/api/transactions/:rowIndex', requireAuth, async (req, res) => {
  try {
    const rowIndex = parseInt(req.params.rowIndex);
    const b = req.body;
    Object.keys(b).forEach(k => { b[k] = trim(b[k]); });

    const existing = await readTransactions();
    const target = existing.find(r => r.rowIndex === rowIndex);
    if (!target) return res.status(404).json({ error: 'Transaction not found' });

    if (req.user.role !== 'admin' && target.Logged_By.toLowerCase() !== req.user.username) {
      return res.status(403).json({ error: 'You can only edit your own transactions' });
    }

    // Enforce store access
    const allowedT2 = await resolveStoreAccess(req.user);
    if (allowedT2 !== null) {
      const setT2 = new Set(allowedT2.map(s => s.toLowerCase()));
      if (!setT2.has(String(b.Store_Delivery).toLowerCase())) {
        return res.status(403).json({ error: `You don't have access to store: ${b.Store_Delivery}` });
      }
    }

    const gross = parseFloat(b.Transacted_Amount_Gross) || 0;
    const disc = parseFloat(b.Transacted_Discount_Net) || 0;
    const net = gross - disc;

    const row = [
      b.Date_Transacted,
      b.Customer_No,
      b.Name,
      b.Region,
      b.Area,
      b.Store_Delivery,
      b.Dept,
      b.Supplier,
      b.Booking_No,
      b.Deals || '',
      gross,
      disc,
      net,
      b.TRX_Number,
      target.Logged_By,
      req.user.username,
    ];

    await sheetsUpdateValues(`${TRX_SHEET}!A${rowIndex}:P${rowIndex}`, [row]);

    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.delete('/api/transactions/:rowIndex', requireAuth, async (req, res) => {
  try {
    if (req.user.role !== 'admin') return res.status(403).json({ error: 'Only admins can delete transactions' });
    const rowIndex = parseInt(req.params.rowIndex);
    const meta = await sheetsGetMeta();
    const ts = meta.sheets.find(s => s.properties.title === TRX_SHEET);
    if (!ts) throw new Error('Transaction sheet not found');

    await sheetsBatchUpdate([{
      deleteDimension: {
        range: {
          sheetId: ts.properties.sheetId,
          dimension: 'ROWS',
          startIndex: rowIndex - 1,
          endIndex: rowIndex,
        },
      },
    }]);

    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ===== BOOKING UPDATE (joins Booking + Transactions) =====
function num(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  return parseFloat(String(v).replace(/[^0-9.\-]/g, '')) || 0;
}

// Build joined report data: every booking + aggregated transaction totals
async function computeBookingUpdate() {
  const [bookings, transactions] = await Promise.all([readBookings(), readTransactions()]);

  // Aggregate transactions per Booking#
  const trxByBooking = {};
  for (const t of transactions) {
    const key = String(t.Booking_No || '').trim().toLowerCase();
    if (!key) continue;
    if (!trxByBooking[key]) trxByBooking[key] = { net: 0, disc: 0, gross: 0, latestDate: '' };
    trxByBooking[key].net += num(t.Transacted_Amount_Net);
    trxByBooking[key].disc += num(t.Transacted_Discount_Net);
    trxByBooking[key].gross += num(t.Transacted_Amount_Gross);
    if (!trxByBooking[key].latestDate || t.Date_Transacted) {
      trxByBooking[key].latestDate = t.Date_Transacted || trxByBooking[key].latestDate;
    }
  }

  return bookings.map(b => {
    const key = String(b.Booking_No || '').trim().toLowerCase();
    const t = trxByBooking[key] || { net: 0, disc: 0, gross: 0, latestDate: '' };
    const booked = num(b.Total_Booked_Amount);
    const balance = booked - t.net;
    const pct = booked > 0 ? (t.net / booked) : 0;
    return {
      Date_Transacted: t.latestDate || '',
      Customer_No: b.Customer_No,
      Name: b.Customer_Name,
      Region: b.Region,
      Area: b.Area,
      Store_Delivery: b.Store_Delivery,
      Dept: b.Dept,
      Supplier: b.Supplier,
      Booking_No: b.Booking_No,
      Deals: b.Deals,
      Total_Booked_Amount: booked,
      Transacted_Amount_Net: t.net,
      Transacted_Discount_Net: t.disc,
      Transacted_Amount_Gross: t.gross,
      Balance_Amount_Tobe_Transacted: balance,
      Pct_Transacted: pct,
    };
  });
}

// API: get live read-only report data (no sheet write)
app.get('/api/booking-update', requireAuth, async (req, res) => {
  try {
    const data = await computeBookingUpdate();
    res.json(data);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/refresh-cache', (req, res) => { invalidateCache(); res.json({ ok: true }); });

// ===== JUSTIFICATIONS =====
const JUSTIFICATION_SHEET = 'Justification';

async function readJustifications() {
  try {
    const res = await sheetsGetValues(`${JUSTIFICATION_SHEET}!A:D`);
    const rows = res.values || [];
    if (rows.length <= 1) return [];
    return rows.slice(1).filter(r => r[0]).map((r, i) => ({
      rowIndex: i + 2,
      Customer_Name: r[0] || '',
      Justification: r[1] || '',
      Updated_By: r[2] || '',
      Updated_Date: r[3] || '',
    }));
  } catch (e) {
    console.error('[Justification] Sheet read error:', e.message);
    return [];
  }
}

app.get('/api/justifications', requireAuth, async (req, res) => {
  try { res.json(await readJustifications()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/justifications', requireAuth, async (req, res) => {
  try {
    const { Customer_Name, Justification } = req.body;
    if (!Customer_Name || !Justification) return res.status(400).json({ error: 'Customer name and justification required' });

    const existing = await readJustifications();
    const match = existing.find(j => j.Customer_Name.toLowerCase() === Customer_Name.toLowerCase());

    if (match) {
      await sheetsUpdateValues(`${JUSTIFICATION_SHEET}!A${match.rowIndex}:D${match.rowIndex}`, [[
        Customer_Name, Justification, req.user.username, todayMDY()
      ]]);
    } else {
      await sheetsAppendValues(`${JUSTIFICATION_SHEET}!A:D`, [[
        Customer_Name, Justification, req.user.username, todayMDY()
      ]]);
    }
    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ===== FRONTEND =====
const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CaMaNaVa Booking Management</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@zxing/library@0.21.0/umd/index.min.js"></script>
<style>
  :root{--bg:#f5f7fa;--card:#fff;--text:#1f2937;--muted:#6b7280;--border:#e5e7eb;--primary:#2563eb;--success:#10b981;--danger:#ef4444;--input:#fff}
  [data-theme="dark"]{--bg:#0f172a;--card:#1e293b;--text:#e2e8f0;--muted:#94a3b8;--border:#334155;--primary:#3b82f6;--input:#0f172a}
  body{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;transition:background .2s}
  .card{background:var(--card);color:var(--text);border:1px solid var(--border);border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.05)}
  .form-control,.form-select{background:var(--input);color:var(--text);border-color:var(--border)}
  .form-control:focus,.form-select:focus{background:var(--input);color:var(--text);box-shadow:0 0 0 .2rem rgba(37,99,235,.15)}
  .form-control:disabled,.form-control[readonly]{background:var(--bg);opacity:.85}
  .navbar-custom{background:var(--card);border-bottom:1px solid var(--border);padding:12px 20px}
  .nav-tabs{border-color:var(--border)}
  .nav-tabs .nav-link{color:var(--muted);border:none;padding:10px 18px}
  .nav-tabs .nav-link.active{color:var(--primary);background:transparent;border-bottom:2px solid var(--primary)}
  .nav-pills .nav-link{color:var(--muted)}
  .nav-pills .nav-link.active{background:var(--primary);color:#fff}
  .stat-card{padding:20px;border-radius:12px}
  .stat-card .label{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.5px}
  .stat-card .value{font-size:26px;font-weight:700;margin-top:6px}
  .toast-container{position:fixed;top:20px;right:20px;z-index:9999}
  .toast-msg{padding:12px 20px;border-radius:8px;margin-bottom:10px;color:#fff;min-width:280px;box-shadow:0 4px 12px rgba(0,0,0,.15);animation:slide .3s}
  .toast-msg.success{background:var(--success)}.toast-msg.error{background:var(--danger)}
  @keyframes slide{from{transform:translateX(100%)}to{transform:translateX(0)}}
  .table-wrap{max-height:520px;overflow:auto}
  .table thead{position:sticky;top:0;background:var(--card);z-index:10}
  .table{color:var(--text)}
  .table th{font-size:11px;text-transform:uppercase;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap}
  .table th:hover{color:var(--primary)}
  .table td{vertical-align:middle;font-size:14px}
  .spinner-overlay{position:fixed;inset:0;background:rgba(0,0,0,.4);display:none;align-items:center;justify-content:center;z-index:9998}
  .spinner-overlay.show{display:flex}
  .btn-primary{background:var(--primary);border-color:var(--primary)}
  .theme-toggle{cursor:pointer;padding:6px 12px;border-radius:6px;background:var(--bg);border:1px solid var(--border);color:var(--text)}
  .pagination button{margin:0 2px;padding:4px 10px;border-radius:6px;border:1px solid var(--border);background:var(--card);color:var(--text);font-size:13px}
  .pagination button.active{background:var(--primary);color:#fff;border-color:var(--primary)}
  .pagination button:disabled{opacity:.5;cursor:not-allowed}
  .action-btn{padding:2px 8px;font-size:12px;margin:0 2px}
  .form-label{font-size:13px;font-weight:500;margin-bottom:4px}
  .required-mark{color:var(--danger)}
  @media(max-width:768px){.stat-card .value{font-size:20px}.table td,.table th{font-size:12px}}
  .owner-badge{position:fixed;bottom:14px;right:14px;background:linear-gradient(135deg,#2563eb,#7c3aed);color:#fff;padding:6px 14px;border-radius:20px;font-size:11px;font-weight:600;letter-spacing:.5px;box-shadow:0 4px 12px rgba(37,99,235,.35);z-index:9997;user-select:none}
  .login-screen{position:fixed;inset:0;background:linear-gradient(135deg,#1e3a8a,#7c3aed);display:flex;align-items:center;justify-content:center;z-index:10000;padding:20px}
  .login-box{background:var(--card);border-radius:16px;padding:32px;width:100%;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.3)}
  .login-box h4{margin-bottom:4px;color:var(--text)}
  .login-box .sub{color:var(--muted);font-size:13px;margin-bottom:24px}
  .app-hidden{display:none}
  .user-chip{display:inline-flex;align-items:center;gap:8px;padding:4px 12px;background:var(--bg);border:1px solid var(--border);border-radius:20px;font-size:13px;color:var(--text)}
  .user-chip .role-tag{background:var(--primary);color:#fff;padding:1px 6px;border-radius:8px;font-size:10px;text-transform:uppercase}
  .user-chip .role-tag.admin{background:#dc2626}
  .scanner-modal{position:fixed;inset:0;background:rgba(0,0,0,.85);display:none;align-items:center;justify-content:center;z-index:10001;padding:20px}
  .scanner-modal.show{display:flex}
  .scanner-box{background:var(--card);border-radius:12px;padding:20px;width:100%;max-width:500px}
  .scanner-video{width:100%;border-radius:8px;background:#000;max-height:60vh}
  .scanner-status{text-align:center;margin-top:12px;color:var(--muted);font-size:13px}
  .calc-display{background:var(--bg);padding:8px 12px;border-radius:6px;border:1px solid var(--border);font-weight:600;color:var(--success)}
</style>
</head>
<body>

<div class="login-screen" id="loginScreen">
  <div class="login-box">
    <h4><i class="fas fa-calendar-check text-primary me-2"></i>CaMaNaVa Booking</h4>
    <div class="sub">Sign in to continue</div>
    <div class="mb-3"><label class="form-label">Username</label><input type="text" class="form-control" id="loginUser" autocomplete="username"></div>
    <div class="mb-3"><label class="form-label">Password</label><input type="password" class="form-control" id="loginPass" autocomplete="current-password"></div>
    <button class="btn btn-primary w-100" id="btnLogin" onclick="doLogin()"><i class="fas fa-sign-in-alt me-1"></i>Login</button>
    <div id="loginError" class="text-danger small mt-2" style="display:none"></div>
  </div>
</div>

<div id="appContainer" class="app-hidden">

<div class="navbar-custom d-flex justify-content-between align-items-center flex-wrap gap-2">
  <h5 class="m-0"><i class="fas fa-calendar-check me-2 text-primary"></i>CaMaNaVa Booking Management</h5>
  <div class="d-flex align-items-center gap-2">
    <span class="user-chip"><i class="fas fa-user"></i><span id="userName"></span><span class="role-tag" id="userRole"></span></span>
    <button class="theme-toggle" onclick="toggleTheme()"><i class="fas fa-moon" id="themeIcon"></i></button>
    <button class="theme-toggle" onclick="doLogout()" title="Logout"><i class="fas fa-sign-out-alt"></i></button>
  </div>
</div>

<div class="container-fluid p-3 p-md-4">
  <ul class="nav nav-tabs mb-3">
    <li class="nav-item"><a class="nav-link active" data-bs-toggle="tab" href="#summaryTab">📊 Booking Summary</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#addTab">➕ Add Booking</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#trxTab">📦 Booking Transaction</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#updateTab" onclick="loadBookingUpdate()">📈 Booking Update</a></li>
  </ul>

  <div class="tab-content">

    <!-- SUMMARY TAB -->
    <div class="tab-pane fade show active" id="summaryTab">
      <div class="row g-3 mb-3">
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Total Bookings</div><div class="value" id="statBookings">0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Total Amount</div><div class="value" id="statAmount">₱0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Customers</div><div class="value" id="statCustomers">0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Suppliers</div><div class="value" id="statSuppliers">0</div></div></div>
      </div>

      <div class="card p-3 mb-3">
        <div class="row g-2 align-items-end">
          <div class="col-6 col-md-3"><label class="form-label">Area</label><select class="form-select form-select-sm" id="filterArea"><option value="">All</option></select></div>
          <div class="col-6 col-md-3"><label class="form-label">Store</label><select class="form-select form-select-sm" id="filterStore"><option value="">All</option></select></div>
          <div class="col-6 col-md-2"><label class="form-label">From</label><input type="date" class="form-control form-control-sm" id="filterFrom"></div>
          <div class="col-6 col-md-2"><label class="form-label">To</label><input type="date" class="form-control form-control-sm" id="filterTo"></div>
          <div class="col-12 col-md-2"><button class="btn btn-sm btn-outline-secondary w-100" onclick="clearFilters()"><i class="fas fa-times me-1"></i>Clear</button></div>
        </div>
      </div>

      <div class="row g-3">
        <div class="col-md-6"><div class="card p-3"><h6>Amount per Supplier</h6><canvas id="chartSupplier" height="200"></canvas></div></div>
        <div class="col-md-6"><div class="card p-3"><h6>Amount per Store</h6><canvas id="chartStore" height="200"></canvas></div></div>
        <div class="col-12"><div class="card p-3"><h6>Booking Trend by Date</h6><canvas id="chartTrend" height="80"></canvas></div></div>
      </div>

      <div class="card p-3 mt-3">
        <h6 class="mb-2">Grouped Summary</h6>
        <ul class="nav nav-pills mb-3" id="groupTabs">
          <li class="nav-item"><a class="nav-link active" data-group="Supplier" href="#" onclick="event.preventDefault();setGroup('Supplier')">By Supplier</a></li>
          <li class="nav-item"><a class="nav-link" data-group="Customer_Name" href="#" onclick="event.preventDefault();setGroup('Customer_Name')">By Customer</a></li>
          <li class="nav-item"><a class="nav-link" data-group="Store_Delivery" href="#" onclick="event.preventDefault();setGroup('Store_Delivery')">By Store</a></li>
        </ul>
        <div class="table-responsive"><table class="table table-sm"><thead><tr><th>Group</th><th class="text-end">Bookings</th><th class="text-end">Total Amount</th></tr></thead><tbody id="groupBody"></tbody></table></div>
      </div>
    </div>

    <!-- ADD BOOKING TAB -->
    <div class="tab-pane fade" id="addTab">
      <div class="card p-3 p-md-4 mb-3">
        <h5 id="formTitle" class="mb-3">Add Booking</h5>
        <input type="hidden" id="editRowIndex">
        <div class="row g-3">
          <div class="col-md-4 col-sm-6"><label class="form-label">Date Booked</label><input type="text" class="form-control" id="Date_Booked" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Customer Name <span class="required-mark">*</span></label><input list="customerList" class="form-control" id="Customer_Name" autocomplete="off" placeholder="Type to search..."><datalist id="customerList"></datalist></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Customer No</label><input type="text" class="form-control" id="Customer_No" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Region <span class="required-mark">*</span></label><input type="text" class="form-control" id="Region" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Area <span class="required-mark">*</span></label><input type="text" class="form-control" id="Area" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Store Delivery <span class="required-mark">*</span></label><input list="storeList" class="form-control" id="Store_Delivery" autocomplete="off"><datalist id="storeList"></datalist></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Dept <span class="required-mark">*</span></label><input list="deptList" class="form-control" id="Dept" autocomplete="off" placeholder="Type to search..."><datalist id="deptList"></datalist></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Supplier <span class="required-mark">*</span></label><input list="supplierList" class="form-control" id="Supplier" autocomplete="off" placeholder="Type to search..."><datalist id="supplierList"></datalist></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Booking #</label><input type="text" class="form-control" id="Booking_No" placeholder="Optional"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Deals <span class="required-mark">*</span></label><input type="text" class="form-control" id="Deals"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Total Booked Amount <span class="required-mark">*</span></label><input type="number" step="0.01" class="form-control" id="Total_Booked_Amount" placeholder="0.00"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Remarks</label><input type="text" class="form-control" id="Remarks"></div>
        </div>
        <div class="mt-3 d-flex gap-2 flex-wrap">
          <button class="btn btn-primary" id="btnSave" onclick="saveBooking()"><i class="fas fa-save me-1"></i>Save</button>
          <button class="btn btn-secondary" onclick="clearForm()"><i class="fas fa-eraser me-1"></i>Clear</button>
        </div>
      </div>

      <div class="card p-3">
        <div class="d-flex justify-content-between align-items-center mb-2 flex-wrap gap-2">
          <h6 class="m-0">All Bookings</h6>
          <div class="d-flex gap-2 flex-wrap">
            <input type="text" class="form-control form-control-sm" id="searchTable" placeholder="Search..." style="width:180px">
            <button class="btn btn-sm btn-success" onclick="exportExcel()"><i class="fas fa-file-excel me-1"></i>Export Excel</button>
            <button class="btn btn-sm btn-outline-primary" onclick="loadBookings()"><i class="fas fa-sync"></i></button>
          </div>
        </div>
        <div class="row g-2 mb-3">
          <div class="col-6 col-md-3"><label class="form-label">Area</label><select class="form-select form-select-sm" id="tblFilterArea"><option value="">All Areas</option></select></div>
          <div class="col-6 col-md-3"><label class="form-label">Store</label><select class="form-select form-select-sm" id="tblFilterStore"><option value="">All Stores</option></select></div>
          <div class="col-12 col-md-2 d-flex align-items-end"><button class="btn btn-sm btn-outline-secondary w-100" onclick="clearTblFilters()"><i class="fas fa-times me-1"></i>Clear Filters</button></div>
        </div>
        <div class="table-wrap">
          <table class="table table-hover table-sm">
            <thead><tr>
              <th data-sort="Date_Booked">Date ⇅</th>
              <th data-sort="Booking_No">Booking# ⇅</th>
              <th data-sort="Customer_Name">Customer ⇅</th>
              <th data-sort="Store_Delivery">Store ⇅</th>
              <th data-sort="Supplier">Supplier ⇅</th>
              <th data-sort="Dept">Dept ⇅</th>
              <th data-sort="Total_Booked_Amount" class="text-end">Amount ⇅</th>
              <th data-sort="Logged_By">Logged By ⇅</th>
              <th>Actions</th>
            </tr></thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
        <div class="d-flex justify-content-between align-items-center mt-2 flex-wrap gap-2">
          <small class="text-muted" id="pageInfo"></small>
          <div class="pagination" id="pagination"></div>
        </div>
      </div>
    </div>

    <!-- BOOKING TRANSACTION TAB -->
    <div class="tab-pane fade" id="trxTab">
      <div class="card p-3 p-md-4 mb-3">
        <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
          <h5 id="trxFormTitle" class="m-0">Add Transaction</h5>
          <button class="btn btn-primary" onclick="openScanner()"><i class="fas fa-camera me-1"></i>Scan Barcode</button>
        </div>
        <input type="hidden" id="trxEditRowIndex">
        <div class="row g-3">
          <div class="col-md-4 col-sm-6">
            <label class="form-label">Booking # <span class="required-mark">*</span></label>
            <div class="input-group">
              <input type="text" class="form-control" id="trx_Booking_No" placeholder="Scan or type Booking#">
              <button class="btn btn-outline-primary" onclick="lookupBooking()"><i class="fas fa-search"></i></button>
            </div>
          </div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Date Transacted <span class="required-mark">*</span></label><input type="date" class="form-control" id="trx_Date_Transacted"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">TRX Number <span class="required-mark">*</span></label><input type="text" class="form-control" id="trx_TRX_Number" placeholder="Transaction reference"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Customer Name</label><input type="text" class="form-control" id="trx_Name" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Customer No</label><input type="text" class="form-control" id="trx_Customer_No" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Region</label><input type="text" class="form-control" id="trx_Region" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Area</label><input type="text" class="form-control" id="trx_Area" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Store Delivery</label><input type="text" class="form-control" id="trx_Store_Delivery" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Dept</label><input type="text" class="form-control" id="trx_Dept" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Supplier</label><input type="text" class="form-control" id="trx_Supplier" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Deals</label><input type="text" class="form-control" id="trx_Deals" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Total Booked Amount</label><input type="text" class="form-control" id="trx_Total_Booked" readonly></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Gross Amount <span class="required-mark">*</span></label><input type="number" step="0.01" class="form-control" id="trx_Gross" placeholder="0.00" oninput="calcNet()"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Discount</label><input type="number" step="0.01" class="form-control" id="trx_Discount" placeholder="0.00" oninput="calcNet()"></div>
          <div class="col-md-4 col-sm-6"><label class="form-label">Net Amount <span class="text-muted small">(auto)</span></label><input type="text" class="form-control calc-display" id="trx_Net" readonly value="0.00"></div>
        </div>
        <div class="mt-3 d-flex gap-2 flex-wrap">
          <button class="btn btn-success" id="trxBtnSave" onclick="saveTransaction()"><i class="fas fa-save me-1"></i>Save</button>
          <button class="btn btn-secondary" onclick="clearTrxForm()"><i class="fas fa-eraser me-1"></i>Clear</button>
        </div>
      </div>

      <div class="card p-3">
        <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
          <h6 class="m-0">All Transactions</h6>
          <div class="d-flex gap-2 flex-wrap">
            <input type="text" class="form-control form-control-sm" id="trxSearch" placeholder="Search..." style="width:180px">
            <button class="btn btn-sm btn-success" onclick="exportTrxExcel()"><i class="fas fa-file-excel me-1"></i>Export</button>
            <button class="btn btn-sm btn-outline-primary" onclick="loadTransactions()"><i class="fas fa-sync"></i></button>
          </div>
        </div>
        <div class="table-wrap">
          <table class="table table-hover table-sm">
            <thead><tr>
              <th>Date</th>
              <th>TRX #</th>
              <th>Booking #</th>
              <th>Customer</th>
              <th>Store</th>
              <th>Supplier</th>
              <th class="text-end">Gross</th>
              <th class="text-end">Discount</th>
              <th class="text-end">Net</th>
              <th>Logged By</th>
              <th>Actions</th>
            </tr></thead>
            <tbody id="trxTableBody"></tbody>
          </table>
        </div>
        <div class="d-flex justify-content-between align-items-center mt-2 flex-wrap gap-2">
          <small class="text-muted" id="trxPageInfo"></small>
          <div class="pagination" id="trxPagination"></div>
        </div>
      </div>
    </div>

    <!-- BOOKING UPDATE TAB -->
    <div class="tab-pane fade" id="updateTab">
      <div class="row g-3 mb-3">
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Total Booked</div><div class="value" id="updStatBooked">₱0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Transacted (Net)</div><div class="value" id="updStatNet">₱0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">Balance</div><div class="value" id="updStatBalance">₱0</div></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card"><div class="label">% Transacted</div><div class="value" id="updStatPct">0%</div></div></div>
      </div>

      <!-- PERFORMANCE BREAKDOWN -->
      <div class="card p-3 mb-3">
        <div class="d-flex justify-content-between align-items-center mb-2 flex-wrap gap-2">
          <h6 class="m-0">Performance Breakdown</h6>
          <div class="d-flex align-items-center gap-2 flex-wrap">
            <ul class="nav nav-pills" id="perfTabs">
              <li class="nav-item"><a class="nav-link active" data-perf="Area" href="#" onclick="event.preventDefault();setPerf('Area')">By Area</a></li>
              <li class="nav-item"><a class="nav-link" data-perf="Store_Delivery" href="#" onclick="event.preventDefault();setPerf('Store_Delivery')">By Store</a></li>
              <li class="nav-item"><a class="nav-link" data-perf="Supplier" href="#" onclick="event.preventDefault();setPerf('Supplier')">By Supplier</a></li>
              <li class="nav-item"><a class="nav-link" data-perf="Name" href="#" onclick="event.preventDefault();setPerf('Name')">By Customer</a></li>
              <li class="nav-item"><a class="nav-link" data-perf="Dept" href="#" onclick="event.preventDefault();setPerf('Dept')">By Dept</a></li>
            </ul>
            <button class="btn btn-sm btn-success" onclick="exportPerfExcel()"><i class="fas fa-file-excel me-1"></i>Export</button>
          </div>
        </div>
        <div class="table-responsive">
          <table class="table table-sm">
            <thead><tr>
              <th id="perfGroupCol" class="perf-sort" data-perf-sort="key" style="cursor:pointer;user-select:none">Area ⇅</th>
              <th class="text-end perf-sort" data-perf-sort="count" style="cursor:pointer;user-select:none">Bookings ⇅</th>
              <th class="text-end perf-sort" data-perf-sort="booked" style="cursor:pointer;user-select:none">Booked ⇅</th>
              <th class="text-end perf-sort" data-perf-sort="net" style="cursor:pointer;user-select:none">Transacted (Net) ⇅</th>
              <th class="text-end perf-sort" data-perf-sort="balance" style="cursor:pointer;user-select:none">Balance ⇅</th>
              <th class="text-end perf-sort" data-perf-sort="pct" style="cursor:pointer;user-select:none;min-width:160px">% Transacted ⇅</th>
              <th id="perfJustifyCol" style="display:none">Justification below <span id="perfJustifyPct">0</span>%</th>
              <th id="perfJustifyActCol" class="text-center" style="display:none;width:70px">Action</th>
            </tr></thead>
            <tbody id="perfBody"></tbody>
          </table>
        </div>
      </div>

      <div class="card p-3">
        <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
          <h6 class="m-0">Booking Update Report <span class="text-muted small">(read-only)</span></h6>
          <div class="d-flex gap-2 flex-wrap">
            <input type="text" class="form-control form-control-sm" id="updSearch" placeholder="Search..." style="width:180px">
            <button class="btn btn-sm btn-success" onclick="exportUpdateExcel()"><i class="fas fa-file-excel me-1"></i>Export Excel</button>
            <button class="btn btn-sm btn-outline-primary" onclick="loadBookingUpdate()" title="Refresh"><i class="fas fa-sync"></i></button>
          </div>
        </div>
        <div class="row g-2 mb-3">
          <div class="col-6 col-md-3"><label class="form-label">Area</label><select class="form-select form-select-sm" id="updFilterArea"><option value="">All Areas</option></select></div>
          <div class="col-6 col-md-3"><label class="form-label">Store</label><select class="form-select form-select-sm" id="updFilterStore"><option value="">All Stores</option></select></div>
          <div class="col-12 col-md-2 d-flex align-items-end"><button class="btn btn-sm btn-outline-secondary w-100" onclick="clearUpdFilters()"><i class="fas fa-times me-1"></i>Clear</button></div>
        </div>
        <div class="table-wrap">
          <table class="table table-hover table-sm">
            <thead><tr>
              <th>Date</th>
              <th>Booking #</th>
              <th>Customer</th>
              <th>Store</th>
              <th>Supplier</th>
              <th class="text-end">Booked</th>
              <th class="text-end">Net Trx</th>
              <th class="text-end">Balance</th>
              <th class="text-end">% Trx</th>
            </tr></thead>
            <tbody id="updTableBody"></tbody>
          </table>
        </div>
        <div class="d-flex justify-content-between align-items-center mt-2 flex-wrap gap-2">
          <small class="text-muted" id="updPageInfo"></small>
          <div class="pagination" id="updPagination"></div>
        </div>
      </div>
    </div>

  </div>
</div>
</div>

<div class="toast-container" id="toastContainer"></div>
<div class="spinner-overlay" id="spinner"><div class="spinner-border text-light"></div></div>
<div class="scanner-modal" id="scannerModal">
  <div class="scanner-box">
    <div class="d-flex justify-content-between align-items-center mb-2">
      <h6 class="m-0"><i class="fas fa-barcode me-2"></i>Scan Barcode</h6>
      <button class="btn-close" onclick="closeScanner()"></button>
    </div>
    <video class="scanner-video" id="scannerVideo" autoplay playsinline muted></video>
    <div class="scanner-status" id="scannerStatus">Point camera at barcode...</div>
    <div class="d-flex gap-2 mt-3">
      <button class="btn btn-sm btn-outline-secondary flex-fill" onclick="switchCamera()"><i class="fas fa-sync me-1"></i>Switch Camera</button>
      <button class="btn btn-sm btn-secondary flex-fill" onclick="closeScanner()">Cancel</button>
    </div>
  </div>
</div>
<div class="scanner-modal" id="justifyModal">
  <div class="scanner-box">
    <div class="d-flex justify-content-between align-items-center mb-2">
      <h6 class="m-0"><i class="fas fa-comment-dots me-2"></i>Justification</h6>
      <button class="btn-close" onclick="closeJustifyModal()"></button>
    </div>
    <div class="mb-3"><label class="form-label">Customer</label><input type="text" class="form-control" id="justifyCustomer" readonly></div>
    <div class="mb-3"><label class="form-label">% Transacted</label><input type="text" class="form-control" id="justifyPct" readonly></div>
    <div class="mb-3"><label class="form-label">Justification <span class="required-mark">*</span></label><textarea class="form-control" id="justifyText" rows="3" placeholder="Enter justification..."></textarea></div>
    <div class="d-flex gap-2">
      <button class="btn btn-primary flex-fill" onclick="saveJustification()"><i class="fas fa-save me-1"></i>Save</button>
      <button class="btn btn-secondary flex-fill" onclick="closeJustifyModal()">Cancel</button>
    </div>
  </div>
</div>
<div class="owner-badge">By Carl_M@17</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script>
let allBookings = [];
let lookups = { customers: [], stores: [], suppliers: [], depts: [] };
let currentPage = 1;
const PAGE_SIZE = 10;
let sortField = 'Date_Booked';
let sortAsc = false;
let currentGroup = 'Supplier';
let charts = {};
let currentUser = null;  // { username, fullName, role }

function getToken(){ return localStorage.getItem('authToken'); }
function setToken(t){ localStorage.setItem('authToken', t); }
function clearToken(){ localStorage.removeItem('authToken'); }

function toast(msg, type='success') {
  const el = document.createElement('div');
  el.className = 'toast-msg ' + type;
  el.textContent = msg;
  document.getElementById('toastContainer').appendChild(el);
  setTimeout(() => el.remove(), 3500);
}
function showSpinner(s){ document.getElementById('spinner').classList.toggle('show', s); }

function toggleTheme(){
  const cur = document.documentElement.getAttribute('data-theme');
  const next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  document.getElementById('themeIcon').className = next === 'dark' ? 'fas fa-sun' : 'fas fa-moon';
  localStorage.setItem('theme', next);
  if (allBookings.length) renderSummary();
}
if (localStorage.getItem('theme') === 'dark') toggleTheme();

function todayStr(){
  const d = new Date();
  return (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear();
}
function fmtPeso(n){
  // Handle values that may come back as "1,234.56" string from Google Sheets
  let num;
  if (typeof n === 'number') num = n;
  else if (n === null || n === undefined || n === '') num = 0;
  else num = parseFloat(String(n).replace(/[^0-9.\-]/g, '')) || 0;
  return '₱' + num.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function parseMDY(s){ if(!s) return null; const [m,d,y]=s.split('/').map(Number); return new Date(y,m-1,d); }

async function api(method, url, body){
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  const t = getToken();
  if (t) opts.headers['x-auth-token'] = t;
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(url, opts);
  const data = await res.json();
  if (res.status === 401) {
    clearToken();
    showLogin();
    throw new Error('Session expired, please log in again');
  }
  if (!res.ok) throw new Error(data.error || 'Request failed');
  return data;
}

function showLogin(){
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('appContainer').classList.add('app-hidden');
  currentUser = null;
}

function showApp(){
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appContainer').classList.remove('app-hidden');
  document.getElementById('userName').textContent = currentUser.fullName;
  const roleEl = document.getElementById('userRole');
  roleEl.textContent = currentUser.role;
  roleEl.className = 'role-tag ' + (currentUser.role === 'admin' ? 'admin' : '');
  // Show scope hint
  const scope = (currentUser.storeAccess || '').trim();
  if (currentUser.role !== 'admin' && scope && scope.toLowerCase() !== 'all'){
    document.getElementById('userName').title = 'Scope: ' + scope;
    if (!document.getElementById('userScope')){
      const span = document.createElement('span');
      span.id = 'userScope';
      span.style.cssText = 'font-size:11px;color:var(--muted);margin-left:4px';
      span.textContent = '· ' + scope;
      document.getElementById('userName').parentNode.insertBefore(span, document.getElementById('userRole'));
    } else {
      document.getElementById('userScope').textContent = '· ' + scope;
    }
  }
}

async function doLogin(){
  const btn = document.getElementById('btnLogin');
  const err = document.getElementById('loginError');
  err.style.display = 'none';
  const u = document.getElementById('loginUser').value.trim();
  const p = document.getElementById('loginPass').value;
  if (!u || !p){ err.textContent = 'Enter username and password'; err.style.display = 'block'; return; }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Signing in...';
  try {
    const res = await fetch('/api/login', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({username: u, password: p}) });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Login failed');
    setToken(data.token);
    currentUser = { username: data.username, fullName: data.fullName, role: data.role };
    showApp();
    document.getElementById('loginPass').value = '';
    await initApp();
  } catch(e){
    err.textContent = e.message;
    err.style.display = 'block';
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-sign-in-alt me-1"></i>Login';
  }
}

async function doLogout(){
  if (!confirm('Log out now?')) return;
  try { await api('POST', '/api/logout'); } catch(e) {}
  clearToken();
  document.getElementById('loginUser').value = '';
  showLogin();
}

// Enter key submits login
document.getElementById('loginPass').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
document.getElementById('loginUser').addEventListener('keydown', e => { if (e.key === 'Enter') document.getElementById('loginPass').focus(); });

async function loadLookups(){
  try {
    lookups = await api('GET', '/api/lookups');
    document.getElementById('customerList').innerHTML = lookups.customers.map(c => '<option value="'+c.name+'">').join('');
    document.getElementById('storeList').innerHTML = lookups.stores.map(s => '<option value="'+s.storeName+'">').join('');
    document.getElementById('supplierList').innerHTML = lookups.suppliers.map(s => '<option value="'+s+'">').join('');
    document.getElementById('deptList').innerHTML = lookups.depts.map(d => '<option value="'+d+'">').join('');

    const fs = document.getElementById('filterStore');
    fs.innerHTML = '<option value="">All</option>' + lookups.stores.map(s => '<option>'+s.storeName+'</option>').join('');
  } catch(e){ toast(e.message, 'error'); }
}

// AUTO-FILL: Customer selected
document.getElementById('Customer_Name').addEventListener('change', e => {
  const c = lookups.customers.find(x => x.name.toLowerCase() === e.target.value.toLowerCase());
  if (c) {
    document.getElementById('Customer_No').value = c.no;
    document.getElementById('Region').value = c.region;
    // Match store via Store ID (most reliable since customer.storeName is "105-VAL" but ListOfStore.storeName is "Valenzuela")
    const st = lookups.stores.find(s => s.storeId === c.storeId);
    if (st) {
      document.getElementById('Store_Delivery').value = st.storeName;
      document.getElementById('Area').value = st.area;
      document.getElementById('Region').value = st.region;
    } else {
      // Fallback to customer's own storeName if no match
      document.getElementById('Store_Delivery').value = c.storeName;
      document.getElementById('Area').value = '';
    }
  } else {
    document.getElementById('Customer_No').value = '';
    document.getElementById('Region').value = '';
    document.getElementById('Area').value = '';
    document.getElementById('Store_Delivery').value = '';
  }
});

// AUTO-FILL: Store changed (user override) -> update Area + Region
document.getElementById('Store_Delivery').addEventListener('change', e => {
  const st = lookups.stores.find(s => s.storeName === e.target.value);
  if (st) {
    document.getElementById('Area').value = st.area;
    document.getElementById('Region').value = st.region;
  }
});

function clearForm(){
  ['Customer_Name','Customer_No','Region','Area','Store_Delivery','Dept','Supplier','Booking_No','Deals','Total_Booked_Amount','Remarks'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('Date_Booked').value = todayStr();
  document.getElementById('editRowIndex').value = '';
  document.getElementById('formTitle').textContent = 'Add Booking';
  document.getElementById('btnSave').innerHTML = '<i class="fas fa-save me-1"></i>Save';
}

async function saveBooking(){
  const btn = document.getElementById('btnSave');
  const payload = {
    Date_Booked: document.getElementById('Date_Booked').value,
    Customer_Name: document.getElementById('Customer_Name').value,
    Customer_No: document.getElementById('Customer_No').value,
    Region: document.getElementById('Region').value,
    Area: document.getElementById('Area').value,
    Store_Delivery: document.getElementById('Store_Delivery').value,
    Dept: document.getElementById('Dept').value,
    Supplier: document.getElementById('Supplier').value,
    Booking_No: document.getElementById('Booking_No').value,
    Deals: document.getElementById('Deals').value,
    Total_Booked_Amount: document.getElementById('Total_Booked_Amount').value,
    Remarks: document.getElementById('Remarks').value,
  };

  const required = ['Customer_Name','Customer_No','Region','Area','Store_Delivery','Dept','Supplier','Deals','Total_Booked_Amount'];
  for (const f of required){
    if (!payload[f]){ toast('Please fill: ' + f.replace(/_/g,' '), 'error'); return; }
  }
  if (isNaN(parseFloat(payload.Total_Booked_Amount))){ toast('Amount must be a number', 'error'); return; }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Saving...';
  showSpinner(true);

  try {
    const editIdx = document.getElementById('editRowIndex').value;
    if (editIdx) {
      await api('PUT', '/api/bookings/' + editIdx, payload);
      toast('Booking updated');
    } else {
      await api('POST', '/api/bookings', payload);
      toast('Booking saved');
    }
    clearForm();
    await loadBookings();
  } catch(e){ toast(e.message, 'error'); }
  finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-save me-1"></i>Save';
    showSpinner(false);
  }
}

function editBooking(rowIndex){
  const b = allBookings.find(x => x.rowIndex === rowIndex);
  if (!b) return;
  document.getElementById('editRowIndex').value = b.rowIndex;
  document.getElementById('Date_Booked').value = b.Date_Booked;
  document.getElementById('Customer_Name').value = b.Customer_Name;
  document.getElementById('Customer_No').value = b.Customer_No;
  document.getElementById('Region').value = b.Region;
  document.getElementById('Area').value = b.Area;
  document.getElementById('Store_Delivery').value = b.Store_Delivery;
  document.getElementById('Dept').value = b.Dept;
  document.getElementById('Supplier').value = b.Supplier;
  document.getElementById('Booking_No').value = b.Booking_No;
  document.getElementById('Deals').value = b.Deals;
  document.getElementById('Total_Booked_Amount').value = b.Total_Booked_Amount;
  document.getElementById('Remarks').value = b.Remarks;
  document.getElementById('formTitle').textContent = 'Edit Booking #' + b.Booking_No;
  document.getElementById('btnSave').innerHTML = '<i class="fas fa-save me-1"></i>Update';
  document.querySelector('a[href="#addTab"]').click();
  window.scrollTo({top:0, behavior:'smooth'});
}

async function deleteBooking(rowIndex, bookingNo){
  if (!confirm('Delete booking #' + bookingNo + '? This cannot be undone.')) return;
  showSpinner(true);
  try {
    await api('DELETE', '/api/bookings/' + rowIndex);
    toast('Booking deleted');
    await loadBookings();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

async function loadBookings(){
  showSpinner(true);
  try {
    allBookings = await api('GET', '/api/bookings');
    renderTable();
    renderSummary();
    populateAreaFilter();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

function populateAreaFilter(){
  const areas = [...new Set(allBookings.map(b => b.Area).filter(Boolean))].sort();
  document.getElementById('filterArea').innerHTML = '<option value="">All</option>' + areas.map(a => '<option>'+a+'</option>').join('');
  // Same areas for table filter
  const tblArea = document.getElementById('tblFilterArea');
  const curArea = tblArea.value;
  tblArea.innerHTML = '<option value="">All Areas</option>' + areas.map(a => '<option>'+a+'</option>').join('');
  tblArea.value = curArea;

  // Table store filter — pull from actual booking data (handles both old and new naming)
  const stores = [...new Set(allBookings.map(b => b.Store_Delivery).filter(Boolean))].sort();
  const tblStore = document.getElementById('tblFilterStore');
  const curStore = tblStore.value;
  tblStore.innerHTML = '<option value="">All Stores</option>' + stores.map(s => '<option>'+s+'</option>').join('');
  tblStore.value = curStore;
}

function getTableFiltered(){
  const q = document.getElementById('searchTable').value.toLowerCase();
  const fArea = document.getElementById('tblFilterArea').value;
  const fStore = (document.getElementById('tblFilterStore').value || '').toLowerCase();
  let data = allBookings;
  if (fArea) data = data.filter(b => b.Area === fArea);
  if (fStore) {
    // Smart match: handles both clean names ("Valenzuela") and old format ("105-VAL")
    data = data.filter(b => {
      const s = String(b.Store_Delivery || '').toLowerCase();
      return s === fStore || s.includes(fStore) || fStore.includes(s);
    });
  }
  if (q) data = data.filter(b => Object.values(b).some(v => String(v).toLowerCase().includes(q)));
  data = [...data].sort((a,b) => {
    let va = a[sortField], vb = b[sortField];
    if (sortField === 'Total_Booked_Amount'){ va = parseFloat(va)||0; vb = parseFloat(vb)||0; }
    else if (sortField === 'Date_Booked'){ va = parseMDY(va)||0; vb = parseMDY(vb)||0; }
    else { va = String(va).toLowerCase(); vb = String(vb).toLowerCase(); }
    if (va < vb) return sortAsc ? -1 : 1;
    if (va > vb) return sortAsc ? 1 : -1;
    return 0;
  });
  return data;
}

function clearTblFilters(){
  document.getElementById('tblFilterArea').value = '';
  document.getElementById('tblFilterStore').value = '';
  document.getElementById('searchTable').value = '';
  currentPage = 1;
  renderTable();
}

function exportExcel(){
  const data = getTableFiltered();
  if (!data.length){ toast('No data to export', 'error'); return; }

  // Build rows with headers exactly as on the Booking sheet (A-L, exclude M and N)
  const headers = ['Date_Booked','Customer_No','Customer_Name','Region','Area','Store_Delivery','Dept','Supplier','Booking_No','Deals','Total_Booked_Amount','Remarks'];
  const rows = [headers, ...data.map(b => headers.map(h => {
    if (h === 'Total_Booked_Amount') return num(b[h]);
    return b[h] || '';
  }))];

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Style header row: dark green fill, white bold font
  const headerStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: '006100' } },
    font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    }
  };
  for (let c = 0; c < headers.length; c++){
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) ws[addr].s = headerStyle;
  }

  // Auto-size columns
  ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 2, 15) }));

  // Currency format on Total_Booked_Amount column (K = index 10)
  const currencyFmt = '#,##0.00';
  for (let r = 1; r <= data.length; r++){
    const addr = XLSX.utils.encode_cell({ r, c: 10 });
    if (ws[addr]) ws[addr].z = currencyFmt;
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BookingData');

  const d = new Date();
  const stamp = \`\${d.getFullYear()}-\${String(d.getMonth()+1).padStart(2,'0')}-\${String(d.getDate()).padStart(2,'0')}_\${String(d.getHours()).padStart(2,'0')}-\${String(d.getMinutes()).padStart(2,'0')}\`;
  XLSX.writeFile(wb, \`BookingData_\${stamp}.xlsx\`);
  toast(\`Exported \${data.length} rows\`);
}

function renderTable(){
  const data = getTableFiltered();
  const totalPages = Math.max(1, Math.ceil(data.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageData = data.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('tableBody');
  if (!pageData.length){
    tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-4">No bookings found</td></tr>';
  } else {
    tbody.innerHTML = pageData.map(b => {
      const isOwner = currentUser && b.Logged_By.toLowerCase() === currentUser.username;
      const isAdmin = currentUser && currentUser.role === 'admin';
      const canEdit = isAdmin || isOwner;
      const canDelete = isAdmin;
      const editorTip = b.Last_Edited_By ? \` title="Last edited by \${b.Last_Edited_By}"\` : '';
      return \`
      <tr>
        <td>\${b.Date_Booked}</td>
        <td><strong>\${b.Booking_No || '—'}</strong></td>
        <td>\${b.Customer_Name}</td>
        <td>\${b.Store_Delivery}</td>
        <td>\${b.Supplier}</td>
        <td>\${b.Dept}</td>
        <td class="text-end">\${fmtPeso(b.Total_Booked_Amount)}</td>
        <td\${editorTip}>\${b.Logged_By}\${b.Last_Edited_By ? ' <i class="fas fa-pen text-muted" style="font-size:10px"></i>' : ''}</td>
        <td class="text-nowrap">
          \${canEdit ? \`<button class="btn btn-sm btn-outline-primary action-btn" onclick="editBooking(\${b.rowIndex})"><i class="fas fa-edit"></i></button>\` : ''}
          \${canDelete ? \`<button class="btn btn-sm btn-outline-danger action-btn" onclick="deleteBooking(\${b.rowIndex}, '\${b.Booking_No || b.rowIndex}')"><i class="fas fa-trash"></i></button>\` : ''}
          \${!canEdit && !canDelete ? '<span class="text-muted small">—</span>' : ''}
        </td>
      </tr>\`;
    }).join('');
  }

  document.getElementById('pageInfo').textContent = \`Showing \${start+1}-\${Math.min(start+PAGE_SIZE, data.length)} of \${data.length}\`;

  const pag = document.getElementById('pagination');
  let html = \`<button \${currentPage===1?'disabled':''} onclick="goPage(\${currentPage-1})">‹</button>\`;
  for (let i=1; i<=totalPages; i++){
    if (i===1 || i===totalPages || Math.abs(i-currentPage)<=1){
      html += \`<button class="\${i===currentPage?'active':''}" onclick="goPage(\${i})">\${i}</button>\`;
    } else if (Math.abs(i-currentPage)===2){
      html += '<span style="padding:0 4px">…</span>';
    }
  }
  html += \`<button \${currentPage===totalPages?'disabled':''} onclick="goPage(\${currentPage+1})">›</button>\`;
  pag.innerHTML = html;
}

function goPage(p){ currentPage = p; renderTable(); }

document.querySelectorAll('th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const f = th.dataset.sort;
    if (sortField === f) sortAsc = !sortAsc;
    else { sortField = f; sortAsc = true; }
    renderTable();
  });
});
document.getElementById('searchTable').addEventListener('input', () => { currentPage = 1; renderTable(); });
document.getElementById('tblFilterArea').addEventListener('change', () => { currentPage = 1; renderTable(); });
document.getElementById('tblFilterStore').addEventListener('change', () => { currentPage = 1; renderTable(); });

// SUMMARY
function getSummaryFiltered(){
  const area = document.getElementById('filterArea').value;
  const store = document.getElementById('filterStore').value;
  const from = document.getElementById('filterFrom').value;
  const to = document.getElementById('filterTo').value;

  // Parse YYYY-MM-DD as LOCAL date (not UTC) to avoid timezone shift
  const parseInputDate = (s, endOfDay) => {
    if (!s) return null;
    const [y, m, d] = s.split('-').map(Number);
    return endOfDay ? new Date(y, m-1, d, 23, 59, 59) : new Date(y, m-1, d, 0, 0, 0);
  };
  const fromDate = parseInputDate(from, false);
  const toDate = parseInputDate(to, true);

  return allBookings.filter(b => {
    if (area && b.Area !== area) return false;
    if (store && b.Store_Delivery !== store) return false;
    if (fromDate || toDate){
      const d = parseMDY(b.Date_Booked);
      if (!d) return false;
      if (fromDate && d < fromDate) return false;
      if (toDate && d > toDate) return false;
    }
    return true;
  });
}

function clearFilters(){
  document.getElementById('filterArea').value = '';
  document.getElementById('filterStore').value = '';
  document.getElementById('filterFrom').value = '';
  document.getElementById('filterTo').value = '';
  renderSummary();
}
['filterArea','filterStore','filterFrom','filterTo'].forEach(id => {
  document.getElementById(id).addEventListener('change', renderSummary);
});

function setGroup(g){
  currentGroup = g;
  document.querySelectorAll('#groupTabs .nav-link').forEach(a => {
    a.classList.toggle('active', a.dataset.group === g);
  });
  renderGroupTable();
}

function renderSummary(){
  const data = getSummaryFiltered();
  document.getElementById('statBookings').textContent = data.length;
  document.getElementById('statAmount').textContent = fmtPeso(data.reduce((s,b)=>s+(parseFloat(b.Total_Booked_Amount)||0),0));
  document.getElementById('statCustomers').textContent = new Set(data.map(b=>b.Customer_Name)).size;
  document.getElementById('statSuppliers').textContent = new Set(data.map(b=>b.Supplier)).size;
  renderGroupTable();
  renderCharts(data);
}

function renderGroupTable(){
  const data = getSummaryFiltered();
  const groups = {};
  data.forEach(b => {
    const k = b[currentGroup] || '(blank)';
    if (!groups[k]) groups[k] = { count: 0, total: 0 };
    groups[k].count++;
    groups[k].total += parseFloat(b.Total_Booked_Amount) || 0;
  });
  const rows = Object.entries(groups).sort((a,b) => b[1].total - a[1].total);
  document.getElementById('groupBody').innerHTML = rows.length
    ? rows.map(([k,v]) => \`<tr><td>\${k}</td><td class="text-end">\${v.count}</td><td class="text-end">\${fmtPeso(v.total)}</td></tr>\`).join('')
    : '<tr><td colspan="3" class="text-center text-muted py-3">No data</td></tr>';
}

function renderCharts(data){
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const textColor = isDark ? '#e2e8f0' : '#1f2937';
  Chart.defaults.color = textColor;

  // Per Supplier
  const sup = {};
  data.forEach(b => { sup[b.Supplier] = (sup[b.Supplier]||0) + (parseFloat(b.Total_Booked_Amount)||0); });
  const supEntries = Object.entries(sup).sort((a,b)=>b[1]-a[1]).slice(0,10);

  // Per Store
  const st = {};
  data.forEach(b => { st[b.Store_Delivery] = (st[b.Store_Delivery]||0) + (parseFloat(b.Total_Booked_Amount)||0); });
  const stEntries = Object.entries(st).sort((a,b)=>b[1]-a[1]).slice(0,10);

  // Trend by date
  const tr = {};
  data.forEach(b => { tr[b.Date_Booked] = (tr[b.Date_Booked]||0) + (parseFloat(b.Total_Booked_Amount)||0); });
  const trEntries = Object.entries(tr).sort((a,b) => parseMDY(a[0]) - parseMDY(b[0]));

  Object.values(charts).forEach(c => c && c.destroy());

  charts.supplier = new Chart(document.getElementById('chartSupplier'), {
    type: 'bar',
    data: { labels: supEntries.map(e=>e[0]), datasets: [{ label: 'Amount', data: supEntries.map(e=>e[1]), backgroundColor: '#3b82f6' }] },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, responsive: true }
  });
  charts.store = new Chart(document.getElementById('chartStore'), {
    type: 'bar',
    data: { labels: stEntries.map(e=>e[0]), datasets: [{ label: 'Amount', data: stEntries.map(e=>e[1]), backgroundColor: '#10b981' }] },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, responsive: true }
  });
  charts.trend = new Chart(document.getElementById('chartTrend'), {
    type: 'line',
    data: { labels: trEntries.map(e=>e[0]), datasets: [{ label: 'Amount', data: trEntries.map(e=>e[1]), borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,0.1)', tension: 0.3, fill: true }] },
    options: { responsive: true, plugins: { legend: { display: false } } }
  });
}

// ===== BOOKING TRANSACTION =====
let allTransactions = [];
let trxCurrentPage = 1;
let trxScanner = null;
let trxCameraIndex = 0;
let trxCameras = [];

function todayISO(){
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function calcNet(){
  const g = parseFloat(document.getElementById('trx_Gross').value) || 0;
  const d = parseFloat(document.getElementById('trx_Discount').value) || 0;
  document.getElementById('trx_Net').value = (g - d).toFixed(2);
}

async function lookupBooking(){
  const bn = document.getElementById('trx_Booking_No').value.trim();
  if (!bn){ toast('Enter or scan a Booking#', 'error'); return; }
  showSpinner(true);
  try {
    const b = await api('GET', '/api/booking-lookup/' + encodeURIComponent(bn));
    document.getElementById('trx_Name').value = b.Customer_Name;
    document.getElementById('trx_Customer_No').value = b.Customer_No;
    document.getElementById('trx_Region').value = b.Region;
    document.getElementById('trx_Area').value = b.Area;
    document.getElementById('trx_Store_Delivery').value = b.Store_Delivery;
    document.getElementById('trx_Dept').value = b.Dept;
    document.getElementById('trx_Supplier').value = b.Supplier;
    document.getElementById('trx_Deals').value = b.Deals;
    document.getElementById('trx_Total_Booked').value = fmtPeso(b.Total_Booked_Amount);
    toast('Booking found: ' + b.Customer_Name);
  } catch(e){
    toast(e.message, 'error');
    // Clear auto-filled fields if lookup failed
    ['trx_Name','trx_Customer_No','trx_Region','trx_Area','trx_Store_Delivery','trx_Dept','trx_Supplier','trx_Deals','trx_Total_Booked'].forEach(id => document.getElementById(id).value = '');
  } finally { showSpinner(false); }
}

document.getElementById('trx_Booking_No').addEventListener('keydown', e => {
  if (e.key === 'Enter'){ e.preventDefault(); lookupBooking(); }
});

function clearTrxForm(){
  ['trx_Booking_No','trx_Date_Transacted','trx_TRX_Number','trx_Name','trx_Customer_No','trx_Region','trx_Area','trx_Store_Delivery','trx_Dept','trx_Supplier','trx_Deals','trx_Total_Booked','trx_Gross','trx_Discount'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('trx_Net').value = '0.00';
  document.getElementById('trx_Date_Transacted').value = todayISO();
  document.getElementById('trxEditRowIndex').value = '';
  document.getElementById('trxFormTitle').textContent = 'Add Transaction';
  document.getElementById('trxBtnSave').innerHTML = '<i class="fas fa-save me-1"></i>Save';
}

async function saveTransaction(){
  const btn = document.getElementById('trxBtnSave');
  const dateInput = document.getElementById('trx_Date_Transacted').value;
  // Convert YYYY-MM-DD to M/D/YYYY to match sheet format
  let dateOut = '';
  if (dateInput){
    const [y,m,d] = dateInput.split('-').map(Number);
    dateOut = m + '/' + d + '/' + y;
  }

  const payload = {
    Date_Transacted: dateOut,
    Customer_No: document.getElementById('trx_Customer_No').value,
    Name: document.getElementById('trx_Name').value,
    Region: document.getElementById('trx_Region').value,
    Area: document.getElementById('trx_Area').value,
    Store_Delivery: document.getElementById('trx_Store_Delivery').value,
    Dept: document.getElementById('trx_Dept').value,
    Supplier: document.getElementById('trx_Supplier').value,
    Booking_No: document.getElementById('trx_Booking_No').value,
    Deals: document.getElementById('trx_Deals').value,
    Transacted_Amount_Gross: document.getElementById('trx_Gross').value,
    Transacted_Discount_Net: document.getElementById('trx_Discount').value || '0',
    Transacted_Amount_Net: document.getElementById('trx_Net').value,
    TRX_Number: document.getElementById('trx_TRX_Number').value,
  };

  if (!payload.Booking_No){ toast('Booking# is required', 'error'); return; }
  if (!payload.Date_Transacted){ toast('Date is required', 'error'); return; }
  if (!payload.TRX_Number){ toast('TRX Number is required', 'error'); return; }
  if (!payload.Customer_No){ toast('Look up the Booking# first', 'error'); return; }
  if (!payload.Transacted_Amount_Gross){ toast('Gross amount is required', 'error'); return; }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Saving...';
  showSpinner(true);
  try {
    const editIdx = document.getElementById('trxEditRowIndex').value;
    if (editIdx){
      await api('PUT', '/api/transactions/' + editIdx, payload);
      toast('Transaction updated');
    } else {
      await api('POST', '/api/transactions', payload);
      toast('Transaction saved');
    }
    clearTrxForm();
    await loadTransactions();
  } catch(e){ toast(e.message, 'error'); }
  finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-save me-1"></i>Save';
    showSpinner(false);
  }
}

async function loadTransactions(){
  showSpinner(true);
  try {
    allTransactions = await api('GET', '/api/transactions');
    renderTrxTable();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

function getTrxFiltered(){
  const q = document.getElementById('trxSearch').value.toLowerCase();
  let data = allTransactions;
  if (q) data = data.filter(t => Object.values(t).some(v => String(v).toLowerCase().includes(q)));
  return [...data].sort((a,b) => {
    const da = parseMDY(a.Date_Transacted) || 0;
    const db = parseMDY(b.Date_Transacted) || 0;
    return db - da;
  });
}

function renderTrxTable(){
  const data = getTrxFiltered();
  const totalPages = Math.max(1, Math.ceil(data.length / PAGE_SIZE));
  if (trxCurrentPage > totalPages) trxCurrentPage = totalPages;
  const start = (trxCurrentPage - 1) * PAGE_SIZE;
  const pageData = data.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('trxTableBody');
  if (!pageData.length){
    tbody.innerHTML = '<tr><td colspan="11" class="text-center text-muted py-4">No transactions found</td></tr>';
  } else {
    tbody.innerHTML = pageData.map(t => {
      const isOwner = currentUser && t.Logged_By.toLowerCase() === currentUser.username;
      const isAdmin = currentUser && currentUser.role === 'admin';
      const canEdit = isAdmin || isOwner;
      const canDelete = isAdmin;
      const editTip = t.Last_Edited_By ? \` title="Last edited by \${t.Last_Edited_By}"\` : '';
      return \`
      <tr>
        <td>\${t.Date_Transacted}</td>
        <td><strong>\${t.TRX_Number}</strong></td>
        <td>\${t.Booking_No}</td>
        <td>\${t.Name}</td>
        <td>\${t.Store_Delivery}</td>
        <td>\${t.Supplier}</td>
        <td class="text-end">\${fmtPeso(t.Transacted_Amount_Gross)}</td>
        <td class="text-end">\${fmtPeso(t.Transacted_Discount_Net)}</td>
        <td class="text-end"><strong>\${fmtPeso(t.Transacted_Amount_Net)}</strong></td>
        <td\${editTip}>\${t.Logged_By}\${t.Last_Edited_By ? ' <i class="fas fa-pen text-muted" style="font-size:10px"></i>' : ''}</td>
        <td class="text-nowrap">
          \${canEdit ? \`<button class="btn btn-sm btn-outline-primary action-btn" onclick="editTransaction(\${t.rowIndex})"><i class="fas fa-edit"></i></button>\` : ''}
          \${canDelete ? \`<button class="btn btn-sm btn-outline-danger action-btn" onclick="deleteTransaction(\${t.rowIndex}, '\${t.TRX_Number}')"><i class="fas fa-trash"></i></button>\` : ''}
          \${!canEdit && !canDelete ? '<span class="text-muted small">—</span>' : ''}
        </td>
      </tr>\`;
    }).join('');
  }

  document.getElementById('trxPageInfo').textContent = \`Showing \${start+1}-\${Math.min(start+PAGE_SIZE, data.length)} of \${data.length}\`;

  const pag = document.getElementById('trxPagination');
  let html = \`<button \${trxCurrentPage===1?'disabled':''} onclick="goTrxPage(\${trxCurrentPage-1})">‹</button>\`;
  for (let i=1; i<=totalPages; i++){
    if (i===1 || i===totalPages || Math.abs(i-trxCurrentPage)<=1){
      html += \`<button class="\${i===trxCurrentPage?'active':''}" onclick="goTrxPage(\${i})">\${i}</button>\`;
    } else if (Math.abs(i-trxCurrentPage)===2){
      html += '<span style="padding:0 4px">…</span>';
    }
  }
  html += \`<button \${trxCurrentPage===totalPages?'disabled':''} onclick="goTrxPage(\${trxCurrentPage+1})">›</button>\`;
  pag.innerHTML = html;
}

function goTrxPage(p){ trxCurrentPage = p; renderTrxTable(); }

document.getElementById('trxSearch').addEventListener('input', () => { trxCurrentPage = 1; renderTrxTable(); });

function editTransaction(rowIndex){
  const t = allTransactions.find(x => x.rowIndex === rowIndex);
  if (!t) return;
  document.getElementById('trxEditRowIndex').value = t.rowIndex;
  document.getElementById('trx_Booking_No').value = t.Booking_No;
  // Convert M/D/YYYY back to YYYY-MM-DD for input[type=date]
  if (t.Date_Transacted){
    const [m,d,y] = t.Date_Transacted.split('/');
    document.getElementById('trx_Date_Transacted').value = y + '-' + String(m).padStart(2,'0') + '-' + String(d).padStart(2,'0');
  }
  document.getElementById('trx_TRX_Number').value = t.TRX_Number;
  document.getElementById('trx_Name').value = t.Name;
  document.getElementById('trx_Customer_No').value = t.Customer_No;
  document.getElementById('trx_Region').value = t.Region;
  document.getElementById('trx_Area').value = t.Area;
  document.getElementById('trx_Store_Delivery').value = t.Store_Delivery;
  document.getElementById('trx_Dept').value = t.Dept;
  document.getElementById('trx_Supplier').value = t.Supplier;
  document.getElementById('trx_Deals').value = t.Deals;
  document.getElementById('trx_Gross').value = t.Transacted_Amount_Gross;
  document.getElementById('trx_Discount').value = t.Transacted_Discount_Net;
  // Look up the original booking to show its Total_Booked_Amount (background, won't block)
  document.getElementById('trx_Total_Booked').value = '...';
  api('GET', '/api/booking-lookup/' + encodeURIComponent(t.Booking_No))
    .then(b => { document.getElementById('trx_Total_Booked').value = fmtPeso(b.Total_Booked_Amount); })
    .catch(() => { document.getElementById('trx_Total_Booked').value = ''; });
  calcNet();
  document.getElementById('trxFormTitle').textContent = 'Edit Transaction ' + t.TRX_Number;
  document.getElementById('trxBtnSave').innerHTML = '<i class="fas fa-save me-1"></i>Update';
  window.scrollTo({top:0, behavior:'smooth'});
}

async function deleteTransaction(rowIndex, trxNum){
  if (!confirm('Delete transaction ' + trxNum + '? This cannot be undone.')) return;
  showSpinner(true);
  try {
    await api('DELETE', '/api/transactions/' + rowIndex);
    toast('Transaction deleted');
    await loadTransactions();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

function exportTrxExcel(){
  const data = getTrxFiltered();
  if (!data.length){ toast('No data to export', 'error'); return; }

  const headers = ['Date_Transacted','CUSTOMER_NO','NAME','REGION','AREA','STORE_DELIVERY','DEPT','SUPPLIER','BOOKING #','DEALS','Transacted_Amount_Gross','Transacted_Discount_Net','Transacted_Amount_Net','TRX_Number'];
  const fields = ['Date_Transacted','Customer_No','Name','Region','Area','Store_Delivery','Dept','Supplier','Booking_No','Deals','Transacted_Amount_Gross','Transacted_Discount_Net','Transacted_Amount_Net','TRX_Number'];
  const rows = [headers, ...data.map(t => fields.map(f => {
    if (['Transacted_Amount_Gross','Transacted_Discount_Net','Transacted_Amount_Net'].includes(f)) return num(t[f]);
    return t[f] || '';
  }))];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const headerStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: '006100' } },
    font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    }
  };
  for (let c = 0; c < headers.length; c++){
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 2, 15) }));
  for (let r = 1; r <= data.length; r++){
    [10,11,12].forEach(c => {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]) ws[addr].z = '#,##0.00';
    });
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BookingTransaction');

  const d = new Date();
  const stamp = \`\${d.getFullYear()}-\${String(d.getMonth()+1).padStart(2,'0')}-\${String(d.getDate()).padStart(2,'0')}_\${String(d.getHours()).padStart(2,'0')}-\${String(d.getMinutes()).padStart(2,'0')}\`;
  XLSX.writeFile(wb, \`BookingTransaction_\${stamp}.xlsx\`);
  toast(\`Exported \${data.length} rows\`);
}

// ===== BARCODE SCANNER =====
async function openScanner(){
  document.getElementById('scannerModal').classList.add('show');
  document.getElementById('scannerStatus').textContent = 'Starting camera...';
  try {
    if (!window.ZXing){ throw new Error('Scanner library failed to load'); }
    trxScanner = new ZXing.BrowserMultiFormatReader();
    const devices = await trxScanner.listVideoInputDevices();
    if (!devices.length) throw new Error('No camera found');
    trxCameras = devices;
    // Prefer rear camera on phones
    const rear = devices.findIndex(d => /back|rear|environment/i.test(d.label));
    trxCameraIndex = rear >= 0 ? rear : devices.length - 1;
    await startScanner();
  } catch(e){
    document.getElementById('scannerStatus').textContent = 'Error: ' + e.message;
    toast(e.message, 'error');
  }
}

async function startScanner(){
  if (!trxScanner || !trxCameras.length) return;
  const deviceId = trxCameras[trxCameraIndex].deviceId;
  document.getElementById('scannerStatus').textContent = 'Point camera at barcode...';
  try {
    await trxScanner.decodeFromVideoDevice(deviceId, 'scannerVideo', (result, err) => {
      if (result){
        const code = result.getText();
        document.getElementById('trx_Booking_No').value = code;
        document.getElementById('scannerStatus').textContent = '✓ Scanned: ' + code;
        closeScanner();
        lookupBooking();
      }
    });
  } catch(e){
    document.getElementById('scannerStatus').textContent = 'Camera error: ' + e.message;
  }
}

async function switchCamera(){
  if (trxCameras.length < 2){ toast('Only one camera available', 'error'); return; }
  trxCameraIndex = (trxCameraIndex + 1) % trxCameras.length;
  if (trxScanner){ trxScanner.reset(); }
  await startScanner();
}

function closeScanner(){
  if (trxScanner){ try { trxScanner.reset(); } catch(e){} trxScanner = null; }
  document.getElementById('scannerModal').classList.remove('show');
}

let currentPerf = 'Area';
let perfSortField = 'booked';
let perfSortAsc = false;
let justifications = {};
let overallPct = 0;

async function loadJustifications(){
  try {
    const data = await api('GET', '/api/justifications');
    justifications = {};
    data.forEach(j => { justifications[j.Customer_Name.toLowerCase()] = j; });
  } catch(e){ console.error('Failed to load justifications:', e); }
}

function openJustifyModal(customerName, pct){
  const existing = justifications[customerName.toLowerCase()];
  document.getElementById('justifyCustomer').value = customerName;
  document.getElementById('justifyPct').value = pct.toFixed(1) + '%';
  document.getElementById('justifyText').value = existing ? existing.Justification : '';
  document.getElementById('justifyModal').classList.add('show');
  document.getElementById('justifyText').focus();
}

function closeJustifyModal(){
  document.getElementById('justifyModal').classList.remove('show');
}

async function saveJustification(){
  const name = document.getElementById('justifyCustomer').value;
  const text = document.getElementById('justifyText').value.trim();
  if (!text){ toast('Please enter a justification', 'error'); return; }
  showSpinner(true);
  try {
    await api('POST', '/api/justifications', { Customer_Name: name, Justification: text });
    toast('Justification saved');
    closeJustifyModal();
    await loadJustifications();
    renderPerformance();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

function setPerf(field){
  currentPerf = field;
  document.querySelectorAll('#perfTabs .nav-link').forEach(a => {
    a.classList.toggle('active', a.dataset.perf === field);
  });
  const labels = { Area: 'Area', Store_Delivery: 'Store', Supplier: 'Supplier', Name: 'Customer', Dept: 'Dept' };
  document.getElementById('perfGroupCol').textContent = (labels[field] || field) + ' ⇅';
  renderPerformance();
}

function renderPerformance(){
  const data = getUpdFiltered();
  const isCustomer = currentPerf === 'Name';
  const groups = {};
  data.forEach(r => {
    const key = r[currentPerf] || '(blank)';
    if (!groups[key]) groups[key] = { count: 0, booked: 0, net: 0 };
    groups[key].count++;
    groups[key].booked += num(r.Total_Booked_Amount);
    groups[key].net += num(r.Transacted_Amount_Net);
  });

  let rows = Object.entries(groups)
    .map(([k, v]) => ({ key: k, ...v, balance: v.booked - v.net, pct: v.booked > 0 ? (v.net / v.booked * 100) : 0 }));

  // Sort by current sort field
  rows.sort((a, b) => {
    let va = a[perfSortField], vb = b[perfSortField];
    if (typeof va === 'string'){ va = va.toLowerCase(); vb = String(vb).toLowerCase(); }
    if (va < vb) return perfSortAsc ? -1 : 1;
    if (va > vb) return perfSortAsc ? 1 : -1;
    return 0;
  });

  // Show/hide justification columns
  document.getElementById('perfJustifyCol').style.display = isCustomer ? '' : 'none';
  document.getElementById('perfJustifyActCol').style.display = isCustomer ? '' : 'none';
  if (isCustomer) document.getElementById('perfJustifyPct').textContent = overallPct.toFixed(1);
  const colSpan = isCustomer ? 8 : 6;

  const tbody = document.getElementById('perfBody');
  if (!rows.length){
    tbody.innerHTML = '<tr><td colspan="' + colSpan + '" class="text-center text-muted py-3">No data</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    const pctCapped = Math.min(r.pct, 100);
    const pctColor = r.pct >= 100 ? 'var(--success)' : r.pct >= 50 ? 'var(--primary)' : 'var(--danger)';
    let justifyCell = '';
    if (isCustomer) {
      const j = justifications[r.key.toLowerCase()];
      const safeKey = r.key.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
      if (r.pct < overallPct) {
        const btnLabel = j ? '<i class="fas fa-edit"></i>' : '<i class="fas fa-plus"></i> Add';
        const jText = j ? j.Justification.replace(/</g,'&lt;') : '<span class="text-muted">—</span>';
        justifyCell = '<td style="white-space:normal;word-break:break-word;max-width:300px">' + jText + '</td>'
          + '<td class="text-center"><button class="btn btn-sm btn-outline-primary action-btn" data-cust="' + safeKey + '" data-pct="' + r.pct + '" onclick="openJustifyModal(this.dataset.cust, parseFloat(this.dataset.pct))">' + btnLabel + '</button></td>';
      } else {
        justifyCell = '<td class="text-muted small">—</td><td></td>';
      }
    }
    return \`
      <tr>
        <td><strong>\${r.key}</strong></td>
        <td class="text-end">\${r.count}</td>
        <td class="text-end">\${fmtPeso(r.booked)}</td>
        <td class="text-end">\${fmtPeso(r.net)}</td>
        <td class="text-end">\${fmtPeso(r.balance)}</td>
        <td class="text-end">
          <div style="display:flex;align-items:center;gap:8px;justify-content:flex-end">
            <div style="flex:1;max-width:100px;height:8px;background:var(--border);border-radius:4px;overflow:hidden">
              <div style="height:100%;width:\${pctCapped}%;background:\${pctColor};transition:width 0.3s"></div>
            </div>
            <span style="color:\${pctColor};font-weight:600;min-width:50px">\${r.pct.toFixed(1)}%</span>
          </div>
        </td>
        \${justifyCell}
      </tr>
    \`;
  }).join('');
}

// Header click → sort toggle
document.querySelectorAll('.perf-sort').forEach(th => {
  th.addEventListener('click', () => {
    const f = th.dataset.perfSort;
    if (perfSortField === f) perfSortAsc = !perfSortAsc;
    else { perfSortField = f; perfSortAsc = f === 'key'; }  // strings default asc, numbers default desc
    renderPerformance();
  });
});

// ===== BOOKING UPDATE (read-only report) =====
let allUpdate = [];
let updCurrentPage = 1;

async function loadBookingUpdate(){
  showSpinner(true);
  try {
    const [data] = await Promise.all([api('GET', '/api/booking-update'), loadJustifications()]);
    allUpdate = data;
    populateUpdFilters();
    renderUpdate();
  } catch(e){ toast(e.message, 'error'); }
  finally { showSpinner(false); }
}

function populateUpdFilters(){
  const areas = [...new Set(allUpdate.map(r => r.Area).filter(Boolean))].sort();
  const stores = [...new Set(allUpdate.map(r => r.Store_Delivery).filter(Boolean))].sort();
  const aSel = document.getElementById('updFilterArea');
  const sSel = document.getElementById('updFilterStore');
  const curA = aSel.value, curS = sSel.value;
  aSel.innerHTML = '<option value="">All Areas</option>' + areas.map(a => '<option>'+a+'</option>').join('');
  sSel.innerHTML = '<option value="">All Stores</option>' + stores.map(s => '<option>'+s+'</option>').join('');
  aSel.value = curA; sSel.value = curS;
}

function getUpdFiltered(){
  const q = document.getElementById('updSearch').value.toLowerCase();
  const fA = document.getElementById('updFilterArea').value;
  const fS = document.getElementById('updFilterStore').value;
  let data = allUpdate;
  if (fA) data = data.filter(r => r.Area === fA);
  if (fS) data = data.filter(r => r.Store_Delivery === fS);
  if (q) data = data.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(q)));
  return data;
}

function clearUpdFilters(){
  document.getElementById('updFilterArea').value = '';
  document.getElementById('updFilterStore').value = '';
  document.getElementById('updSearch').value = '';
  updCurrentPage = 1;
  renderUpdate();
}

function fmtPct(v){
  const n = num(v) * 100;
  return n.toFixed(1) + '%';
}
function num(v){
  if (typeof v === 'number') return v;
  if (!v) return 0;
  return parseFloat(String(v).replace(/[^0-9.\-]/g, '')) || 0;
}

function renderUpdate(){
  const data = getUpdFiltered();

  // Stat cards
  const booked = data.reduce((s,r) => s + num(r.Total_Booked_Amount), 0);
  const net = data.reduce((s,r) => s + num(r.Transacted_Amount_Net), 0);
  const balance = booked - net;
  const pct = booked > 0 ? (net / booked * 100) : 0;
  overallPct = pct;
  document.getElementById('updStatBooked').textContent = fmtPeso(booked);
  document.getElementById('updStatNet').textContent = fmtPeso(net);
  document.getElementById('updStatBalance').textContent = fmtPeso(balance);
  document.getElementById('updStatPct').textContent = pct.toFixed(1) + '%';

  // Render performance breakdown alongside main report
  renderPerformance();

  // Table pagination
  const totalPages = Math.max(1, Math.ceil(data.length / PAGE_SIZE));
  if (updCurrentPage > totalPages) updCurrentPage = totalPages;
  const start = (updCurrentPage - 1) * PAGE_SIZE;
  const pageData = data.slice(start, start + PAGE_SIZE);

  const tbody = document.getElementById('updTableBody');
  if (!pageData.length){
    tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-4">No data</td></tr>';
  } else {
    tbody.innerHTML = pageData.map(r => {
      const pctVal = num(r.Pct_Transacted) * 100;
      const pctColor = pctVal >= 100 ? 'var(--success)' : pctVal >= 50 ? 'var(--primary)' : 'var(--danger)';
      return \`
        <tr>
          <td>\${r.Date_Transacted || '—'}</td>
          <td><strong>\${r.Booking_No || '—'}</strong></td>
          <td>\${r.Name}</td>
          <td>\${r.Store_Delivery}</td>
          <td>\${r.Supplier}</td>
          <td class="text-end">\${fmtPeso(r.Total_Booked_Amount)}</td>
          <td class="text-end">\${fmtPeso(r.Transacted_Amount_Net)}</td>
          <td class="text-end">\${fmtPeso(r.Balance_Amount_Tobe_Transacted)}</td>
          <td class="text-end" style="color:\${pctColor};font-weight:600">\${pctVal.toFixed(1)}%</td>
        </tr>
      \`;
    }).join('');
  }

  document.getElementById('updPageInfo').textContent = \`Showing \${data.length ? start+1 : 0}-\${Math.min(start+PAGE_SIZE, data.length)} of \${data.length}\`;

  const pag = document.getElementById('updPagination');
  let html = \`<button \${updCurrentPage===1?'disabled':''} onclick="goUpdPage(\${updCurrentPage-1})">‹</button>\`;
  for (let i=1; i<=totalPages; i++){
    if (i===1 || i===totalPages || Math.abs(i-updCurrentPage)<=1){
      html += \`<button class="\${i===updCurrentPage?'active':''}" onclick="goUpdPage(\${i})">\${i}</button>\`;
    } else if (Math.abs(i-updCurrentPage)===2){
      html += '<span style="padding:0 4px">…</span>';
    }
  }
  html += \`<button \${updCurrentPage===totalPages?'disabled':''} onclick="goUpdPage(\${updCurrentPage+1})">›</button>\`;
  pag.innerHTML = html;
}

function goUpdPage(p){ updCurrentPage = p; renderUpdate(); }

document.getElementById('updSearch').addEventListener('input', () => { updCurrentPage = 1; renderUpdate(); });
document.getElementById('updFilterArea').addEventListener('change', () => { updCurrentPage = 1; renderUpdate(); });
document.getElementById('updFilterStore').addEventListener('change', () => { updCurrentPage = 1; renderUpdate(); });

function exportPerfExcel(){
  const data = getUpdFiltered();
  const isCustomer = currentPerf === 'Name';
  const groups = {};
  data.forEach(r => {
    const key = r[currentPerf] || '(blank)';
    if (!groups[key]) groups[key] = { count: 0, booked: 0, net: 0 };
    groups[key].count++;
    groups[key].booked += num(r.Total_Booked_Amount);
    groups[key].net += num(r.Transacted_Amount_Net);
  });
  const rows = Object.entries(groups).map(([k, v]) => ({
    key: k, ...v, balance: v.booked - v.net, pct: v.booked > 0 ? (v.net / v.booked * 100) : 0
  }));
  if (!rows.length){ toast('No data to export', 'error'); return; }

  const labels = { Area: 'Area', Store_Delivery: 'Store', Supplier: 'Supplier', Name: 'Customer', Dept: 'Dept' };
  const groupLabel = labels[currentPerf] || currentPerf;
  let headers = [groupLabel, 'Bookings', 'Booked', 'Transacted (Net)', 'Balance', '% Transacted'];
  if (isCustomer) headers.push('Justification below ' + overallPct.toFixed(1) + '%');

  const xlRows = [headers, ...rows.map(r => {
    const row = [r.key, r.count, r.booked, r.net, r.balance, r.pct / 100];
    if (isCustomer) {
      const j = justifications[r.key.toLowerCase()];
      row.push(r.pct < overallPct ? (j ? j.Justification : '') : '');
    }
    return row;
  })];

  const ws = XLSX.utils.aoa_to_sheet(xlRows);
  ws['!cols'] = headers.map((h, i) => ({ wch: i === 0 ? 30 : isCustomer && i === headers.length - 1 ? 40 : 18 }));
  for (let r = 1; r <= rows.length; r++){
    const pctAddr = XLSX.utils.encode_cell({ r, c: 5 });
    if (ws[pctAddr]) ws[pctAddr].z = '0.0%';
    [2,3,4].forEach(c => {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]) ws[addr].z = '#,##0.00';
    });
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Performance_' + groupLabel);
  const d = new Date();
  const stamp = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  XLSX.writeFile(wb, 'Performance_' + groupLabel + '_' + stamp + '.xlsx');
  toast('Exported ' + rows.length + ' rows');
}

function exportUpdateExcel(){
  const data = getUpdFiltered();
  if (!data.length){ toast('No data to export', 'error'); return; }

  const headers = ['Date_Transacted','CUSTOMER_NO','NAME','REGION','AREA','STORE_DELIVERY','DEPT','SUPPLIER','BOOKING #','DEALS','Total_Booked_Amount','Transacted_Amount_Net','Transacted_Discount_Net','Transacted_Amount_Gross','Balance_Amount_Tobe_Transacted','% Transacted'];
  const fields = ['Date_Transacted','Customer_No','Name','Region','Area','Store_Delivery','Dept','Supplier','Booking_No','Deals','Total_Booked_Amount','Transacted_Amount_Net','Transacted_Discount_Net','Transacted_Amount_Gross','Balance_Amount_Tobe_Transacted','Pct_Transacted'];
  const rows = [headers, ...data.map(r => fields.map(f => {
    if (['Total_Booked_Amount','Transacted_Amount_Net','Transacted_Discount_Net','Transacted_Amount_Gross','Balance_Amount_Tobe_Transacted','Pct_Transacted'].includes(f)) return num(r[f]);
    return r[f] || '';
  }))];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const headerStyle = {
    fill: { patternType: 'solid', fgColor: { rgb: '006100' } },
    font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 11 },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    }
  };
  for (let c = 0; c < headers.length; c++){
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 2, 15) }));
  // Currency format for cols K-O (10-14), % format for col P (15)
  for (let r = 1; r <= data.length; r++){
    [10,11,12,13,14].forEach(c => {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]) ws[addr].z = '#,##0.00';
    });
    const pAddr = XLSX.utils.encode_cell({ r, c: 15 });
    if (ws[pAddr]) ws[pAddr].z = '0.0%';
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BookingUpdate');

  const d = new Date();
  const stamp = \`\${d.getFullYear()}-\${String(d.getMonth()+1).padStart(2,'0')}-\${String(d.getDate()).padStart(2,'0')}_\${String(d.getHours()).padStart(2,'0')}-\${String(d.getMinutes()).padStart(2,'0')}\`;
  XLSX.writeFile(wb, \`BookingUpdate_\${stamp}.xlsx\`);
  toast(\`Exported \${data.length} rows\`);
}

async function initApp(){
  document.getElementById('Date_Booked').value = todayStr();
  document.getElementById('trx_Date_Transacted').value = todayISO();
  await loadLookups();
  await loadBookings();
  await loadTransactions();
  // Auto-refresh every 30s for multi-user sync
  if (window._refreshInterval) clearInterval(window._refreshInterval);
  window._refreshInterval = setInterval(() => { loadBookings(); loadTransactions(); }, 30000);
}

// INIT — try auto-login if token exists
(async function(){
  const t = getToken();
  if (!t){ showLogin(); return; }
  try {
    const res = await fetch('/api/me', { headers: { 'x-auth-token': t } });
    if (!res.ok) throw new Error('expired');
    const me = await res.json();
    currentUser = me;
    showApp();
    await initApp();
  } catch(e){
    clearToken();
    showLogin();
  }
})();
</script>
</body>
</html>`;

app.get('/', (req, res) => res.send(HTML));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Booking app running on port ${PORT}`));
