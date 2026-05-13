// Booking Management Web App
// Single-file Express server + embedded frontend
// Backend: Google Sheets via service account
// Deploy on Railway

const express = require('express');
const { google } = require('googleapis');
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

const auth = new google.auth.JWT(
  SERVICE_ACCOUNT_EMAIL,
  null,
  PRIVATE_KEY,
  ['https://www.googleapis.com/auth/spreadsheets']
);
const sheets = google.sheets({ version: 'v4', auth });

// ===== CACHE FOR LOOKUPS (refreshed every 60s) =====
let cache = { data: null, ts: 0 };
const CACHE_MS = 60 * 1000;

async function getLookups() {
  if (cache.data && Date.now() - cache.ts < CACHE_MS) return cache.data;

  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID,
    ranges: [
      `${CUSTOMER_SHEET}!A:E`,
      `${STORE_SHEET}!A:D`,
      `${SUPPLIER_SHEET}!A:C`,
    ],
  });

  const custRows = res.data.valueRanges[0].values || [];
  const storeRows = res.data.valueRanges[1].values || [];
  const supRows = res.data.valueRanges[2].values || [];

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
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${MAIN_SHEET}!A:N`,
  });
  const rows = res.data.values || [];
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
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Users!A:D',
  });
  const rows = res.data.values || [];
  const users = rows.slice(1).filter(r => r[0]).map(r => ({
    username: String(r[0]).trim().toLowerCase(),
    password: String(r[1] || ''),
    fullName: r[2] || r[0],
    role: (r[3] || 'user').toLowerCase(),
  }));
  userCache.data = users;
  userCache.ts = Date.now();
  return users;
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
    sessions.set(token, { username: u.username, fullName: u.fullName, role: u.role, ts: Date.now() });
    res.json({ token, username: u.username, fullName: u.fullName, role: u.role });
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
  res.json({ username: u.username, fullName: u.fullName, role: u.role });
});

// ===== API ROUTES =====
app.get('/api/lookups', requireAuth, async (req, res) => {
  try { res.json(await getLookups()); }
  catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.get('/api/bookings', requireAuth, async (req, res) => {
  try { res.json(await readBookings()); }
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

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${MAIN_SHEET}!A:N`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] },
    });

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

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${MAIN_SHEET}!A${rowIndex}:N${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [row] },
    });
    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.delete('/api/bookings/:rowIndex', requireAuth, async (req, res) => {
  try {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Only admins can delete bookings' });
    }
    const rowIndex = parseInt(req.params.rowIndex);
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const ms = meta.data.sheets.find(s => s.properties.title === MAIN_SHEET);
    if (!ms) throw new Error('Main sheet not found');

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: ms.properties.sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex,
            },
          },
        }],
      },
    });
    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post('/api/refresh-cache', (req, res) => { invalidateCache(); res.json({ ok: true }); });

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
        <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
          <h6 class="m-0">All Bookings</h6>
          <div class="d-flex gap-2">
            <input type="text" class="form-control form-control-sm" id="searchTable" placeholder="Search..." style="width:200px">
            <button class="btn btn-sm btn-outline-primary" onclick="loadBookings()"><i class="fas fa-sync"></i></button>
          </div>
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

  </div>
</div>
</div>

<div class="toast-container" id="toastContainer"></div>
<div class="spinner-overlay" id="spinner"><div class="spinner-border text-light"></div></div>
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
function fmtPeso(n){ return '₱' + Number(n||0).toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2}); }
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
}

function getTableFiltered(){
  const q = document.getElementById('searchTable').value.toLowerCase();
  let data = allBookings;
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

// SUMMARY
function getSummaryFiltered(){
  const area = document.getElementById('filterArea').value;
  const store = document.getElementById('filterStore').value;
  const from = document.getElementById('filterFrom').value;
  const to = document.getElementById('filterTo').value;
  return allBookings.filter(b => {
    if (area && b.Area !== area) return false;
    if (store && b.Store_Delivery !== store) return false;
    if (from){
      const d = parseMDY(b.Date_Booked);
      if (d && d < new Date(from)) return false;
    }
    if (to){
      const d = parseMDY(b.Date_Booked);
      if (d && d > new Date(to)) return false;
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

async function initApp(){
  document.getElementById('Date_Booked').value = todayStr();
  await loadLookups();
  await loadBookings();
  // Auto-refresh every 30s for multi-user sync
  if (window._refreshInterval) clearInterval(window._refreshInterval);
  window._refreshInterval = setInterval(loadBookings, 30000);
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
