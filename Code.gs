/**
 * DASHBOARD MANAGEMENT FEE - PT BKK
 * Sumber data: Google Spreadsheet, sheet "Data_Laporan"
 * Penyesuaian UI dan perhitungan Summary x2 dilakukan di Index.html.
 * Angka dasar di sini tetap asli agar detail outlet tidak berubah.
 */
const CONFIG = {
  SPREADSHEET_ID: '1VVc5kpn8yD4A_5LtnFHf-o_epconJTMU8S2RW2ud_Qw',
  SHEET_NAME: 'Data_Laporan',
  HEADER_ROW: 1
};

const FEE_FIELDS = [
  { key: 'sa', label: 'Fee SA', target: 'TARGET FEE SA', actual: 'ACTUAL FEE SA' },
  { key: 'vf', label: 'Fee VF', target: 'TARGET FEE VF', actual: 'ACTUAL FEE VF' },
  { key: 'digipos', label: 'Fee Digipos', target: 'TARGET FEE DIGIPOS', actual: 'ACTUAL FEE DIGIPOS' },
  { key: 'growth', label: 'Fee Growth', target: 'TARGET FEE GROWTH', actual: 'ACTUAL FEE GROWTH' }
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Management Fee Dashboard | PT BKK')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(sessionToken) {
  requireSession_(sessionToken);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  if (values.length <= CONFIG.HEADER_ROW) {
    return { rows: [], summary: emptySummary_(), filters: { clusters: [], taps: [], salesforces: [] }, updatedAt: new Date().toISOString(), source: CONFIG.SHEET_NAME };
  }

  const headers = values[CONFIG.HEADER_ROW - 1].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h.toUpperCase()] = i; });

  const required = [
    'NAMA OUTLET', 'OUTLET ID', 'NO RS', 'KAB', 'KEC', 'SALESFORCE',
    'CLUSTER', 'TAP', 'TOTAL TARGET FEE', 'TOTAL ACTUAL FEE',
    ...FEE_FIELDS.flatMap(f => [f.target, f.actual])
  ];
  const missing = required.filter(h => idx[h] === undefined);
  if (missing.length) throw new Error('Kolom wajib tidak ditemukan: ' + missing.join(', '));

  const rows = [];
  for (let r = CONFIG.HEADER_ROW; r < values.length; r++) {
    const line = values[r];
    const outletName = clean_(line[idx['NAMA OUTLET']]);
    if (!outletName) continue;

    const item = {
      no: r - CONFIG.HEADER_ROW + 1,
      outletId: clean_(line[idx['OUTLET ID']]),
      noRs: clean_(line[idx['NO RS']]),
      namaOutlet: outletName,
      kab: clean_(line[idx['KAB']]),
      kec: clean_(line[idx['KEC']]),
      salesforce: clean_(line[idx['SALESFORCE']]),
      cluster: clean_(line[idx['CLUSTER']]),
      tap: clean_(line[idx['TAP']]),
      target: toNumber_(line[idx['TOTAL TARGET FEE']]),
      actual: toNumber_(line[idx['TOTAL ACTUAL FEE']]),
      fees: {}
    };
    FEE_FIELDS.forEach(f => {
      item.fees[f.key] = {
        target: toNumber_(line[idx[f.target.toUpperCase()]]),
        actual: toNumber_(line[idx[f.actual.toUpperCase()]])
      };
    });
    rows.push(item);
  }

  const summary = summarize_(rows);
  const unique = field => [...new Set(rows.map(x => x[field]).filter(Boolean))].sort();
  return {
    rows, summary,
    filters: { clusters: unique('cluster'), taps: unique('tap'), salesforces: unique('salesforce') },
    updatedAt: new Date().toISOString(),
    source: CONFIG.SHEET_NAME
  };
}

function summarize_(rows) {
  const total = emptySummary_();
  const byTap = {};
  const byCluster = {};
  rows.forEach(row => {
    total.outlets++;
    total.target += row.target;
    total.actual += row.actual;
    FEE_FIELDS.forEach(f => {
      total.fees[f.key].target += row.fees[f.key].target;
      total.fees[f.key].actual += row.fees[f.key].actual;
    });
    addGroup_(byTap, row.tap, row);
    addGroup_(byCluster, row.cluster, row);
  });
  total.achievement = total.target ? total.actual / total.target * 100 : 0;
  total.gap = total.actual - total.target;
  total.byTap = Object.keys(byTap).sort().map(k => finalizeGroup_(k, byTap[k]));
  total.byCluster = Object.keys(byCluster).sort().map(k => finalizeGroup_(k, byCluster[k]));
  total.byTap.sort((a, b) => b.achievement - a.achievement);
  total.byCluster.sort((a, b) => b.achievement - a.achievement);
  FEE_FIELDS.forEach(f => {
    const v = total.fees[f.key];
    v.achievement = v.target ? v.actual / v.target * 100 : 0;
    v.gap = v.actual - v.target;
  });
  return total;
}

function emptySummary_() {
  const fees = {};
  FEE_FIELDS.forEach(f => fees[f.key] = { label: f.label, target: 0, actual: 0, achievement: 0, gap: 0 });
  return { outlets: 0, target: 0, actual: 0, achievement: 0, gap: 0, fees, byTap: [], byCluster: [] };
}

function addGroup_(map, key, row) {
  const name = key || '(Kosong)';
  if (!map[name]) {
    map[name] = {
      outlets: 0, target: 0, actual: 0,
      fees: { sa: { target: 0, actual: 0 }, vf: { target: 0, actual: 0 },
              digipos: { target: 0, actual: 0 }, growth: { target: 0, actual: 0 } }
    };
  }
  const g = map[name];
  g.outlets++;
  g.target += row.target;
  g.actual += row.actual;
  FEE_FIELDS.forEach(f => {
    g.fees[f.key].target += row.fees[f.key].target;
    g.fees[f.key].actual += row.fees[f.key].actual;
  });
}

function finalizeGroup_(name, g) {
  return {
    name, outlets: g.outlets, target: g.target, actual: g.actual,
    achievement: g.target ? g.actual / g.target * 100 : 0,
    gap: g.actual - g.target,
    fees: g.fees
  };
}

function clean_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function toNumber_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined || v === '') return 0;
  let s = String(v).trim().replace(/\s/g, '');
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else if (s.includes(',')) s = s.replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}


/** LOGIN SERVER-SIDE — semua endpoint data wajib memeriksa sesi. */
const AUTH = { USER_PREFIX:'BKK_USER_', SESSION_SECONDS:21600, FAIL_LIMIT:5, LOCK_SECONDS:900 };

function authHash_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0'+(b & 255).toString(16)).slice(-2)).join('');
}
function passwordHash_(password, salt) {
  // SHA-256 iteratif dengan salt berbeda per akun; untuk keamanan lebih tinggi
  // di organisasi gunakan Google Workspace/Identity Provider.
  let digest = salt + ':' + password;
  for (let i=0; i<1500; i++) digest=authHash_(digest);
  return digest;
}
function userKey_(username) { return AUTH.USER_PREFIX + String(username).toLowerCase(); }
function readUser_(username) {
  const raw=PropertiesService.getScriptProperties().getProperty(userKey_(username));
  return raw ? JSON.parse(raw) : null;
}
function saveUser_(u) { PropertiesService.getScriptProperties().setProperty(userKey_(u.username),JSON.stringify(u)); }
function safeUser_(u) { return {username:u.username,role:u.role,active:u.active!==false}; }
function validUsername_(name) { return /^[a-z0-9._-]{3,40}$/.test(name); }
function validPassword_(p) { return typeof p==='string' && p.length>=12 && p.length<=128; }

/**
 * LANGKAH PERTAMA: Project Settings > Script properties:
 * BKK_INITIAL_ADMIN_USERNAME = admin
 * BKK_INITIAL_ADMIN_PASSWORD = password unik minimal 12 karakter
 * Setelah itu jalankan setupFirstAdmin() dari editor Apps Script.
 * Properti password awal otomatis dihapus setelah akun dibuat.
 */
function setupFirstAdmin() {
  const props=PropertiesService.getScriptProperties();
  const username=String(props.getProperty('BKK_INITIAL_ADMIN_USERNAME')||'').trim().toLowerCase();
  const password=props.getProperty('BKK_INITIAL_ADMIN_PASSWORD');
  if (!validUsername_(username) || !validPassword_(password)) throw new Error('Isi Script Properties BKK_INITIAL_ADMIN_USERNAME dan BKK_INITIAL_ADMIN_PASSWORD (minimal 12 karakter).');
  const all=props.getProperties();
  if (Object.keys(all).some(k=>k.startsWith(AUTH.USER_PREFIX))) throw new Error('Akun sudah ada. Gunakan manajemen user di dashboard.');
  const salt=Utilities.getUuid()+Utilities.getUuid();
  saveUser_({username,role:'admin',active:true,salt,hash:passwordHash_(password,salt)});
  props.deleteProperty('BKK_INITIAL_ADMIN_USERNAME');
  props.deleteProperty('BKK_INITIAL_ADMIN_PASSWORD');
  return 'Akun administrator berhasil dibuat. Silakan login.';
}
function loginDashboard(username,password) {
  username=String(username||'').trim().toLowerCase();
  if (!validUsername_(username) || typeof password!=='string' || password.length>128) throw new Error('Username atau password salah.');
  const cache=CacheService.getScriptCache();
  const attemptsKey='fail_'+authHash_(username);
  const entry=JSON.parse(cache.get(attemptsKey)||'{"count":0}');
  if(entry.count>=AUTH.FAIL_LIMIT) throw new Error('Terlalu banyak percobaan login. Coba lagi dalam 15 menit.');
  const user=readUser_(username);
  const correct=user && user.active!==false && passwordHash_(password,user.salt)===user.hash;
  if(!correct) {
    cache.put(attemptsKey,JSON.stringify({count:entry.count+1}),AUTH.LOCK_SECONDS);
    throw new Error('Username atau password salah.');
  }
  cache.remove(attemptsKey);
  const token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  cache.put('session_'+authHash_(token),JSON.stringify({username}),AUTH.SESSION_SECONDS);
  return {token,user:safeUser_(user),expiresInSeconds:AUTH.SESSION_SECONDS};
}
function requireSession_(token,role) {
  if(typeof token!=='string'|| !/^[a-f0-9]{64}$/.test(token)) throw new Error('Sesi tidak valid. Silakan login.');
  const session=CacheService.getScriptCache().get('session_'+authHash_(token));
  if(!session) throw new Error('Sesi berakhir. Silakan login kembali.');
  const u=readUser_(JSON.parse(session).username);
  if(!u ||u.active===false) throw new Error('Akses akun sudah dinonaktifkan.');
  if(role && u.role!==role) throw new Error('Hanya administrator yang memiliki akses.');
  return safeUser_(u);
}
function validateDashboardSession(token) {return requireSession_(token);}
function logoutDashboard(token) {
  if(typeof token==='string' && /^[a-f0-9]{64}$/.test(token)) CacheService.getScriptCache().remove('session_'+authHash_(token));
  return true;
}
function listDashboardUsers(token) {
  requireSession_(token,'admin');
  const props=PropertiesService.getScriptProperties().getProperties();
  return Object.keys(props).filter(k=>k.startsWith(AUTH.USER_PREFIX)).map(k=>safeUser_(JSON.parse(props[k]))).sort((a,b)=>a.username.localeCompare(b.username));
}
function saveDashboardUser(token,username,password,role) {
  requireSession_(token,'admin');
  username=String(username||'').trim().toLowerCase();
  if(!validUsername_(username)) throw new Error('Username harus 3–40 karakter: huruf kecil, angka, titik, garis bawah atau minus.');
  if(!validPassword_(password)) throw new Error('Password minimal 12 karakter, maksimal 128 karakter.');
  if(role!=='admin' && role!=='viewer') throw new Error('Role tidak valid.');
  const existing=readUser_(username);
  const current=requireSession_(token,'admin');
  if(username===current.username && role!=='admin') throw new Error('Tidak dapat menurunkan role administrator sendiri.');
  if(existing && existing.role==='admin' && role!=='admin') {
    const activeAdmins=listDashboardUsers(token).filter(x=>x.role==='admin'&&x.active);
    if(activeAdmins.length<=1 && existing.active!==false) throw new Error('Minimal satu administrator aktif harus dipertahankan.');
  }
  const salt=Utilities.getUuid()+Utilities.getUuid();
  saveUser_({username,role,active:existing?existing.active!==false:true,salt,hash:passwordHash_(password,salt)});
  return listDashboardUsers(token);
}
function setDashboardUserActive(token,username,active) {
  const current=requireSession_(token,'admin');
  username=String(username||'').toLowerCase();
  const u=readUser_(username);
  if(!u) throw new Error('Pengguna tidak ditemukan.');
  if(u.username===current.username && !active) throw new Error('Tidak dapat menonaktifkan akun sendiri.');
  if(u.role==='admin' && !active) {
    const all=listDashboardUsers(token);
    if(all.filter(x=>x.role==='admin'&&x.active).length<=1)throw new Error('Minimal satu administrator harus aktif.');
  }
  u.active=Boolean(active);saveUser_(u);
  return listDashboardUsers(token);
}
