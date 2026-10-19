const appInsights = require('applicationinsights');
if (process.env.APPLICATIONINSIGHTS_CONNECTION_STRING) {
  appInsights.setup(process.env.APPLICATIONINSIGHTS_CONNECTION_STRING)
    .setAutoCollectRequests(true)
    .setAutoCollectConsole(true, true)
    .start();
}

const express = require('express');
const sql = require('mssql');
const crypto = require('crypto');
const path = require('path');
const { BlobServiceClient } = require('@azure/storage-blob');
const { DefaultAzureCredential } = require('@azure/identity');
const multer = require('multer');
const { parse } = require('csv-parse/sync');

const MAX_BULK_ROWS = 25; // keeps one upload under Azure's ~230s request timeout
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 1024 * 1024 } });

const app = express();
app.use(express.json());
// Serve ONLY the portal page. (express.static(__dirname) would also expose server.js and package.json.)
app.get(['/', '/index.html'], (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// Set LOGIC_APP_URL as an App Service setting, then delete the fallback below.
const LOGIC_APP_URL = process.env.LOGIC_APP_URL ||
  'https://prod-87.eastus.logic.azure.com:443/workflows/98725403991e4980a919359584b71fba/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=-c3m9l_EzxuxS2vQJxWYi8Y0RDDIsZkL4nZJAfPW0E8';
const TEMPLATE_NAME = 'AutoOnboard Enterprise Policies and Governance Framework.pdf';

// ---------- Database ----------
const config = {
  user: process.env.DB_USER || 'autodbuser2026',
  password: process.env.DB_PASS, // comes from the App Service setting / Key Vault reference
  server: process.env.DB_HOST || 'sqlserver-autoonboard-01.database.windows.net',
  database: process.env.DB_NAME || 'sqldb-autoonboard',
  options: { encrypt: true, trustServerCertificate: false }
};

let poolPromise;
function getPool() {
  if (!poolPromise) {
    poolPromise = new sql.ConnectionPool(config).connect().catch(err => {
      poolPromise = null;
      throw err;
    });
  }
  return poolPromise;
}

async function audit(name, email, department, step, status, detail) {
  try {
    const pool = await getPool();
    await pool.request()
      .input('Name', sql.VarChar(100), name)
      .input('Email', sql.VarChar(100), email)
      .input('Department', sql.VarChar(50), department || null)
      .input('Step', sql.VarChar(30), step)
      .input('Status', sql.VarChar(10), status)
      .input('Detail', sql.VarChar(500), detail ? String(detail).slice(0, 500) : null)
      .query('INSERT INTO OnboardingAudit (Name, Email, Department, Step, Status, Detail) VALUES (@Name, @Email, @Department, @Step, @Status, @Detail)');
  } catch (e) {
    console.error('Audit write failed:', e.message);
  }
}

// ---------- Microsoft Entra ID via Microsoft Graph ----------
const GRAPH = {
  tenant: process.env.GRAPH_TENANT_ID,
  clientId: process.env.GRAPH_CLIENT_ID,
  secret: process.env.GRAPH_CLIENT_SECRET,
  domain: process.env.GRAPH_DOMAIN
};

let GROUP_MAP = {
  'DevOps & Cloud Architecture': '9128d01e-9b9d-420f-bb15-304a3458523c',
  'Engineering': 'ac47b960-f2e9-45f5-87b0-59cc45da2833',
  'Human Resources': 'f9f5aef8-85ed-45a2-bb9a-ae2de0894182',
  'Finance': '0271d00f-0e2d-452e-b3cd-d7498251af0c',
  'Marketing': '33089e86-33eb-492d-8075-e54d7b2e0da7',
  'Operations': 'ad997236-39ba-4447-a974-eb65854cd2a7'
};
try {
  if (process.env.GROUP_MAP) GROUP_MAP = Object.assign({}, GROUP_MAP, JSON.parse(process.env.GROUP_MAP));
} catch (e) {
  console.error('GROUP_MAP setting is not valid JSON, using built-in map');
}

const PRIMARY_DEPARTMENTS = ['DevOps & Cloud Architecture', 'Engineering', 'Human Resources', 'Finance', 'Marketing', 'Operations'];
const OTHER_DEPARTMENTS = ['Legal', 'Quality Assurance', 'Product Management', 'Security'];
const ALL_DEPARTMENTS = PRIMARY_DEPARTMENTS.concat(OTHER_DEPARTMENTS);

// Lets a CSV say "HR" or "devops" instead of the full department name.
const DEPT_ALIASES = {
  'devops': 'DevOps & Cloud Architecture', 'devops & cloud architecture': 'DevOps & Cloud Architecture',
  'devops and cloud architecture': 'DevOps & Cloud Architecture', 'cloud': 'DevOps & Cloud Architecture',
  'engineering': 'Engineering',
  'hr': 'Human Resources', 'human resources': 'Human Resources',
  'finance': 'Finance', 'marketing': 'Marketing', 'operations': 'Operations', 'ops': 'Operations',
  'legal': 'Legal', 'qa': 'Quality Assurance', 'quality assurance': 'Quality Assurance',
  'product': 'Product Management', 'product management': 'Product Management', 'security': 'Security'
};
function normalizeDept(value) {
  const v = String(value || '').trim().toLowerCase();
  if (!v) return null;
  const hit = DEPT_ALIASES[v];
  if (hit) return hit;
  return ALL_DEPARTMENTS.find(k => k.toLowerCase() === v) || null;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

function makeNickname(name, email) {
  let nick = String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
  if (!nick) nick = String(email || '').split('@')[0].toLowerCase().replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
  return (nick || 'newuser').slice(0, 40);
}

function generateTempPassword() {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ', lower = 'abcdefghijkmnpqrstuvwxyz';
  const digits = '23456789', symbols = '!@#$%*?';
  const pick = s => s[crypto.randomInt(s.length)];
  const all = upper + lower + digits + symbols;
  const chars = [pick(upper), pick(lower), pick(digits), pick(symbols)];
  while (chars.length < 16) chars.push(pick(all));
  for (let i = chars.length - 1; i > 0; i--) {
    const j = crypto.randomInt(i + 1);
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join('');
}

// Cache the Graph token so a bulk upload doesn't request a new one per row.
let tokenCache = { value: null, expires: 0 };
async function getGraphToken() {
  if (tokenCache.value && Date.now() < tokenCache.expires) return tokenCache.value;
  const res = await fetch('https://login.microsoftonline.com/' + GRAPH.tenant + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: GRAPH.clientId,
      client_secret: GRAPH.secret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    })
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error('Could not get a Graph token: ' + (data.error_description || data.error || res.status).toString().split('\r')[0]);
  tokenCache = { value: data.access_token, expires: Date.now() + ((data.expires_in || 3600) - 300) * 1000 };
  return data.access_token;
}

async function graph(token, method, gpath, body, extraHeaders) {
  const res = await fetch('https://graph.microsoft.com/v1.0' + gpath, {
    method,
    headers: Object.assign({ Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, extraHeaders || {}),
    body: body ? JSON.stringify(body) : undefined
  });
  if (res.status === 204) return {};
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = new Error((data.error && data.error.message) || ('Graph error ' + res.status));
    err.status = res.status;
    throw err;
  }
  return data;
}

async function createEntraUser(name, email, department) {
  if (!GRAPH.tenant || !GRAPH.clientId || !GRAPH.secret || !GRAPH.domain) {
    return { skipped: true, reason: 'Graph settings are not all set' };
  }
  const token = await getGraphToken();
  const password = generateTempPassword();
  const baseNick = makeNickname(name, email);

  let nick = baseNick, created;
  for (let attempt = 0; attempt < 2 && !created; attempt++) {
    try {
      created = await graph(token, 'POST', '/users', {
        accountEnabled: true,
        displayName: name,
        mailNickname: nick,
        userPrincipalName: nick + '@' + GRAPH.domain,
        otherMails: [email],
        passwordProfile: { forceChangePasswordNextSignIn: true, password }
      });
    } catch (err) {
      if (attempt === 0 && /already exists|another object with the same value/i.test(err.message)) {
        nick = baseNick + crypto.randomInt(100, 1000);
      } else {
        throw err;
      }
    }
  }

  const result = { upn: created.userPrincipalName, password, groupAdded: false };
  const groupId = GROUP_MAP[department];
  if (!groupId) {
    result.groupWarning = 'No group is mapped to department "' + department + '"';
    return result;
  }
  try {
    await graph(token, 'POST', '/groups/' + groupId + '/members/$ref', {
      '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + created.id
    });
    result.groupAdded = true;
  } catch (err) {
    result.groupWarning = 'User created but could not be added to the group: ' + err.message;
  }
  return result;
}

async function recordRole(email, entra, department) {
  try {
    const pool = await getPool();
    await pool.request()
      .input('Email', sql.VarChar(100), email)
      .input('Upn', sql.VarChar(150), entra.upn || null)
      .input('Department', sql.VarChar(50), department || null)
      .input('GroupId', sql.VarChar(40), GROUP_MAP[department] || null)
      .input('GroupAdded', sql.Bit, entra.groupAdded ? 1 : 0)
      .query('INSERT INTO UserRoles (Email, UserPrincipalName, Department, EntraGroupId, GroupAdded) VALUES (@Email, @Upn, @Department, @GroupId, @GroupAdded)');
  } catch (e) {
    console.error('UserRoles write failed:', e.message);
  }
}

async function createUserFolder(email) {
  const svc = new BlobServiceClient(
    `https://${process.env.STORAGE_ACCOUNT_NAME}.blob.core.windows.net`,
    new DefaultAzureCredential()
  );
  const buf = await svc.getContainerClient('templates')
    .getBlobClient(TEMPLATE_NAME).downloadToBuffer();
  await svc.getContainerClient('orientation-files')
    .getBlockBlobClient(`${email.toLowerCase()}/${TEMPLATE_NAME}`)
    .uploadData(buf, { blobHTTPHeaders: { blobContentType: 'application/pdf' } });
}

// ---------- Offboarding ----------
async function findEntraUser(token, email) {
  // 1) the UPN we stored when the person was onboarded
  let upn = null;
  try {
    const pool = await getPool();
    const r = await pool.request().input('Email', sql.VarChar(100), email)
      .query('SELECT TOP 1 UserPrincipalName FROM UserRoles WHERE Email = @Email');
    if (r.recordset.length) upn = r.recordset[0].UserPrincipalName;
  } catch (e) { console.error('UserRoles lookup failed:', e.message); }
  if (upn) {
    try { return await graph(token, 'GET', '/users/' + encodeURIComponent(upn) + '?$select=id,userPrincipalName'); }
    catch (e) { if (e.status !== 404) throw e; }
  }
  // 2) fall back to the personal email stored on the Entra user
  const q = "/users?$filter=otherMails/any(c:c eq '" + email.replace(/'/g, "''") + "')&$select=id,userPrincipalName&$count=true";
  const found = await graph(token, 'GET', q, null, { ConsistencyLevel: 'eventual' });
  if (found.value && found.value.length > 1) throw new Error('More than one Entra user matches ' + email + '; remove them manually');
  return found.value && found.value[0] ? found.value[0] : null;
}

async function archiveUserFiles(email) {
  const svc = new BlobServiceClient(
    `https://${process.env.STORAGE_ACCOUNT_NAME}.blob.core.windows.net`,
    new DefaultAzureCredential()
  );
  const src = svc.getContainerClient('orientation-files');
  const dst = svc.getContainerClient('archived-files');
  await dst.createIfNotExists();
  const prefix = email.toLowerCase() + '/';
  const stamp = new Date().toISOString().slice(0, 10);
  let moved = 0;
  for await (const b of src.listBlobsFlat({ prefix })) {
    const srcBlob = src.getBlockBlobClient(b.name);
    const buf = await srcBlob.downloadToBuffer();
    await dst.getBlockBlobClient(prefix + 'offboarded-' + stamp + '/' + b.name.slice(prefix.length))
      .uploadData(buf, { blobHTTPHeaders: { blobContentType: (b.properties && b.properties.contentType) || 'application/octet-stream' } });
    await srcBlob.delete();
    moved++;
  }
  return moved;
}

// mode: 'disable' (reversible) or 'delete'. Returns { ok, httpStatus, error, name, steps[] }
async function offboardPerson(email, mode, actor) {
  const pool = await getPool();
  const rec = await pool.request().input('Email', sql.VarChar(100), email)
    .query('SELECT TOP 1 * FROM Classmates WHERE Email = @Email');
  if (!rec.recordset.length) return { ok: false, httpStatus: 404, error: 'No onboarded person has that email' };
  const person = rec.recordset[0];
  if (person.Status === 'Offboarded') return { ok: false, httpStatus: 409, error: person.Name + ' is already offboarded' };
  const name = person.Name, dept = person.Department;

  const steps = [];
  const add = async (step, status, detail) => {
    steps.push({ step, status, detail });
    await audit(name, email, dept, step, status, (detail || '') + (actor ? ' [by ' + actor + ']' : ''));
  };

  // 1+2. Entra: leave department groups, sign out, disable or delete
  let entraDone = false;
  if (!GRAPH.tenant || !GRAPH.clientId || !GRAPH.secret || !GRAPH.domain) {
    await add('Offboard-Entra', 'Skipped', 'Graph settings are not all set');
    entraDone = true;
  } else {
    try {
      const token = await getGraphToken();
      const user = await findEntraUser(token, email);
      if (!user) {
        await add('Offboard-Entra', 'Skipped', 'No Entra account found for this email');
        entraDone = true;
      } else {
        let removed = 0, groupErr = null;
        for (const gid of [...new Set(Object.values(GROUP_MAP))]) {
          try { await graph(token, 'DELETE', '/groups/' + gid + '/members/' + user.id + '/$ref'); removed++; }
          catch (e) { if (e.status !== 404) groupErr = e.message; } // 404 = was not a member
        }
        await add('Offboard-Groups', groupErr ? 'Failed' : 'Success', groupErr || ('Removed from ' + removed + ' group(s)'));
        try { await graph(token, 'POST', '/users/' + user.id + '/revokeSignInSessions'); } catch (e) { console.error('revoke failed:', e.message); }
        if (mode === 'delete') await graph(token, 'DELETE', '/users/' + user.id);
        else await graph(token, 'PATCH', '/users/' + user.id, { accountEnabled: false });
        await add('Offboard-Entra', 'Success', (mode === 'delete' ? 'Deleted ' : 'Disabled ') + user.userPrincipalName);
        entraDone = true;
      }
    } catch (e) {
      console.error('Offboard Entra failed:', e.message);
      await add('Offboard-Entra', 'Failed', e.message);
    }
  }

  // 3. Blob Storage: archive then remove the orientation files
  try {
    const n = await archiveUserFiles(email);
    await add('Offboard-Blob', 'Success', n + ' file(s) archived to archived-files');
  } catch (e) {
    console.error('Offboard blob failed:', e.message);
    await add('Offboard-Blob', 'Failed', e.message);
  }

  // 4. Only mark the person offboarded if their Entra access really is gone
  if (entraDone) {
    try {
      await pool.request().input('Email', sql.VarChar(100), email)
        .query("UPDATE Classmates SET Status = 'Offboarded', OffboardedAt = SYSUTCDATETIME() WHERE Email = @Email");
      await add('Offboard-DB', 'Success', 'Marked as offboarded');
    } catch (e) {
      await add('Offboard-DB', 'Failed', e.message + ' (did you run the ALTER TABLE script?)');
    }
  }

  return { ok: entraDone && steps.every(s => s.status !== 'Failed'), httpStatus: 200, name, mode, steps };
}

// ---------- One shared onboarding flow (single form AND bulk CSV use this) ----------
// Returns { ok, httpStatus, error, upn, problems[] }. ok = saved to the database.
async function onboardPerson(name, email, department) {
  const problems = [];

  // 1. Database
  try {
    const pool = await getPool();
    await pool.request()
      .input('Name', sql.VarChar(100), name)
      .input('Email', sql.VarChar(100), email)
      .input('Department', sql.VarChar(50), department)
      .query('INSERT INTO Classmates (Name, Email, Department) VALUES (@Name, @Email, @Department)');
  } catch (err) {
    console.error('INSERT failed:', err.message);
    if (err.number === 2627 || err.number === 2601) {
      await audit(name, email, department, 'Saved', 'Failed', 'Email already registered');
      return { ok: false, httpStatus: 409, error: 'That email is already registered' };
    }
    await audit(name, email, department, 'Saved', 'Failed', err.message);
    return { ok: false, httpStatus: 500, error: 'Database error: ' + err.message };
  }
  await audit(name, email, department, 'Saved', 'Success', null);

  // 2. Entra ID account + department group
  let entra = null;
  try {
    entra = await createEntraUser(name, email, department);
    if (entra.skipped) {
      problems.push('Entra ID step skipped: ' + entra.reason);
      await audit(name, email, department, 'EntraUser', 'Skipped', entra.reason);
    } else {
      await audit(name, email, department, 'EntraUser', 'Success', entra.upn);
      await recordRole(email, entra, department);
      await audit(name, email, department, 'GroupAdded', entra.groupAdded ? 'Success' : 'Failed',
        entra.groupAdded ? department : entra.groupWarning);
      if (!entra.groupAdded) problems.push(entra.groupWarning);
    }
  } catch (err) {
    console.error('Entra step failed:', err.message);
    entra = null;
    problems.push('Entra ID account could not be created: ' + err.message);
    await audit(name, email, department, 'EntraUser', 'Failed', err.message);
  }

  // 3. Welcome email (Logic App)
  if (LOGIC_APP_URL) {
    try {
      const mail = await fetch(LOGIC_APP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name, email, department,
          upn: entra && entra.upn ? entra.upn : null,
          password: entra && entra.password ? entra.password : null
        })
      });
      await audit(name, email, department, 'Email', mail.ok ? 'Success' : 'Failed', 'Logic App status ' + mail.status);
      if (!mail.ok) problems.push('Welcome email failed (Logic App status ' + mail.status + ')');
    } catch (emailErr) {
      console.error('Logic App trigger failed:', emailErr.message);
      await audit(name, email, department, 'Email', 'Failed', emailErr.message);
      problems.push('Welcome email failed: ' + emailErr.message);
    }
  }

  // 4. Orientation files (Blob Storage)
  try {
    await createUserFolder(email);
    await audit(name, email, department, 'Blob', 'Success', email.toLowerCase() + '/');
  } catch (e) {
    console.error('Blob step failed', e.message);
    await audit(name, email, department, 'Blob', 'Failed', e.message);
    problems.push('Orientation files failed: ' + e.message);
  }

  return { ok: true, upn: entra && entra.upn ? entra.upn : null, problems };
}

// ---------- Access control ----------
// Layer 1: App Service Authentication (Easy Auth) blocks anyone who is not signed in.
// Layer 2 (this code): only the people listed in the ADMIN_USERS app setting may use the API.
// Reason: new hires get accounts in the same tenant, so "signed in" alone does not mean "admin".
const ADMIN_USERS = (process.env.ADMIN_USERS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
const signedInUser = req => String(req.headers['x-ms-client-principal-name'] || '').trim();
const isAdmin = req => ADMIN_USERS.length === 0 || ADMIN_USERS.includes(signedInUser(req).toLowerCase());

app.get('/health', (req, res) => res.send('ok'));
app.get('/api/me', (req, res) => res.json({ user: signedInUser(req) || null, admin: isAdmin(req), enforced: ADMIN_USERS.length > 0 }));
app.use('/api', (req, res, next) => {
  if (isAdmin(req)) return next();
  console.warn('Blocked non-admin API call from', signedInUser(req) || '(no identity)');
  res.status(403).json({ error: 'Not authorised. Sign in with an administrator account.' });
});

// ---------- API Routes ----------

// One person at a time
app.post('/api/onboard', async (req, res) => {
  const body = req.body || {};
  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim();
  const department = normalizeDept(body.department);

  if (!name || !email) return res.status(400).json({ error: 'name and email are required' });
  if (!EMAIL_RE.test(email)) return res.status(400).json({ error: 'That email address does not look valid' });
  if (!department) return res.status(400).json({ error: 'Unknown department' });

  const r = await onboardPerson(name, email, department);
  if (!r.ok) return res.status(r.httpStatus).json({ error: r.error });

  let message = 'Saved! ' + name + ' was added to the database.';
  if (r.upn) message += '\n\nEntra ID account created.\nSign-in: ' + r.upn + '\n(Temporary password sent via email.)';
  if (r.problems.length) message += '\n\n' + r.problems.join('\n');
  res.json({ message, upn: r.upn, problems: r.problems });
});

// Offboard one person: { email, mode: 'disable' | 'delete' }
app.post('/api/offboard', async (req, res) => {
  const email = String((req.body || {}).email || '').trim();
  const mode = (req.body || {}).mode === 'delete' ? 'delete' : 'disable';
  if (!EMAIL_RE.test(email)) return res.status(400).json({ error: 'A valid email is required' });
  try {
    const r = await offboardPerson(email, mode, signedInUser(req) || null);
    res.status(r.httpStatus || 200).json(r);
  } catch (e) {
    console.error('Offboard error:', e.message);
    res.status(500).json({ error: 'Offboarding failed: ' + e.message });
  }
});

// Bulk CSV: Name, Email, Department
app.post('/api/onboard-bulk', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No CSV file uploaded' });

    let records;
    try {
      records = parse(req.file.buffer, {
        columns: header => header.map(h => String(h).replace(/^\uFEFF/, '').trim().toLowerCase()), // Excel adds a hidden BOM
        skip_empty_lines: true,
        trim: true,
        relax_column_count: true
      });
    } catch (e) {
      return res.status(400).json({ error: 'That file is not a valid CSV: ' + e.message });
    }

    if (!records.length) return res.status(400).json({ error: 'The CSV has no rows' });
    if (records.length > MAX_BULK_ROWS) {
      return res.status(400).json({ error: 'Too many rows (' + records.length + '). Upload at most ' + MAX_BULK_ROWS + ' people at a time.' });
    }

    const summary = { total: records.length, success: 0, partial: 0, failed: 0, details: [] };
    const seen = new Set();

    for (let i = 0; i < records.length; i++) {
      const row = records[i];
      const line = i + 2; // row 1 is the header
      const name = String(row.name || row.displayname || row['full name'] || '').trim();
      const email = String(row.email || row['personal email'] || '').trim();
      const department = normalizeDept(row.department);

      const fail = reason => {
        summary.failed++;
        summary.details.push({ line, name, email: email || '(blank)', status: 'Failed', reason });
      };

      if (!name || !email) { fail('Name and email are required'); continue; }
      if (!EMAIL_RE.test(email)) { fail('Invalid email address'); continue; }
      if (!department) { fail('Unknown or missing department "' + (row.department || '') + '"'); continue; }
      if (seen.has(email.toLowerCase())) { fail('Duplicate email inside this file'); continue; }
      seen.add(email.toLowerCase());

      try {
        const r = await onboardPerson(name, email, department);
        if (!r.ok) { fail(r.error); continue; }
        if (r.problems.length) {
          summary.partial++;
          summary.details.push({ line, name, email, department, status: 'Partial', upn: r.upn, reason: r.problems.join('; ') });
        } else {
          summary.success++;
          summary.details.push({ line, name, email, department, status: 'Success', upn: r.upn });
        }
      } catch (rowErr) {
        console.error('Bulk row error:', rowErr.message);
        fail(rowErr.message);
      }
    }

    res.json(summary);
  } catch (err) {
    console.error('Bulk onboard error:', err);
    res.status(500).json({ error: 'Failed to process CSV file: ' + err.message });
  }
});

app.get('/api/classmates', async (req, res) => {
  try {
    const pool = await getPool();
    const result = await pool.request().query('SELECT * FROM Classmates ORDER BY Id DESC');
    res.json(result.recordset);
  } catch (err) {
    console.error('SELECT failed:', err.message);
    res.status(500).json({ error: 'Database error: ' + err.message });
  }
});

// Catches upload errors (file too large etc.) and returns JSON instead of an HTML error page
app.use((err, req, res, next) => {
  if (err && err.code === 'LIMIT_FILE_SIZE') return res.status(400).json({ error: 'File is too large (max 1 MB)' });
  console.error('Unhandled error:', err && err.message);
  res.status(500).json({ error: 'Server error: ' + (err && err.message) });
});

const PORT = process.env.PORT || 8080;
if (require.main === module) {
  app.listen(PORT, () => console.log('Server running on port ' + PORT));
}
module.exports = { app, offboardPerson, makeNickname, generateTempPassword, createEntraUser, createUserFolder, normalizeDept };
