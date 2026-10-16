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
const { BlobServiceClient } = require('@azure/storage-blob');
const { DefaultAzureCredential } = require('@azure/identity');
const multer = require('multer');
const { parse } = require('csv-parse/sync');

const MAX_BULK_ROWS = 25; // keeps one upload under Azure's ~230s request timeout
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 1024 * 1024 } });

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

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
  if (process.env.GROUP_MAP) GROUP_MAP = JSON.parse(process.env.GROUP_MAP);
} catch (e) {
  console.error('GROUP_MAP setting is not valid JSON, using built-in map');
}

// Lets a CSV say "HR" or "devops" instead of the full department name.
const DEPT_ALIASES = {
  'devops': 'DevOps & Cloud Architecture', 'devops & cloud architecture': 'DevOps & Cloud Architecture',
  'devops and cloud architecture': 'DevOps & Cloud Architecture', 'cloud': 'DevOps & Cloud Architecture',
  'engineering': 'Engineering',
  'hr': 'Human Resources', 'human resources': 'Human Resources',
  'finance': 'Finance', 'marketing': 'Marketing', 'operations': 'Operations', 'ops': 'Operations'
};
function normalizeDept(value) {
  const v = String(value || '').trim().toLowerCase();
  if (!v) return null;
  const hit = DEPT_ALIASES[v];
  if (hit && GROUP_MAP[hit]) return hit;
  const exact = Object.keys(GROUP_MAP).find(k => k.toLowerCase() === v);
  return exact || null;
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

async function graph(token, method, path, body) {
  const res = await fetch('https://graph.microsoft.com/v1.0' + path, {
    method,
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
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

// ---------- API Routes ----------
app.get('/health', (req, res) => res.send('ok'));

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
module.exports = { app, makeNickname, generateTempPassword, createEntraUser, createUserFolder, normalizeDept };
