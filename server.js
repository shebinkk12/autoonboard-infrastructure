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
const upload = multer({ storage: multer.memoryStorage() });

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

const LOGIC_APP_URL = 'https://prod-87.eastus.logic.azure.com:443/workflows/98725403991e4980a919359584b71fba/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=-c3m9l_EzxuxS2vQJxWYi8Y0RDDIsZkL4nZJAfPW0E8';
const TEMPLATE_NAME = 'AutoOnboard Enterprise Policies and Governance Framework.pdf';

// ---------- Database Configuration ----------
const config = {
  user: process.env.DB_USER || 'autodbuser2026',
  password: process.env.DB_PASS || 'AutoOnboard@2026#Passs',
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

// ---------- Audit helpers ----------
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

async function getGraphToken() {
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
      if (attempt === 0 && /already exists/i.test(err.message)) {
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

// ---------- API Routes ----------
app.get('/health', (req, res) => res.send('ok'));

app.post('/api/onboard', async (req, res) => {
  const body = req.body || {};
  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim();
  const department = String(body.department || '').trim();
  if (!name || !email) return res.status(400).json({ error: 'name and email are required' });

  try {
    const pool = await getPool();
    await pool.request()
      .input('Name', sql.VarChar(100), name)
      .input('Email', sql.VarChar(100), email)
      .input('Department', sql.VarChar(50), department || null)
      .query('INSERT INTO Classmates (Name, Email, Department) VALUES (@Name, @Email, @Department)');
  } catch (err) {
    console.error('INSERT failed:', err.message);
    if (err.number === 2627 || err.number === 2601) {
      await audit(name, email, department, 'Saved', 'Failed', 'Email already registered');
      return res.status(409).json({ error: 'That email is already registered' });
    }
    await audit(name, email, department, 'Saved', 'Failed', err.message);
    return res.status(500).json({ error: 'Database error: ' + err.message });
  }
  await audit(name, email, department, 'Saved', 'Success', null);

  let message = 'Saved! ' + name + ' was added to the database.';
  let entra = null;
  try {
    entra = await createEntraUser(name, email, department);
    if (entra.skipped) {
      message += '\n\nEntra ID step skipped: ' + entra.reason + '.';
      await audit(name, email, department, 'EntraUser', 'Skipped', entra.reason);
    } else {
      message += '\n\nEntra ID account created.\nSign-in: ' + entra.upn + '\n(Temporary password sent via email.)';
      await audit(name, email, department, 'EntraUser', 'Success', entra.upn);
      await recordRole(email, entra, department);
      await audit(name, email, department, 'GroupAdded', entra.groupAdded ? 'Success' : 'Failed',
        entra.groupAdded ? department : entra.groupWarning);
    }
  } catch (err) {
    console.error('Entra step failed:', err.message);
    message += '\n\nEntra ID account could not be created: ' + err.message;
    await audit(name, email, department, 'EntraUser', 'Failed', err.message);
  }

  if (LOGIC_APP_URL) {
    try {
      const mail = await fetch(LOGIC_APP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name,
          email,
          department,
          upn: entra && entra.upn ? entra.upn : null,
          password: entra && entra.password ? entra.password : null
        })
      });
      await audit(name, email, department, 'Email', mail.ok ? 'Success' : 'Failed', 'Logic App status ' + mail.status);
    } catch (emailErr) {
      console.error('Logic App trigger failed:', emailErr.message);
      await audit(name, email, department, 'Email', 'Failed', emailErr.message);
    }
  }

  try {
    await createUserFolder(email);
    await audit(name, email, department, 'Blob', 'Success', email.toLowerCase() + '/');
  } catch (e) {
    console.error('Blob step failed', e.message);
    await audit(name, email, department, 'Blob', 'Failed', e.message);
  }

  res.json({ message, upn: entra && entra.upn ? entra.upn : null });
});

// ---------- Bulk CSV Onboarding Route ----------
app.post('/api/onboard-bulk', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No CSV file uploaded' });
    }

    const records = parse(req.file.buffer, {
      columns: true,
      skip_empty_lines: true,
      trim: true
    });

    const summary = { total: records.length, success: 0, failed: 0, details: [] };

    for (const row of records) {
      const name = String(row.Name || row.DisplayName || '').trim();
      const email = String(row.Email || row.email || '').trim();
      const department = String(row.Department || row.department || 'DevOps & Cloud Architecture').trim();

      if (!name || !email) {
        summary.failed++;
        summary.details.push({ email: email || 'unknown', status: 'Failed', reason: 'Name and email are required' });
        continue;
      }

      try {
        const pool = await getPool();
        try {
          await pool.request()
            .input('Name', sql.VarChar(100), name)
            .input('Email', sql.VarChar(100), email)
            .input('Department', sql.VarChar(50), department || null)
            .query('INSERT INTO Classmates (Name, Email, Department) VALUES (@Name, @Email, @Department)');
        } catch (dbErr) {
          if (dbErr.number === 2627 || dbErr.number === 2601) {
            await audit(name, email, department, 'Saved', 'Failed', 'Email already registered');
            summary.failed++;
            summary.details.push({ email, status: 'Failed', reason: 'Email already registered' });
            continue;
          }
          throw dbErr;
        }
        await audit(name, email, department, 'Saved', 'Success', null);

        let entra = null;
        try {
          entra = await createEntraUser(name, email, department);
          if (entra.skipped) {
            await audit(name, email, department, 'EntraUser', 'Skipped', entra.reason);
          } else {
            await audit(name, email, department, 'EntraUser', 'Success', entra.upn);
            await recordRole(email, entra, department);
            await audit(name, email, department, 'GroupAdded', entra.groupAdded ? 'Success' : 'Failed',
              entra.groupAdded ? department : entra.groupWarning);
          }
        } catch (entraErr) {
          await audit(name, email, department, 'EntraUser', 'Failed', entraErr.message);
        }

        if (LOGIC_APP_URL) {
          try {
            await fetch(LOGIC_APP_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                name,
                email,
                department,
                upn: entra && entra.upn ? entra.upn : null,
                password: entra && entra.password ? entra.password : null
              })
            });
            await audit(name, email, department, 'Email', 'Success', 'Logic App triggered');
          } catch (emailErr) {
            await audit(name, email, department, 'Email', 'Failed', emailErr.message);
          }
        }

        try {
          await createUserFolder(email);
          await audit(name, email, department, 'Blob', 'Success', email.toLowerCase() + '/');
        } catch (blobErr) {
          await audit(name, email, department, 'Blob', 'Failed', blobErr.message);
        }

        summary.success++;
        summary.details.push({ email, status: 'Success', upn: entra ? entra.upn : null });
      } catch (rowErr) {
        summary.failed++;
        summary.details.push({ email, status: 'Failed', reason: rowErr.message });
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

const PORT = process.env.PORT || 8080;
if (require.main === module) {
  app.listen(PORT, () => console.log('Server running on port ' + PORT));
}
module.exports = { app, makeNickname, generateTempPassword, createEntraUser, createUserFolder };