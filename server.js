const express = require('express');
const sql = require('mssql');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

const LOGIC_APP_URL = 'https://prod-87.eastus.logic.azure.com:443/workflows/98725403991e4980a919359584b71fba/triggers/When_an_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_an_HTTP_request_is_received%2Frun&sv=1.0&sig=-c3m9l_EzxuxS2vQJxWYi8Y0RDDIsZkL4nZJAfPW0E8';

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

// ---------- API Routes ----------
app.get('/health', (req, res) => res.send('ok'));

app.post('/api/onboard', async (req, res) => {
  const body = req.body || {};
  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim();
  const department = String(body.department || '').trim();
  if (!name || !email) return res.status(400).json({ error: 'name and email are required' });

  // 1. Save entry to Azure SQL
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
      return res.status(409).json({ error: 'That email is already registered' });
    }
    return res.status(500).json({ error: 'Database error: ' + err.message });
  }

  // 2. Trigger Logic App Welcome Email
  if (LOGIC_APP_URL) {
    try {
      await fetch(LOGIC_APP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, email, department })
      });
    } catch (emailErr) {
      console.error('Logic App trigger failed:', emailErr.message);
    }
  }

  // 3. Create Entra ID account
  let message = 'Saved! ' + name + ' was added to the database.';
  let entra = null;
  try {
    entra = await createEntraUser(name, email, department);
    if (entra.skipped) {
      message += '\n\nEntra ID step skipped: ' + entra.reason + '.';
    } else {
      message += '\n\nEntra ID account created.\nSign-in: ' + entra.upn +
        '\nTemporary password: ' + entra.password;
    }
  } catch (err) {
    console.error('Entra step failed:', err.message);
    message += '\n\nEntra ID account could not be created: ' + err.message;
  }

  res.json({ message, upn: entra && entra.upn ? entra.upn : null });
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
module.exports = { app, makeNickname, generateTempPassword, createEntraUser };