const sql = require('mssql');

const sqlConfig = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
    server: process.env.DB_SERVER,
    pool: { max: 10, min: 0, idleTimeoutMillis: 30000 },
    options: { encrypt: true, trustServerCertificate: false }
};

module.exports = async function (context, req) {
    context.log('Processing Onboarding Request...');
    const { name, email, department } = req.body || {};

    if (!name || !email || !department) {
        context.res = {
            status: 400,
            headers: { 'Content-Type': 'application/json' },
            body: { message: 'Missing required fields: name, email, or department' }
        };
        return;
    }

    try {
        let pool = await sql.connect(sqlConfig);
        const result = await pool.request()
            .input('Name', sql.NVarChar, name)
            .input('Email', sql.NVarChar, email)
            .input('Department', sql.NVarChar, department)
            .query(`
                INSERT INTO OnboardingLogs (Name, Email, Department)
                OUTPUT INSERTED.Id, INSERTED.CreatedDate
                VALUES (@Name, @Email, @Department)
            `);

        const insertedRecord = result.recordset[0];
        context.res = {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: {
                message: 'User successfully onboarded and saved to Azure SQL!',
                data: {
                    id: insertedRecord.Id,
                    name,
                    email,
                    department,
                    createdDate: insertedRecord.CreatedDate
                }
            }
        };
    } catch (err) {
        context.log.error('Database Error:', err.message);
        context.res = {
            status: 500,
            headers: { 'Content-Type': 'application/json' },
            body: { message: 'Database insertion failed', error: err.message }
        };
    }
};
