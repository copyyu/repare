const express = require('express');
const router = express.Router();
const pool = require('../db');
const { google } = require('googleapis');
const Joi = require('joi');

// Google Sheets API Setup
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_CREDENTIALS_PATH || './credentials.json',
  scopes: 'https://www.googleapis.com/auth/spreadsheets',
});

// Schemas
const ticketStatusSchema = Joi.object({
  status: Joi.string().valid(
    'Pending',
    'In Progress',
    'Completed',
    'Rejected',
    'Waiting',
    'Scheduled'
  ).required()
});
const ticketMessageSchema = Joi.object({
  message: Joi.string().required()
});

// Middleware
const validateTicketStatus = (req, res, next) => {
  const { error } = ticketStatusSchema.validate(req.body);
  if (error) return res.status(400).json({ error: 'Invalid data', details: error.details });
  next();
};
const validateTicketMessage = (req, res, next) => {
  const { error } = ticketMessageSchema.validate(req.body);
  if (error) return res.status(400).json({ error: 'Invalid message', details: error.details });
  next();
};

// ฟังก์ชันจัดรูปแบบวันที่
function formatDate(dateString) {
  if (!dateString) return '-';
  try {
    const date = new Date(dateString);
    return date.toISOString()
      .replace('T', ' ')
      .replace(/\.\d+Z$/, '');
  } catch (e) {
    console.error('Error formatting date:', e);
    return dateString;
  }
}

// Routes
// GET: ดึงทุกคอลัมน์แบบ dynamic
router.get('/', async (req, res) => {
  let client;
  let query = '';
  let params = [];
  try {
    client = await pool.connect();

    // ดึง column ทั้งหมดในตาราง sheet1
    const colResult = await client.query(`
      SELECT column_name 
      FROM information_schema.columns 
      WHERE table_name = 'sheet1'
      ORDER BY ordinal_position
    `);
    const columns = colResult.rows.map(row => row.column_name);

    // สร้าง select column string
    const selectCols = columns.map(col => `"${col}"`).join(', ');

    const { type, status, search } = req.query;
    query = `SELECT ${selectCols} FROM sheet1 WHERE 1=1`;

    if (type) {
      query += ` AND "Type" = $${params.length + 1}`;
      params.push(type);
    }
    if (status) {
      query += ` AND "สถานะ" = $${params.length + 1}`;
      params.push(status);
    }
    if (search) {
      query += ` AND (
        "Ticket ID"::TEXT LIKE $${params.length + 1} OR 
        "User ID" LIKE $${params.length + 1} OR
        "ชื่อ" LIKE $${params.length + 1} OR
        "อีเมล" LIKE $${params.length + 1} OR
        "เบอร์ติดต่อ" LIKE $${params.length + 1}
      )`;
      params.push(`%${search}%`);
    }
    query += ` ORDER BY "วันที่แจ้ง" DESC`;

    const result = await client.query(query, params);

    // แปลงวันที่ให้อ่านง่ายขึ้น (ถ้ามี key "วันที่แจ้ง")
    const formattedData = result.rows.map(row => {
      const obj = { ...row };
      if ('วันที่แจ้ง' in obj) {
        obj['วันที่แจ้ง'] = formatDate(obj['วันที่แจ้ง']);
      }
      return obj;
    });

    res.json(formattedData);
  } catch (err) {
    console.error('Database error:', {
      message: err.message,
      stack: err.stack,
      query: query,
      params: params
    });
    res.status(500).json({ 
      error: 'Failed to fetch tickets',
      details: {
        message: err.message,
        query: query,
        params: params
      }
    });
  } finally {
    if (client) client.release();
  }
});

// ส่วน PUT, DELETE, message อื่นๆ ใช้ logic เดิม (แต่ให้คืน field ทั้งหมดเหมือน GET)
router.put('/:ticket_id', validateTicketStatus, async (req, res) => {
  const { ticket_id } = req.params;
  const { status } = req.body;
  let client;
  let query = '';
  try {
    client = await pool.connect();
    await client.query('BEGIN');

    // ดึง columns ทั้งหมด
    const colResult = await client.query(`
      SELECT column_name 
      FROM information_schema.columns 
      WHERE table_name = 'sheet1'
      ORDER BY ordinal_position
    `);
    const columns = colResult.rows.map(row => row.column_name);
    const selectCols = columns.map(col => `"${col}"`).join(', ');

    query = `
      UPDATE sheet1 SET "สถานะ" = $1 
      WHERE "Ticket ID" = $2 RETURNING ${selectCols}
    `;
    const result = await client.query(query, [status, ticket_id]);

    // update Google Sheet ถ้ากำหนดไว้
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'Sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      if (rowIndex !== -1) {
        // H = 8
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          range: `Sheet1!H${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[status]] },
        });
      }
    }
    await client.query('COMMIT');
    const obj = { ...result.rows[0] };
    if ('วันที่แจ้ง' in obj) obj['วันที่แจ้ง'] = formatDate(obj['วันที่แจ้ง']);
    res.json(obj);
  } catch (err) {
    await client.query('ROLLBACK');
    res.status(500).json({ error: 'Failed to update ticket', details: err.message });
  } finally {
    if (client) client.release();
  }
});

router.put('/:ticket_id/message', validateTicketMessage, async (req, res) => {
  const { ticket_id } = req.params;
  const { message } = req.body;
  let client;
  let query = '';
  try {
    client = await pool.connect();
    await client.query('BEGIN');

    // ดึง columns ทั้งหมด
    const colResult = await client.query(`
      SELECT column_name 
      FROM information_schema.columns 
      WHERE table_name = 'sheet1'
      ORDER BY ordinal_position
    `);
    const columns = colResult.rows.map(row => row.column_name);
    const selectCols = columns.map(col => `"${col}"`).join(', ');

    // ตรวจสอบว่ามี TEXTBOX หรือไม่
    if (!columns.map(c => c.toUpperCase()).includes('TEXTBOX')) {
      throw new Error('Column TEXTBOX does not exist in the table');
    }

    query = `
      UPDATE sheet1 SET "TEXTBOX" = $1 
      WHERE "Ticket ID" = $2 RETURNING ${selectCols}
    `;
    const result = await client.query(query, [message, ticket_id]);

    // update Google Sheet ถ้ากำหนดไว้
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'Sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      if (rowIndex !== -1) {
        // M = 13
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          range: `Sheet1!M${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[message]] },
        });
      }
    }
    await client.query('COMMIT');
    const obj = { ...result.rows[0] };
    if ('วันที่แจ้ง' in obj) obj['วันที่แจ้ง'] = formatDate(obj['วันที่แจ้ง']);
    res.json(obj);
  } catch (err) {
    await client.query('ROLLBACK');
    res.status(500).json({ error: 'Failed to update message', details: err.message });
  } finally {
    if (client) client.release();
  }
});

router.delete('/:ticket_id', async (req, res) => {
  const { ticket_id } = req.params;
  let client;
  try {
    client = await pool.connect();
    await client.query('BEGIN');
    await client.query(
      'DELETE FROM sheet1 WHERE "Ticket ID" = $1',
      [ticket_id]
    );

    // ลบใน Google Sheet ถ้ามี
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'Sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      if (rowIndex !== -1) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          resource: {
            requests: [{
              deleteDimension: {
                range: {
                  sheetId: 0,
                  dimension: 'ROWS',
                  startIndex: rowIndex,
                  endIndex: rowIndex + 1
                }
              }
            }]
          }
        });
      }
    }

    await client.query('COMMIT');
    res.json({ success: true, message: 'Ticket deleted' });
  } catch (err) {
    await client.query('ROLLBACK');
    res.status(500).json({ error: 'Failed to delete ticket', details: err.message });
  } finally {
    if (client) client.release();
  }
});

module.exports = router;