const express = require('express');
const cors = require('cors');
const app = express();
const pool = require('./db');
const google = require('googleapis');
const auth = require('./auth');

// Middleware
app.use(cors());
app.use(express.json());

// Import routers
const ticketsRouter = require('./routes/tickets');
const excelSyncRouter = require('./routes/excelSync');

// ใช้ router
app.use('/api/tickets', ticketsRouter);      // สำหรับ tickets CRUD
app.use('/api', excelSyncRouter);            // สำหรับ /api/excel-sync

// Webhook สำหรับรับข้อมูล Ticket ใหม่
app.post('/webhook', async (req, res) => {
  const {
    ticket_id, user_id, email, name, phone,
    department, created_at, status,
    appointment, requeste, report, textbox
  } = req.body;

  try {
    const query = `
      INSERT INTO sheet1(
        "Ticket ID", "User ID", "อีเมล", "ชื่อ", "เบอร์ติดต่อ", 
        "แผนก", "วันที่แจ้ง", "สถานะ", 
        "Appointment", "Requeste", "Report", "TEXTBOX"
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
    `;

    const values = [
      ticket_id, user_id, email, name, phone,
      department, created_at, status,
      appointment, requeste, report, textbox
    ];

    await pool.query(query, values);

    // อัพเดท Google Sheet
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      
      if (rowIndex === -1) {
        // ถ้าไม่มีข้อมูลใน Sheet ให้เพิ่มแถวใหม่
        await sheets.spreadsheets.values.append({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
<<<<<<< HEAD
          range: `sheet1!H${rowIndex + 1}`,
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
    console.error('[PUT] update error:', err);
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
  console.log('[PUT] /:ticket_id/message', { ticket_id, message });
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
      console.error('[PUT] TEXTBOX column missing');
      throw new Error('Column TEXTBOX does not exist in the table');
    }

    query = `
      UPDATE sheet1 SET "TEXTBOX" = $1 
      WHERE "Ticket ID" = $2 RETURNING ${selectCols}
    `;
    console.log('[PUT] update message query:', query, [message, ticket_id]);
    const result = await client.query(query, [message, ticket_id]);
    console.log('[PUT] update message result:', result.rows);

    // update Google Sheet ถ้ากำหนดไว้
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      if (rowIndex !== -1) {
        // M = 13
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          range: `sheet1!M${rowIndex + 1}`,
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
    console.error('[PUT] update message error:', err);
    await client.query('ROLLBACK');
    res.status(500).json({ error: 'Failed to update message', details: err.message });
  } finally {
    if (client) client.release();
  }
});

router.delete('/:ticket_id', async (req, res) => {
  const { ticket_id } = req.params;
  let client;
  console.log('[DELETE] /:ticket_id', { ticket_id });
  try {
    client = await pool.connect();
    await client.query('BEGIN');
    const deleteResult = await client.query(
      'DELETE FROM sheet1 WHERE "Ticket ID" = $1',
      [ticket_id]
    );
    console.log('[DELETE] delete result:', deleteResult.rowCount);

    // ลบใน Google Sheet ถ้ามี
    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === ticket_id);
      if (rowIndex !== -1) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
=======
          range: 'sheet1!A:L',
          valueInputOption: 'USER_ENTERED',
>>>>>>> cd92c261aeff0209daeced5b42a035f167bd602d
          resource: {
            values: [[
              ticket_id, user_id, email, name, phone,
              department, created_at, status,
              appointment, requeste, report, textbox
            ]]
          }
        });
      } else {
        // ถ้ามีข้อมูลแล้ว ให้อัพเดทแถวที่มีอยู่
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          range: `sheet1!A${rowIndex + 1}:L${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          resource: {
            values: [[
              ticket_id, user_id, email, name, phone,
              department, created_at, status,
              appointment, requeste, report, textbox
            ]]
          }
        });
      }
    }

    console.log('✅ บันทึกลง PostgreSQL และ Google Sheet เรียบร้อย');
    res.status(200).send('✅ บันทึกสำเร็จ');
  } catch (err) {
    console.error('❌ บันทึกล้มเหลว:', err.message);
    res.status(500).send('❌ เกิดข้อผิดพลาดขณะบันทึกข้อมูล');
  }
});

// Error handler (ควรไว้ล่างสุด)
app.use((err, req, res, next) => {
  console.error('❌ ข้อผิดพลาดเซิร์ฟเวอร์:', err);
  res.status(500).json({ error: 'เกิดข้อผิดพลาดในเซิร์ฟเวอร์' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 เซิร์ฟเวอร์ทำงานที่ http://localhost:${PORT}`);
});
