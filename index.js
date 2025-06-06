const express = require('express');
const cors = require('cors');
const app = express();
const pool = require('./db');
const { google } = require('googleapis');
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
    appointment, requeste, report, type, textbox
  } = req.body;

  try {
    const query = `
      INSERT INTO sheet1(
        "Ticket ID", "User ID", "อีเมล", "ชื่อ", "เบอร์ติดต่อ", 
        "แผนก", "วันที่แจ้ง", "สถานะ", 
        "Appointment", "Requeste", "Report", "Type", "TEXTBOX"
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13)
    `;

    const values = [
      ticket_id, user_id, email, name, phone,
      department, created_at, status,
      appointment, requeste, report, type, textbox
    ];

    await pool.query(query, values);

    // เรียก sync กลับ Google Sheet อัตโนมัติ
    await updateGoogleSheetFromPostgres(ticket_id);

    console.log('✅ บันทึกลง PostgreSQL และ Google Sheet เรียบร้อย');
    res.status(200).send('✅ บันทึกสำเร็จ');
  } catch (err) {
    console.error('❌ บันทึกล้มเหลว:', err.message);
    res.status(500).send('❌ เกิดข้อผิดพลาดขณะบันทึกข้อมูล');
  }
});

// ฟังก์ชัน sync ข้อมูลจาก PostgreSQL กลับไป Google Sheet
async function updateGoogleSheetFromPostgres(ticketId) {
  try {
    const result = await pool.query(
      'SELECT * FROM sheet1 WHERE "Ticket ID" = $1',
      [ticketId]
    );
    if (result.rows.length === 0) {
      console.log('❌ ไม่พบข้อมูล Ticket ใน PostgreSQL');
      return;
    }
    const row = result.rows[0];
    const values = [
      row['Ticket ID'], row['User ID'], row['อีเมล'], row['ชื่อ'], row['เบอร์ติดต่อ'],
      row['แผนก'], row['วันที่แจ้ง'], row['สถานะ'],
      row['Appointment'], row['Requeste'], row['Report'], row['Type'], row['TEXTBOX']
    ];
    const sheets = google.sheets({ version: 'v4', auth });
    const findRowResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: 'sheet1!A:A',
    });
    const rows = findRowResponse.data.values || [];
    const rowIndex = rows.findIndex(r => r[0] === ticketId);
    if (rowIndex === -1) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'sheet1!A:M',
        valueInputOption: 'USER_ENTERED',
        resource: { values: [values] }
      });
      console.log('✅ เพิ่มข้อมูลใหม่ใน Google Sheet สำเร็จ');
    } else {
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: `sheet1!A${rowIndex + 1}:M${rowIndex + 1}`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [values] }
      });
      console.log('✅ อัปเดตข้อมูลใน Google Sheet สำเร็จ');
    }
  } catch (err) {
    console.error('❌ อัปเดต Google Sheet ล้มเหลว:', err.message);
  }
}

// Error handler (ควรไว้ล่างสุด)
app.use((err, req, res, next) => {
  console.error('❌ ข้อผิดพลาดเซิร์ฟเวอร์:', err);
  res.status(500).json({ error: 'เกิดข้อผิดพลาดในเซิร์ฟเวอร์' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 เซิร์ฟเวอร์ทำงานที่ http://localhost:${PORT}`);
});
