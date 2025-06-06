const express = require('express');
const cors = require('cors');
const app = express();
const pool = require('../db');
const { google } = require('googleapis');
const auth = require('../auth');

// Middleware
app.use(cors());
app.use(express.json());

// Import routers
const excelSyncRouter = require('./excelSync');

// ใช้ router
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
          range: 'sheet1!A:L',
          valueInputOption: 'USER_ENTERED',
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

module.exports = app; // หรือ module.exports = router; ถ้าใช้ express.Router()
