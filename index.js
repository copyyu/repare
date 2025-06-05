const express = require('express');
const cors = require('cors');
const app = express();
const pool = require('./db');

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

    console.log('✅ บันทึกลง PostgreSQL เรียบร้อย');
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