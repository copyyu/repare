const pool = require('./db') // ต้องเชื่อมกับ db.js ที่คุณสร้างไว้
const { google } = require('googleapis')
const auth = require('./auth')
require('dotenv').config()

async function syncSheetToPostgres() {
  try {
    const sheets = google.sheets({ version: 'v4', auth })
    
    // ดึงข้อมูลทั้งหมดจาก Google Sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: 'sheet1!A:M', // ดึงข้อมูลทั้งหมดจากคอลัมน์ A ถึง L
    })

    const rows = response.data.values || []
    console.log(`ดึงข้อมูลจาก Google Sheet ได้ทั้งหมด ${rows.length} แถว`)

    if (rows.length) {
      // เริ่ม transaction
      const client = await pool.connect()
      try {
        await client.query('BEGIN')

        // อัพเดทข้อมูลใน PostgreSQL
        for (let i = 1; i < rows.length; i++) { // ข้าม header (i = 0)
          const row = rows[i]
          const [
            ticket_id, user_id, email, name, phone,
            department, created_at, status,
            appointment, requeste, report, type, textbox
          ] = row

          await client.query(
            `INSERT INTO sheet1 (
              "Ticket ID", "User ID", "อีเมล", "ชื่อ", "เบอร์ติดต่อ",
              "แผนก", "วันที่แจ้ง", "สถานะ",
              "Appointment", "Requeste", "Report", "Type", "TEXTBOX"
            ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13)
            ON CONFLICT ("Ticket ID") DO UPDATE SET
              "User ID" = EXCLUDED."User ID",
              "อีเมล" = EXCLUDED."อีเมล",
              "ชื่อ" = EXCLUDED."ชื่อ",
              "เบอร์ติดต่อ" = EXCLUDED."เบอร์ติดต่อ",
              "แผนก" = EXCLUDED."แผนก",
              "วันที่แจ้ง" = EXCLUDED."วันที่แจ้ง",
              "สถานะ" = EXCLUDED."สถานะ",
              "Appointment" = EXCLUDED."Appointment",
              "Requeste" = EXCLUDED."Requeste",
              "Report" = EXCLUDED."Report",
              "Type" = EXCLUDED."Type",
              "TEXTBOX" = EXCLUDED."TEXTBOX"`,
            [
              ticket_id, user_id, email, name, phone,
              department, created_at, status,
              appointment, requeste, report, type, textbox
            ]
          )
        }

        await client.query('COMMIT')
        console.log('✅ ซิงค์ข้อมูลจาก Google Sheet ไปยัง PostgreSQL เรียบร้อย')
      } catch (err) {
        await client.query('ROLLBACK')
        throw err
      } finally {
        client.release()
      }
    }
  } catch (err) {
    console.error('❌ เกิดข้อผิดพลาดในการซิงค์ข้อมูล:', err)
    throw err
  }
}

module.exports = syncSheetToPostgres
