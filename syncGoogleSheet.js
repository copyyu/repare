const pool = require('./db') // ต้องเชื่อมกับ db.js ที่คุณสร้างไว้
const { google } = require('googleapis')
require('dotenv').config()

async function syncSheetToPostgres() {
  // 1. Auth กับ Google Sheet
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json', // ต้องมีไฟล์นี้จาก Google Cloud Console
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  })

  const client = await auth.getClient()
  const sheets = google.sheets({ version: 'v4', auth: client })

  // 2. ดึงข้อมูลจาก Google Sheet
  const spreadsheetId = '1Hd2SV8sVZoPjyn3MhjoIQKSi_WIohvZkU4-nD7dpBGk'
  const range = 'Sheet1!A2:H' // ข้าม header (A2 ถึง H)

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  })

  const rows = response.data.values
  console.log(`ดึงข้อมูลจาก Google Sheet ได้ทั้งหมด ${rows.length} แถว`)

  if (rows.length) {
    for (let row of rows) {
      const ticket_id = row[0] || ''
      const user_id = row[1] || ''
      const email = row[2] || ''
      const name = row[3] || ''
      const phone = row[4] || ''
      const department = row[5] || ''
      const created_at = row[6] || ''
      const status = row[7] || ''

      await pool.query(
        `INSERT INTO "sheet1" (
          "Ticket ID", "User ID", "อีเมล", "ชื่อ", "เบอร์ติดต่อ", "แผนก", "วันที่แจ้ง", "สถานะ"
        ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
        ON CONFLICT ("Ticket ID") DO NOTHING`,
        [ticket_id, user_id, email, name, phone, department, created_at, status]
      )
    }
  }
}

syncSheetToPostgres().catch(console.error)
