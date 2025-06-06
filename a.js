// ฟังก์ชัน sync ข้อมูลจาก PostgreSQL กลับไป Google Sheet
const { google } = require('googleapis');
const pool = require('./db');
const auth = require('./auth');
require('dotenv').config();

/**
 * อัปเดตข้อมูลใน Google Sheet ตามข้อมูลใน PostgreSQL (sheet1)
 * @param {string} ticketId - Ticket ID ที่ต้องการอัปเดต
 */
async function updateGoogleSheetFromPostgres(ticketId) {
  try {
    // 1. ดึงข้อมูลจาก PostgreSQL
    const result = await pool.query(
      'SELECT * FROM sheet1 WHERE "Ticket ID" = $1',
      [ticketId]
    );
    if (result.rows.length === 0) {
      console.log('❌ ไม่พบข้อมูล Ticket ใน PostgreSQL');
      return;
    }
    const row = result.rows[0];
    // 2. เตรียมข้อมูลสำหรับ Google Sheet
    const values = [
      row['Ticket ID'], row['User ID'], row['อีเมล'], row['ชื่อ'], row['เบอร์ติดต่อ'],
      row['แผนก'], row['วันที่แจ้ง'], row['สถานะ'],
      row['Appointment'], row['Requeste'], row['Report'], row['Type'], row['TEXTBOX']
    ];
    // 3. ค้นหา rowIndex ใน Google Sheet
    const sheets = google.sheets({ version: 'v4', auth });
    const findRowResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: 'sheet1!A:A',
    });
    const rows = findRowResponse.data.values || [];
    const rowIndex = rows.findIndex(r => r[0] === ticketId);
    if (rowIndex === -1) {
      // ถ้าไม่มีข้อมูล ให้เพิ่มแถวใหม่
      await sheets.spreadsheets.values.append({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'sheet1!A:M',
        valueInputOption: 'USER_ENTERED',
        resource: { values: [values] }
      });
      console.log('✅ เพิ่มข้อมูลใหม่ใน Google Sheet สำเร็จ');
    } else {
      // ถ้ามีข้อมูลแล้ว ให้อัปเดตแถวเดิม
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

// ตัวอย่างการเรียกใช้ (เช่น หลัง update ใน PostgreSQL)
// updateGoogleSheetFromPostgres('TICKET_ID_123');

module.exports = { updateGoogleSheetFromPostgres };
