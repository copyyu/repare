const express = require('express');
const router = express.Router();
const pool = require('../db');
const { google } = require('googleapis');
const auth = require('../auth');

// รับข้อมูลจาก Power Automate (Excel Online → PostgreSQL)
router.post('/excel-sync', async (req, res) => {
  const { TicketID, Name, Status } = req.body;

  try {
    await pool.query(
      `INSERT INTO tickets ("Ticket ID", "Name", "Status")
       VALUES ($1, $2, $3)
       ON CONFLICT ("Ticket ID") DO UPDATE SET "Name"=$2, "Status"=$3`,
      [TicketID, Name, Status]
    );

    if (process.env.GOOGLE_SHEET_ID) {
      const sheets = google.sheets({ version: 'v4', auth });
      // หา rowIndex ของ Ticket ID ใน Google Sheet
      const findRowResponse = await sheets.spreadsheets.values.get({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: 'Sheet1!A:A',
      });
      const rows = findRowResponse.data.values || [];
      const rowIndex = rows.findIndex(row => row[0] === TicketID);
      if (rowIndex !== -1) {
        // สมมติคอลัมน์ H คือ "สถานะ"
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.GOOGLE_SHEET_ID,
          range: `Sheet1!H${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: [[Status]] },
        });
      }
    }

    res.status(200).json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;