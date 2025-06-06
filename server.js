const cron = require('node-cron');
const syncSheetToPostgres = require('./syncGoogleSheet');

// รันการซิงค์ข้อมูลทุก 5 นาที
cron.schedule('*/1 * * * *', async () => {
  console.log('🔄 เริ่มการซิงค์ข้อมูลจาก Google Sheet...');
  try {
    await syncSheetToPostgres();
  } catch (err) {
    console.error('❌ การซิงค์ข้อมูลล้มเหลว:', err);
  }
});

// รันการซิงค์ครั้งแรกเมื่อเริ่มเซิร์ฟเวอร์
console.log('🚀 เริ่มการซิงค์ข้อมูลครั้งแรก...');
syncSheetToPostgres().catch(console.error);
