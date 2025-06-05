const cron = require('node-cron');
const syncSheetToPostgres = require('./syncGoogleSheet');

cron.schedule('*/5 * * * *', () => {
  console.log('Syncing data from Google Sheet...');
  syncSheetToPostgres();
});
