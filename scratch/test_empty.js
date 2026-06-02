require('dotenv').config();
const { google } = require('googleapis');
const path = require('path');

async function testCreateEmpty() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  const creds = raw.trim().startsWith('{') ? JSON.parse(raw) : require(path.resolve(raw));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });

  try {
    console.log('Attempting to create an empty Google Sheet...');
    const res = await drive.files.create({
      requestBody: {
        name: 'Empty Sheet',
        mimeType: 'application/vnd.google-apps.spreadsheet',
      },
    });
    console.log('Success! ID:', res.data.id);
  } catch (e) {
    console.error('Failed:', e.message);
  }
}

testCreateEmpty();
