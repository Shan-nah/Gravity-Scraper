require('dotenv').config();
const { google } = require('googleapis');
const path = require('path');
const { Readable } = require('stream');

async function testCreate() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  const creds = raw.trim().startsWith('{') ? JSON.parse(raw) : require(path.resolve(raw));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });

  try {
    console.log('Attempting to create a tiny file...');
    const res = await drive.files.create({
      requestBody: { name: 'test.txt' },
      media: {
        mimeType: 'text/plain',
        body: Readable.from(['hello world']),
      },
    });
    console.log('Success! ID:', res.data.id);
  } catch (e) {
    console.error('Failed:', e.message);
  }
}

testCreate();
