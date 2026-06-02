require('dotenv').config();
const { google } = require('googleapis');
const path = require('path');

async function checkDrive() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!raw) {
    console.error('No GOOGLE_SERVICE_ACCOUNT_KEY found');
    return;
  }

  try {
    const creds = raw.trim().startsWith('{') ? JSON.parse(raw) : require(path.resolve(raw));
    const auth = new google.auth.GoogleAuth({
      credentials: creds,
      scopes: ['https://www.googleapis.com/auth/drive'],
    });
    const drive = google.drive({ version: 'v3', auth });

    console.log('Listing files...');
    const res = await drive.files.list({
      pageSize: 10,
      fields: 'nextPageToken, files(id, name, size)',
    });
    const files = res.data.files;
    if (files.length) {
      console.log('Files found:');
      files.forEach((file) => {
        console.log(`${file.name} (${file.id}) - ${file.size || 'unknown'} bytes`);
      });
    } else {
      console.log('No files found.');
    }

    console.log('\nChecking about...');
    const about = await drive.about.get({
      fields: 'storageQuota',
    });
    console.log('Storage Quota:', JSON.stringify(about.data.storageQuota, null, 2));

  } catch (e) {
    console.error('Error:', e.message);
  }
}

checkDrive();
