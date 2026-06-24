const http = require('http');
const fs = require('fs');
const path = require('path');

const boundary = 'FormBoundary' + Date.now();

function filePart(fieldName, filePath) {
  const fileName = path.basename(filePath);
  const fileData = fs.readFileSync(filePath);
  const header = Buffer.from(
    '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="' + fieldName + '"; filename="' + fileName + '"\r\n' +
    'Content-Type: application/octet-stream\r\n\r\n'
  );
  return Buffer.concat([header, fileData, Buffer.from('\r\n')]);
}

const parts = [
  filePart('rawFile',       'C:/Users/Vivob/Downloads/Raw File.xlsx'),
  filePart('prevWorking',   'C:/Users/Vivob/Downloads/Working Copy - 052026.xlsx'),
  filePart('nalMasterlist', 'C:/Users/Vivob/Downloads/Masterlist NAL.xlsx'),
  Buffer.from('--' + boundary + '--\r\n')
];
const body = Buffer.concat(parts);

const opts = {
  hostname: 'localhost', port: 3000, path: '/api/summary', method: 'POST',
  headers: {
    'Content-Type': 'multipart/form-data; boundary=' + boundary,
    'Content-Length': body.length
  }
};

const req = http.request(opts, res => {
  let d = '';
  res.on('data', c => d += c);
  res.on('end', () => {
    try { console.log(JSON.stringify(JSON.parse(d), null, 2)); }
    catch(e) { console.log(d); }
  });
});
req.on('error', e => console.error('ERROR:', e.message));
req.write(body);
req.end();
