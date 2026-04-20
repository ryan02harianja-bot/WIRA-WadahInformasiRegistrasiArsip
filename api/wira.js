// ============================================================
// WIRA - Wadah Informasi Registrasi Arsip
// api/wira.js - Vercel Serverless Function
// Pengganti Code.gs (Google Apps Script)
// ============================================================

const { google } = require('googleapis');

// ============================================================
// KONFIGURASI — isi di Vercel Environment Variables
// SPREADSHEET_ID  : ID Google Spreadsheet kamu
// FOLDER_ID       : ID Google Drive Folder untuk PDF
// GOOGLE_SA_EMAIL : email service account (xxx@xxx.iam.gserviceaccount.com)
// GOOGLE_SA_KEY   : private key service account (-----BEGIN PRIVATE KEY-----....)
// ============================================================

function getAuth() {
  const auth = new google.auth.JWT(
    process.env.GOOGLE_SA_EMAIL,
    null,
    process.env.GOOGLE_SA_KEY.split('\\n').join('\n'),
    [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ]
  );
  return auth;
}

function getSheetsClient(auth) {
  return google.sheets({ version: 'v4', auth });
}

function getDriveClient(auth) {
  return google.drive({ version: 'v3', auth });
}

// ============================================================
// FORMAT TANGGAL  dd/mm/yyyy
// ============================================================
function formatTanggal(val) {
  if (!val) return '';
  const d = (val instanceof Date) ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val);
  const dd   = String(d.getDate()).padStart(2, '0');
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function extractFileId(url) {
  const match = String(url || '').match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ============================================================
// AMBIL DATA DASHBOARD
// ============================================================
async function ambilDataDashboard(auth) {
  const sheets = getSheetsClient(auth);
  const resp   = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: 'Sheet1'
  });

  const rows = resp.data.values || [];
  const hariIni = new Date();
  const blnIni  = hariIni.getMonth();
  const thnIni  = hariIni.getFullYear();
  const tglLalu = new Date();
  tglLalu.setMonth(hariIni.getMonth() - 1);
  const blnLalu = tglLalu.getMonth();
  const thnLalu = tglLalu.getFullYear();

  let countBulanIni  = 0;
  let countBulanLalu = 0;
  const semua = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;

    const tglSuratRaw = row[2] || '';
    let d = tglSuratRaw ? new Date(tglSuratRaw) : null;

    if (d && !isNaN(d.getTime())) {
      if (d.getMonth() === blnIni  && d.getFullYear() === thnIni)  countBulanIni++;
      if (d.getMonth() === blnLalu && d.getFullYear() === thnLalu) countBulanLalu++;
    }

    semua.push({
      rowIndex:  i + 1,
      no:        row[1] || '',
      tglSurat:  row[2] ? formatTanggal(row[2]) : '',
      tglTerima: row[3] ? formatTanggal(row[3]) : '',
      asal:      row[4] || '',
      perihal:   row[5] || '',
      link:      row[6] || '',
      status:    row[7] || 'Belum Diproses'
    });
  }

  const hasil       = [...semua].reverse();
  const instansiList = [...new Set(semua.map(d => d.asal).filter(Boolean))].sort();

  return {
    statistik:    { ini: countBulanIni, lalu: countBulanLalu, total: semua.length },
    terbaru:      hasil.slice(0, 20),
    semua:        hasil,
    instansiList
  };
}

// ============================================================
// UPLOAD SURAT (base64 PDF → Google Drive)
// ============================================================
async function uploadSurat(auth, obj) {
  if (!obj.noSurat || !obj.asalSurat || !obj.perihal) {
    return { ok: false, msg: 'Data tidak lengkap (nomor, asal, perihal wajib diisi).' };
  }
  if (!obj.fileSurat) {
    return { ok: false, msg: 'File PDF tidak ditemukan.' };
  }

  const drive  = getDriveClient(auth);
  const sheets = getSheetsClient(auth);

  // Upload ke Drive
  const buffer   = Buffer.from(obj.fileSurat, 'base64');
  const { Readable } = require('stream');
  const stream   = Readable.from(buffer);

  const driveResp = await drive.files.create({
    requestBody: {
      name:    obj.fileName,
      parents: [process.env.FOLDER_ID]
    },
    media: {
      mimeType: obj.mimeType,
      body:     stream
    },
    fields: 'id, webViewLink'
  });

  const fileId  = driveResp.data.id;
  const fileUrl = driveResp.data.webViewLink;

  // Set permission publik
  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' }
  });

  const tglSuratVal  = obj.tglSurat  ? new Date(obj.tglSurat).toISOString()  : '';
  const tglTerimaVal = obj.tglTerima ? new Date(obj.tglTerima).toISOString() : '';

  // Simpan ke Spreadsheet
  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range:         'Sheet1',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        new Date().toISOString(),
        obj.noSurat,
        tglSuratVal,
        tglTerimaVal,
        obj.asalSurat,
        obj.perihal,
        fileUrl,
        'Belum Diproses'
      ]]
    }
  });

  return { ok: true, msg: 'Berhasil Mengarsipkan!' };
}

// ============================================================
// EDIT SURAT
// ============================================================
async function editSurat(auth, obj) {
  if (!obj.rowIndex || !obj.noSurat || !obj.asalSurat || !obj.perihal) {
    return { ok: false, msg: 'Data tidak lengkap.' };
  }

  const sheets = getSheetsClient(auth);
  const row    = parseInt(obj.rowIndex);

  const tglSuratVal  = obj.tglSurat  ? new Date(obj.tglSurat).toISOString()  : '';
  const tglTerimaVal = obj.tglTerima ? new Date(obj.tglTerima).toISOString() : '';

  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range:         `Sheet1!B${row}:H${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        obj.noSurat,
        tglSuratVal,
        tglTerimaVal,
        obj.asalSurat,
        obj.perihal,
        '', // link tidak berubah saat edit — kolom G dilewati dengan cara ambil dulu
        obj.status || 'Belum Diproses'
      ]]
    }
  });

  // Ambil link existing supaya tidak tertimpa
  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `Sheet1!G${row}`
  });
  const existingLink = (existing.data.values && existing.data.values[0] && existing.data.values[0][0]) || '';

  // Update ulang kolom B-H dengan link yang benar
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range:         `Sheet1!B${row}:H${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        obj.noSurat,
        tglSuratVal,
        tglTerimaVal,
        obj.asalSurat,
        obj.perihal,
        existingLink,
        obj.status || 'Belum Diproses'
      ]]
    }
  });

  return { ok: true, msg: 'Data berhasil diperbarui!' };
}

// ============================================================
// HAPUS SURAT
// ============================================================
async function hapusSurat(auth, rowIndex) {
  const sheets = getSheetsClient(auth);
  const drive  = getDriveClient(auth);
  const row    = parseInt(rowIndex);

  // Ambil link PDF dulu
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.SPREADSHEET_ID,
      range: `Sheet1!G${row}`
    });
    const link = resp.data.values && resp.data.values[0] && resp.data.values[0][0];
    if (link) {
      const fileId = extractFileId(link);
      if (fileId) {
        await drive.files.delete({ fileId }).catch(() => {});
      }
    }
  } catch (e) { /* abaikan jika gagal ambil link */ }

  // Dapatkan sheetId (tab pertama)
  const meta = await sheets.spreadsheets.get({ spreadsheetId: process.env.SPREADSHEET_ID });
  const sheetId = meta.data.sheets[0].properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: process.env.SPREADSHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId,
            dimension: 'ROWS',
            startIndex: row - 1,
            endIndex:   row
          }
        }
      }]
    }
  });

  return { ok: true, msg: 'Data berhasil dihapus!' };
}

// ============================================================
// UPDATE STATUS
// ============================================================
async function updateStatus(auth, rowIndex, status) {
  const sheets = getSheetsClient(auth);
  const row    = parseInt(rowIndex);

  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range:         `Sheet1!H${row}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[status]] }
  });

  return { ok: true, msg: 'Status diperbarui!' };
}

// ============================================================
// EKSPOR CSV
// ============================================================
async function eksporData(auth) {
  const sheets = getSheetsClient(auth);
  const resp   = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: 'Sheet1'
  });

  const rows   = resp.data.values || [];
  const header = ['Timestamp','No Surat','Tgl Surat','Tgl Terima','Asal Instansi','Perihal','Link PDF','Status'];
  let csv      = header.join(';') + '\n';

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;
    const mapped = row.map((cell, idx) => {
      let val = (idx === 2 || idx === 3) ? formatTanggal(cell) : String(cell || '');
      return '"' + val.replace(/"/g, '""') + '"';
    });
    // Pastikan 8 kolom
    while (mapped.length < 8) mapped.push('""');
    csv += mapped.slice(0, 8).join(';') + '\n';
  }

  return { ok: true, csv };
}

// ============================================================
// MAIN HANDLER
// ============================================================
module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const auth = getAuth();

    // ── GET ──
    if (req.method === 'GET') {
      const action = req.query.action;
      if (action === 'muatDashboard') {
        const data = await ambilDataDashboard(auth);
        return res.status(200).json(data);
      }
      return res.status(400).json({ error: 'Action tidak dikenal.' });
    }

    // ── POST ──
    if (req.method === 'POST') {
      const body   = req.body;
      const action = body.action;

      if (action === 'uploadSurat')  return res.status(200).json(await uploadSurat(auth, body));
      if (action === 'editSurat')    return res.status(200).json(await editSurat(auth, body));
      if (action === 'hapusSurat')   return res.status(200).json(await hapusSurat(auth, body.rowIndex));
      if (action === 'updateStatus') return res.status(200).json(await updateStatus(auth, body.rowIndex, body.status));
      if (action === 'eksporData')   return res.status(200).json(await eksporData(auth));

      return res.status(400).json({ error: 'Action tidak dikenal.' });
    }

    return res.status(405).json({ error: 'Method not allowed.' });

  } catch (e) {
    console.error('WIRA API Error:', e);
    return res.status(500).json({ ok: false, msg: 'Server error: ' + e.message });
  }
};
