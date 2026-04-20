// ============================================================
// WIRA - api/wira.js
// ============================================================
const { google } = require('googleapis');

function getAuth() {
  let key = (process.env.GOOGLE_SA_KEY || '').replace(/\\n/g, '\n');
  return new google.auth.JWT({
    email:  process.env.GOOGLE_SA_EMAIL,
    key,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ]
  });
}

const sheets = (auth) => google.sheets({ version: 'v4', auth });
const drive  = (auth) => google.drive({ version: 'v3', auth });

function fmt(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d)) return String(val);
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

function fid(url) {
  const m = String(url||'').match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

async function getRows(auth) {
  const r = await sheets(auth).spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID, range: 'Sheet1'
  });
  return r.data.values || [];
}

async function ambilDataDashboard(auth) {
  const rows = await getRows(auth);
  const now  = new Date();
  const bIni = now.getMonth(), yIni = now.getFullYear();
  const prev = new Date(now); prev.setMonth(bIni - 1);
  const bLalu = prev.getMonth(), yLalu = prev.getFullYear();
  let cIni = 0, cLalu = 0;
  const semua = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[1]) continue;
    const d = r[2] ? new Date(r[2]) : null;
    if (d && !isNaN(d)) {
      if (d.getMonth()===bIni  && d.getFullYear()===yIni)  cIni++;
      if (d.getMonth()===bLalu && d.getFullYear()===yLalu) cLalu++;
    }
    semua.push({ rowIndex:i+1, no:r[1]||'', tglSurat:fmt(r[2]), tglTerima:fmt(r[3]),
      asal:r[4]||'', perihal:r[5]||'', link:r[6]||'', status:r[7]||'Belum Diproses' });
  }
  const hasil = [...semua].reverse();
  const instansiList = [...new Set(semua.map(d=>d.asal).filter(Boolean))].sort();
  return { statistik:{ini:cIni,lalu:cLalu,total:semua.length}, terbaru:hasil.slice(0,20), semua:hasil, instansiList };
}

async function simpanSurat(auth, obj) {
  if (!obj.noSurat || !obj.asalSurat || !obj.perihal || !obj.fileUrl)
    return { ok:false, msg:'Data tidak lengkap.' };
  await sheets(auth).spreadsheets.values.append({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: 'Sheet1', valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[
      new Date().toISOString(), obj.noSurat,
      obj.tglSurat  ? new Date(obj.tglSurat).toISOString()  : '',
      obj.tglTerima ? new Date(obj.tglTerima).toISOString() : '',
      obj.asalSurat, obj.perihal, obj.fileUrl, 'Belum Diproses'
    ]]}
  });
  return { ok:true, msg:'Berhasil Mengarsipkan!' };
}

async function editSurat(auth, obj) {
  if (!obj.rowIndex || !obj.noSurat || !obj.asalSurat || !obj.perihal)
    return { ok:false, msg:'Data tidak lengkap.' };
  const row = parseInt(obj.rowIndex);
  const ex = await sheets(auth).spreadsheets.values.get({
    spreadsheetId: process.env.SPREADSHEET_ID, range: `Sheet1!G${row}`
  });
  const link = ex.data.values?.[0]?.[0] || '';
  await sheets(auth).spreadsheets.values.update({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `Sheet1!B${row}:H${row}`, valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[
      obj.noSurat,
      obj.tglSurat  ? new Date(obj.tglSurat).toISOString()  : '',
      obj.tglTerima ? new Date(obj.tglTerima).toISOString() : '',
      obj.asalSurat, obj.perihal, link,
      obj.status || 'Belum Diproses'
    ]]}
  });
  return { ok:true, msg:'Data berhasil diperbarui!' };
}

async function hapusSurat(auth, rowIndex) {
  const row = parseInt(rowIndex);
  const meta = await sheets(auth).spreadsheets.get({ 
    spreadsheetId: process.env.SPREADSHEET_ID 
  });
  const sheetId = meta.data.sheets[0].properties.sheetId;
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: process.env.SPREADSHEET_ID,
    requestBody: { requests: [{ deleteDimension: {
      range: { sheetId, dimension:'ROWS', startIndex:row-1, endIndex:row }
    }}]}
  });
  return { ok:true, msg:'Data berhasil dihapus!' };
}

async function updateStatus(auth, rowIndex, status) {
  await sheets(auth).spreadsheets.values.update({
    spreadsheetId: process.env.SPREADSHEET_ID,
    range: `Sheet1!H${parseInt(rowIndex)}`, valueInputOption: 'RAW',
    requestBody: { values: [[status]] }
  });
  return { ok:true, msg:'Status diperbarui!' };
}

async function eksporData(auth) {
  const rows = await getRows(auth);
  let csv = ['Timestamp','No Surat','Tgl Surat','Tgl Terima','Asal Instansi','Perihal','Link PDF','Status'].join(';') + '\n';
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; if (!r||!r[1]) continue;
    const mapped = r.map((c,idx) => '"' + ((idx===2||idx===3)?fmt(c):String(c||'')).replace(/"/g,'""') + '"');
    while (mapped.length < 8) mapped.push('""');
    csv += mapped.slice(0,8).join(';') + '\n';
  }
  return { ok:true, csv };
}

// ── PARSE BODY HELPER ──
async function parseBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => { data += chunk.toString(); });
    req.on('end', () => {
      try { resolve(data ? JSON.parse(data) : {}); }
      catch(e) { reject(new Error('Invalid JSON: ' + data.slice(0,100))); }
    });
    req.on('error', reject);
  });
}

// ── MAIN HANDLER ──
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const auth = getAuth();

    if (req.method === 'GET') {
      if (req.query.action === 'muatDashboard')
        return res.json(await ambilDataDashboard(auth));
      return res.status(400).json({ error: 'Action tidak dikenal.' });
    }

    if (req.method === 'POST') {
      const body = await parseBody(req);
      const { action } = body;
      if (action === 'simpanSurat' || action === 'uploadSurat')
        return res.json(await simpanSurat(auth, body));
      if (action === 'editSurat')
        return res.json(await editSurat(auth, body));
      if (action === 'hapusSurat')
        return res.json(await hapusSurat(auth, body.rowIndex));
      if (action === 'updateStatus')
        return res.json(await updateStatus(auth, body.rowIndex, body.status));
      if (action === 'eksporData')
        return res.json(await eksporData(auth));
      return res.status(400).json({ error: 'Action tidak dikenal: ' + action });
    }

    return res.status(405).json({ error: 'Method not allowed.' });

  } catch(e) {
    console.error('WIRA API Error:', e.message);
    return res.status(500).json({ ok:false, msg:'Server error: ' + e.message });
  }
};
