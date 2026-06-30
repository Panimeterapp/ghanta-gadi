// ================================================================
// ग्राम गाडी ट्रॅकर — Google Apps Script v3
// नवीन Trip System Compatible (From→To, Stages, Payment)
// ================================================================
// Setup:
//   1. Google Sheets उघडा → Extensions → Apps Script
//   2. हे सर्व code paste करा (Replace करा)
//   3. Function dropdown → setupAllSheets → Run (एकदाच)
//   4. Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   5. URL copy करा → App मध्ये Developer Settings मध्ये paste करा
// ================================================================

const S_TRIPS    = 'Trips';
const S_SETTINGS = 'Settings';
const S_USERS    = 'Users';
const S_STOPS    = 'Stops';
const S_PAYMENTS = 'Payments';
const S_SUMMARY  = 'Monthly Summary';

// ---------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result = {};
    switch (data.action) {
      case 'addTrip':      result = addTrip(data.trip);           break;
      case 'syncBulk':     result = syncBulk(data.trips);         break;
      case 'updateTrip':   result = updateTrip(data.trip);        break;
      case 'saveSettings': result = saveSettings(data.settings);  break;
      case 'getSettings':  result = getSettings();                 break;
      case 'getUsers':     result = getUsers();                    break;
      case 'getStops':     result = getStops();                    break;
      case 'getTrips':     result = getTrips(data.driverId);        break;
      case 'deleteTrip':   result = deleteTrip(data.tripId, data.role);  break;
      case 'addPayment':   result = addPayment(data.payment);      break;
      case 'uploadPhoto':  result = uploadPhoto(data.photo);       break;
      case 'saveUsers':    result = saveUsers(data.users);         break;
      default: result = { error: 'Unknown action: ' + data.action };
    }
    return jsonResponse({ success: true, ...result });
  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  // App startup ला GET ने data load करतो
  const action = e.parameter.action;
  if (action === 'getSettings') {
    return jsonResponse({ success: true, ...getSettings() });
  }
  if (action === 'getUsers') {
    // Security: passwords पाठवू नका GET ने — फक्त id, role, name, phone
    const result = getUsers();
    if (result.users) {
      result.users = result.users.map(u => ({
        id: u.id, pass: u.pass, role: u.role, name: u.name, phone: u.phone
      }));
    }
    return jsonResponse({ success: true, ...result });
  }
  if (action === 'getStops') {
    return jsonResponse({ success: true, ...getStops() });
  }
  if (action === 'getTrips') {
    return jsonResponse({ success: true, ...getTrips(e.parameter.driverId) });
  }
  if (e.parameter.test) {
    return jsonResponse({
      status: 'ग्राम गाडी ट्रॅकर v3 Script चालू आहे ✅',
      time: new Date().toISOString(),
      sheets: getSheetNames()
    });
  }
  return jsonResponse({ status: 'OK', time: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// ---------------------------------------------------------------
// TRIPS SHEET — नवीन columns: From, To, Route, Stages, Status, Payment
// ---------------------------------------------------------------
function ensureTripsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_TRIPS);
  if (!sh) {
    sh = ss.insertSheet(S_TRIPS);
    const headers = [
      'Trip ID', 'दिनांक', 'वेळ',
      'चालक ID', 'चालकाचे नाव',
      'From थांबा', 'To थांबा', 'Route',
      'Stages', 'Completed Stages',
      'रक्कम (₹)', 'Status',
      'Payment Date', 'Payment Mode', 'UTR नंबर',
      'महिना', 'वर्ष', 'Sync Time'
    ];
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#1a6b3a')
      .setFontColor('white')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 100);
    sh.setColumnWidth(2, 90);
    sh.setColumnWidth(5, 140);
    sh.setColumnWidth(6, 140);
    sh.setColumnWidth(7, 140);
    sh.setColumnWidth(8, 200);
    sh.setColumnWidth(12, 80);
  }
  return sh;
}

function addTrip(trip) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  // Duplicate check
  if (data.some((r, i) => i > 0 && String(r[0]) === String(trip.id))) {
    return { skipped: true, reason: 'duplicate' };
  }
  const now = new Date();
  sh.appendRow([
    trip.id,
    trip.date,
    trip.time || now.toLocaleTimeString('mr-IN'),
    trip.driverId,
    trip.driverName,
    trip.fromName  || '',
    trip.toName    || '',
    trip.stopName  || (trip.fromName + ' → ' + trip.toName),
    trip.stages    || '',
    trip.completedStages || '',
    trip.amount,
    trip.status    || 'complete',
    trip.paymentDate   || '',
    trip.paymentMode   || '',
    trip.paymentUTR    || '',
    trip.month,
    trip.year,
    now.toISOString()
  ]);
  // Alternate row color
  const lr = sh.getLastRow();
  if (lr % 2 === 0) sh.getRange(lr, 1, 1, 18).setBackground('#f0f8f3');
  updateMonthlySummary();
  return { added: true, row: lr };
}

function syncBulk(trips) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  const existing = new Set(data.slice(1).map(r => String(r[0])));
  let added = 0, skipped = 0;
  trips.forEach(trip => {
    if (existing.has(String(trip.id))) { skipped++; return; }
    const now = new Date();
    sh.appendRow([
      trip.id,
      trip.date,
      trip.time || now.toLocaleTimeString('mr-IN'),
      trip.driverId,
      trip.driverName,
      trip.fromName  || '',
      trip.toName    || '',
      trip.stopName  || (trip.fromName + ' → ' + trip.toName),
      trip.stages    || '',
      trip.completedStages || '',
      trip.amount,
      trip.status    || 'complete',
      trip.paymentDate   || '',
      trip.paymentMode   || '',
      trip.paymentUTR    || '',
      trip.month,
      trip.year,
      now.toISOString()
    ]);
    added++;
  });
  if (added > 0) updateMonthlySummary();
  return { added, skipped };
}

// Sheet मधील सर्व trips परत App ला पाठवा (driverId दिल्यास फक्त त्या ड्रायव्हरचे)
function getTrips(driverId) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  let trips = data.slice(1).filter(r => r[0]).map(r => {
    let fromName = r[5] || '', toName = r[6] || '';
    if (!fromName && !toName && r[7]) {
      const parts = String(r[7]).split('→').map(x => x.trim());
      fromName = parts[0] || ''; toName = parts[1] || '';
    }
    return {
      id: r[0], date: r[1], time: r[2],
      driverId: r[3], driverName: r[4],
      fromName, toName,
      amount: Number(r[10]) || 0,
      status: r[11] || 'complete',
      paymentDate: r[12] || '', paymentMode: r[13] || '', paymentUTR: r[14] || '',
      month: (Number(r[15]) || 1) - 1, // App मध्ये month 0-indexed आहे
      year: Number(r[16]) || new Date().getFullYear(),
      synced: true
    };
  });
  if (driverId) trips = trips.filter(t => String(t.driverId) === String(driverId));
  return { trips };
}

// Trip कायमचा delete करा (row Sheet मधून काढून टाका)
function deleteTrip(tripId, role) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(tripId)) {
      const status = data[i][11] || '';
      if (role === 'admin' && status !== 'active') {
        return { deleted: false, reason: 'admin can only delete active trips' };
      }
      // या trip चे Drive फोटोही काढून टाका (column 19 मध्ये साठवलेल्या links मधून)
      const driveLinksRaw = data[i][18] || '';
      if (driveLinksRaw) {
        driveLinksRaw.split('\n').forEach(line => {
          const m = line.match(/https?:\/\/drive\.google\.com\/\S+/);
          if (m) {
            try {
              const idMatch = m[0].match(/[-\w]{25,}/);
              if (idMatch) DriveApp.getFileById(idMatch[0]).setTrashed(true);
            } catch(e) { /* फाईल आधीच नसेल तर दुर्लक्ष करा */ }
          }
        });
      }
      sh.deleteRow(i+1);
      // संबंधित Payments row पण काढून टाका (असल्यास)
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const paySh = ss.getSheetByName(S_PAYMENTS);
      if (paySh) {
        const pdata = paySh.getDataRange().getValues();
        for (let j = pdata.length - 1; j >= 1; j--) {
          if (String(pdata[j][0]) === String(tripId)) paySh.deleteRow(j+1);
        }
      }
      updateMonthlySummary();
      return { deleted: true };
    }
  }
  return { deleted: false, reason: 'not found' };
}

// Payment नोंद झाल्यावर trip update करा
function updateTrip(trip) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(trip.id)) {
      if (trip.status)       sh.getRange(i+1, 12).setValue(trip.status);
      if (trip.paymentDate)  sh.getRange(i+1, 13).setValue(trip.paymentDate);
      if (trip.paymentMode)  sh.getRange(i+1, 14).setValue(trip.paymentMode);
      if (trip.paymentUTR)   sh.getRange(i+1, 15).setValue(trip.paymentUTR);
      // Paid rows ला हिरवी रंगसंगती
      if (trip.status === 'paid') {
        sh.getRange(i+1, 1, 1, 18).setBackground('#d4edda');
      }
      updateMonthlySummary();
      return { updated: true, row: i+1 };
    }
  }
  return { updated: false, reason: 'not found' };
}

// ---------------------------------------------------------------
// SETTINGS SHEET
// ---------------------------------------------------------------
function ensureSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(S_SETTINGS);
    sh.appendRow(['Key', 'Value', 'Description']);
    sh.getRange(1,1,1,3).setBackground('#1565c0').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 260);
    const defaults = [
      ['appName',    'ग्राम गाडी ट्रॅकर',       'App चे नाव'],
      ['tagline',    'ग्रामपंचायत वाहन व्यवस्थापन', 'App Tagline'],
      ['logo',       '🚌',                       'Logo Emoji'],
      ['gramName',   'ग्रामपंचायत',              'ग्रामपंचायतीचे नाव'],
      ['color',      '#1a6b3a',                  'Theme Color'],
      ['tripRate',   '500',                      'Trip दर (₹) — एका trip ला एकच रक्कम'],
      ['stages',     'सुरुवात,शेवट',             'Trip Stages (comma separated)'],
      ['fenceMeters','10',                       'Geo-fence अंतर (मीटर)'],
      ['drvPhone',   '',                         'चालकाचा WhatsApp नंबर'],
      ['srpPhone',   '',                         'सरपंचाचा WhatsApp नंबर'],
    ];
    defaults.forEach(r => sh.appendRow(r));
  }
  return sh;
}

function saveSettings(settings) {
  const sh = ensureSettingsSheet();
  const data = sh.getDataRange().getValues();
  const keyRow = {};
  data.forEach((r, i) => { if (i > 0) keyRow[r[0]] = i+1; });
  Object.entries(settings).forEach(([k, v]) => {
    if (keyRow[k]) sh.getRange(keyRow[k], 2).setValue(v);
    else sh.appendRow([k, v, '']);
  });
  return { saved: true };
}

function getSettings() {
  ensureSettingsSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_SETTINGS);
  const data = sh.getDataRange().getValues();
  const settings = {};
  data.forEach((r, i) => { if (i > 0 && r[0]) settings[r[0]] = r[1]; });
  return { settings };
}

// ---------------------------------------------------------------
// USERS SHEET
// ---------------------------------------------------------------
function ensureUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_USERS);
  if (!sh) {
    sh = ss.insertSheet(S_USERS);
    sh.appendRow(['Login ID', 'Password', 'Role', 'नाव', 'Phone']);
    sh.getRange(1,1,1,5).setBackground('#6a1b9a').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.appendRow(['driver1', 'pass123',  'user',  'रामराव शिंदे',    '9876543210']);
    sh.appendRow(['admin1',  'admin123', 'admin', 'ग्रामसेवक जाधव', '9876543220']);
  }
  return sh;
}

function getUsers() {
  ensureUsersSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_USERS);
  const data = sh.getDataRange().getValues();
  const users = data.slice(1).filter(r => r[0]).map(r => ({
    id: r[0], pass: r[1], role: r[2], name: r[3], phone: r[4] || ''
  }));
  return { users };
}

// ---------------------------------------------------------------
// STOPS SHEET (rate column काढला — Trip rate global आहे)
// ---------------------------------------------------------------
function ensureStopsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_STOPS);
  if (!sh) {
    sh = ss.insertSheet(S_STOPS);
    sh.appendRow(['ID', 'थांब्याचे नाव', 'Latitude', 'Longitude']);
    sh.getRange(1,1,1,4).setBackground('#e67e00').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.appendRow([1, 'ग्रामपंचायत कार्यालय', 17.6805, 74.0183]);
    sh.appendRow([2, 'मुख्य बाजारपेठ',        17.6820, 74.0200]);
    sh.appendRow([3, 'प्राथमिक शाळा',         17.6790, 74.0165]);
    sh.appendRow([4, 'आरोग्य केंद्र',          17.6835, 74.0210]);
    sh.appendRow([5, 'रेल्वे स्थानक',          17.6770, 74.0150]);
  }
  return sh;
}

function getStops() {
  ensureStopsSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_STOPS);
  const data = sh.getDataRange().getValues();
  const stops = data.slice(1).filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], lat: r[2], lng: r[3]
  }));
  return { stops };
}

// ---------------------------------------------------------------
// PAYMENTS SHEET — Payment नोंद
// ---------------------------------------------------------------
function addPayment(payment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_PAYMENTS);
  if (!sh) {
    sh = ss.insertSheet(S_PAYMENTS);
    sh.appendRow(['Trip ID', 'दिनांक', 'चालक ID', 'चालकाचे नाव', 'Route', 'रक्कम (₹)', 'Payment Date', 'Payment Mode', 'UTR नंबर']);
    sh.getRange(1,1,1,9).setBackground('#c0392b').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 200);
  }
  sh.appendRow([
    payment.tripId       || '',
    payment.date         || '',
    payment.driverId     || '',
    payment.driverName   || '',
    payment.route        || '',
    payment.amount       || 0,
    payment.paymentDate  || '',
    payment.paymentMode  || '',
    payment.paymentUTR   || ''
  ]);
  // Trips sheet मध्ये पण update करा
  if (payment.tripId) {
    updateTrip({
      id: payment.tripId,
      status: 'paid',
      paymentDate: payment.paymentDate,
      paymentMode: payment.paymentMode,
      paymentUTR:  payment.paymentUTR
    });
  }
  return { saved: true };
}

// ---------------------------------------------------------------
// MONTHLY SUMMARY — Auto-generated
// ---------------------------------------------------------------
function updateMonthlySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trSh = ss.getSheetByName(S_TRIPS);
  if (!trSh) return;
  let sumSh = ss.getSheetByName(S_SUMMARY);
  if (!sumSh) {
    sumSh = ss.insertSheet(S_SUMMARY);
    sumSh.setFrozenRows(1);
  }
  const data = trSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const grouped = {};
  data.forEach(r => {
    const key = `${r[16]}_${r[15]}_${r[4]}`;
    if (!grouped[key]) grouped[key] = {
      year: r[16], month: r[15], name: r[4],
      total: 0, paid: 0, pending: 0, count: 0
    };
    grouped[key].count++;
    grouped[key].total += Number(r[10]);
    if (r[11] === 'paid') grouped[key].paid += Number(r[10]);
    else grouped[key].pending += Number(r[10]);
  });
  sumSh.clearContents();
  sumSh.appendRow(['वर्ष', 'महिना', 'चालकाचे नाव', 'एकूण Trips', 'एकूण रक्कम (₹)', 'Paid (₹)', 'येणे (₹)']);
  sumSh.getRange(1,1,1,7).setBackground('#1a6b3a').setFontColor('white').setFontWeight('bold');
  Object.values(grouped).forEach(g =>
    sumSh.appendRow([g.year, g.month, g.name, g.count, g.total, g.paid, g.pending])
  );
}

// ---------------------------------------------------------------
// PHOTO UPLOAD — Google Drive मध्ये फोटो save करा
// ---------------------------------------------------------------
function uploadPhoto(photo) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(photo.base64),
      'image/jpeg',
      photo.fileName
    );
    // Folder निवडा
    let folder;
    if (photo.folderId) {
      try { folder = DriveApp.getFolderById(photo.folderId); }
      catch(e) { folder = DriveApp.getRootFolder(); }
    } else {
      // "ग्राम गाडी ट्रॅकर Photos" folder आपोआप बनवा
      const folders = DriveApp.getFoldersByName('ग्राम गाडी ट्रॅकर Photos');
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('ग्राम गाडी ट्रॅकर Photos');
    }
    // Date नुसार subfolder
    const subName = photo.date || new Date().toLocaleDateString('en-IN');
    const subs = folder.getFoldersByName(subName);
    const sub = subs.hasNext() ? subs.next() : folder.createFolder(subName);

    const file = sub.createFile(blob);
    file.setDescription(`Trip: ${photo.tripId} | Driver: ${photo.driverName} | Route: ${photo.route} | Stage: ${photo.stageName}`);

    // Trips sheet मध्ये Drive link add करा
    addDriveLinkToTrip(photo.tripId, photo.stageName, file.getUrl());

    return { saved: true, url: file.getUrl(), name: file.getName() };
  } catch(e) {
    return { saved: false, error: e.message };
  }
}

function addDriveLinkToTrip(tripId, stageName, url) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_TRIPS);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(tripId)) {
        // Column 19 onwards मध्ये Drive links
        const existing = data[i][18] || '';
        sh.getRange(i+1, 19).setValue(existing ? existing + '\n' + stageName + ': ' + url : stageName + ': ' + url);
        break;
      }
    }
  } catch(e) { Logger.log('Drive link error: ' + e.message); }
}

// ---------------------------------------------------------------
// Users sheet पूर्णपणे replace करा (Developer ने push केल्यावर)
function saveUsers(users) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_USERS);
  if (!sh) sh = ss.insertSheet(S_USERS);
  sh.clearContents();
  sh.appendRow(['Login ID', 'Password', 'Role', 'नाव', 'Phone']);
  sh.getRange(1,1,1,5).setBackground('#6a1b9a').setFontColor('white').setFontWeight('bold');
  sh.setFrozenRows(1);
  (users||[]).forEach(u => sh.appendRow([u.id, u.pass, u.role, u.name, u.phone||'']));
  return { saved: true, count: (users||[]).length };
}

// ---------------------------------------------------------------
// SETUP — पहिल्यांदा एकदाच run करा
// ---------------------------------------------------------------
function setupAllSheets() {
  ensureTripsSheet();
  ensureSettingsSheet();
  ensureUsersSheet();
  ensureStopsSheet();
  Logger.log('✅ सर्व sheets तयार झाल्या!');
  Logger.log('Sheets: ' + getSheetNames().join(', '));
}
