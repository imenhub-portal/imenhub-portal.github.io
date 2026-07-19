// Last updated by Claude
// --- CONFIGURATION ---
// ⬇️ UPDATED: Hardcoded URL to prevent access errors ⬇️
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw1PrhjSzHsd__rxbEcXZp4iyF4v9e41pVYg1j7ewo6ocnPNw3CsdTt-j7UFb0NM-0l/exec";

// Student-facing portal (GitHub Pages copy of the UI — no Google banner, no
// authuser issues; ?lab= deep-link handled by index.html's URLSearchParams fallback)
const PORTAL_URL = "https://imenhub-portal.github.io/i-nstrumen/";

const SHEET_IDS = {
  EQUIPMENT: 'Equipment',
  BOOKINGS: 'Bookings',
  LOGS: 'Logs',
  CONFIG: 'Config'
};

// --- SERVING THE APP ---
function doGet(e) {
  // JSON API mode — used by external tools for live sync
  if (e && e.parameter && e.parameter.format === 'json') {
    const data = getSystemData();
    return ContentService
      .createTextOutput(JSON.stringify({ equipment: data.equipment, labs: data.labs }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Equipment Glossary page
  if (e && e.parameter && e.parameter.page === 'glossary') {
    return HtmlService.createHtmlOutputFromFile('ijana')
      .setTitle('i-Nstrumen — Equipment Glossary')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Normal HTML app
  const params = {
    lab: (e && e.parameter && e.parameter.lab) || null,
    view: (e && e.parameter && e.parameter.view) || null
  };

  // ── SINGLE-SOURCE UI ─────────────────────────────────────────────────────
  // The UI is maintained in ONE place only: the GitHub repo. This fetches the
  // latest index.html from GitHub and serves it, so UI updates never need to
  // be pasted into Apps Script again. If GitHub is unreachable, we fall back
  // to the last locally-pasted 'index' file (emergency copy).
  const UI_SOURCES = [
    'https://raw.githubusercontent.com/imenhub-portal/imenhub-portal.github.io/main/i-nstrumen/index.html',
    'https://raw.githubusercontent.com/imenhub-portal/imenhub-portal.github.io/master/i-nstrumen/index.html'
  ];
  for (let i = 0; i < UI_SOURCES.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(UI_SOURCES[i], { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        // Inject server params by replacing the template scriptlet literally
        // (the GitHub copy contains the raw scriptlet text).
        const html = resp.getContentText()
          .replace('<?!= JSON.stringify(urlParams) ?>', JSON.stringify(params));
        return HtmlService.createHtmlOutput(html)
          .setTitle('IMEN Lab Manager')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (err) { /* try next source / fall back */ }
  }

  // Fallback: last pasted local copy (only used if GitHub is unreachable)
  const template = HtmlService.createTemplateFromFile('index');
  template.urlParams = params;
  return template.evaluate()
    .setTitle('IMEN Lab Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- JSON API (GitHub-Pages hosted UI) ---
// The frontend shim in index.html POSTs {fn, args} as text/plain (a CORS
// "simple request" — no preflight, which Apps Script cannot answer).
// SECURITY: dispatch is an explicit switch — ONLY these three entry points are
// callable from outside; arbitrary function names are rejected.
function doPost(e) {
  let payload;
  try {
    const req  = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = Array.isArray(req.args) ? req.args : [];
    let result;
    switch (req.fn) {
      case 'getInitialData':       result = getInitialData(); break;
      case 'handleFrontendAction': result = handleFrontendAction(args[0], args[1]); break;
      case 'fetchUpdates':         result = fetchUpdates(args[0]); break;
      default: throw new Error('Unknown API function: ' + req.fn);
    }
    payload = { ok: true, result: result };
  } catch (err) {
    payload = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
                       .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 1. PUBLIC API
// ==========================================

function getInitialData() {
  return getSystemData(); 
}

function handleFrontendAction(actionType, payload) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(5000);
  if (!acquired) {
    return { success: false, error: "Server busy. Please try again in a moment." };
  }
  
  try {
    switch (actionType) {
      case 'Usage': return saveLog(payload);
      case 'Book': return saveBooking(payload);
      case 'UpdateBooking': return updateBooking(payload); 
      case 'Report': {
        const result = saveLog(payload.log);
        // Apply explicit equipment update (supports partial process-level maintenance)
        if (payload.equipmentUpdate) {
          updateEquipmentStatus(payload.equipmentUpdate.id, payload.equipmentUpdate.status, payload.equipmentUpdate.maintenanceReason);
        }
        return result;
      }
      case 'EditEquipment': return saveEquipment(payload);
      case 'DeleteEquipment': return deleteEquipment(payload.id);
      case 'EditLab': return renameLab(payload.oldName, payload.newName);
      case 'AddLab': return addLab(payload.name);
      case 'DeleteLab': return deleteLab(payload.name);
      case 'EditCoordinator':
        return saveConfig({
            coordinators: payload.allCoordinators,
            officialEmail: payload.officialEmail
        });
      case 'EditTechStaff':
        return saveConfig({ techStaff: payload.allTechStaff });
      case 'Archive': return archiveSystem();
      case 'CancelBooking': return cancelBooking(payload);
      case 'FindMyBookings': return findMyBookings(payload);
      case 'GetSmartMatch': return getSmartMatchEquipment(payload);

      // ── i-Menian Crossbridge ────────────────────────────────────────────────
      case 'SearchIMenian': {
        var IMENIAN_ID  = '16mqyApWABuMmYUumLXOLsAzVLwwu9V4MlQs73RYZgNY';
        var labFilter   = (payload.lab   || '').trim().toUpperCase();
        var queryFilter = (payload.query || '').trim().toLowerCase();
        var cacheKey    = 'imenian_' + labFilter.replace(/\W/g, '_');
        var cache       = CacheService.getScriptCache();

        // ── Return cached list if available (avoids re-reading the sheet) ──
        var cached = cache.get(cacheKey);
        if (cached) {
          var all = JSON.parse(cached);
          if (!queryFilter) return all;
          return all.filter(function(u) {
            return u.name.toLowerCase().indexOf(queryFilter) !== -1;
          });
        }

        // ── First load: read only the columns we need (skip Photo column) ──
        var iSS    = SpreadsheetApp.openById(IMENIAN_ID);
        var iSheet = iSS.getSheets()[0];
        var iHdr   = iSheet.getRange(1, 1, 1, iSheet.getLastColumn()).getValues()[0];

        var iName  = iHdr.indexOf('Name');
        var iEmail = iHdr.indexOf('Email');
        var iPhone = iHdr.indexOf('Phone');
        var iMatric= iHdr.indexOf('StudentID');
        var iCat   = iHdr.indexOf('Category');
        var iSup   = iHdr.indexOf('SupervisorName');
        var iLab   = iHdr.indexOf('LabName');
        var iPhoto = iHdr.indexOf('Photo');

        // Read all needed columns in ONE call, excluding nothing — but
        // we purposely skip the Photo column by not including it in results
        // (Photo is fetched separately on demand via GetIMenianPhoto)
        var lastRow  = iSheet.getLastRow();
        var lastCol  = iSheet.getLastColumn();
        var allData  = iSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

        var seen    = {};
        var results = [];

        for (var r = 0; r < allData.length; r++) {
          var row      = allData[r];
          var rowLab   = (row[iLab]  || '').toString().trim().toUpperCase();
          var rowName  = (row[iName] || '').toString().trim();
          var rowEmail = (row[iEmail]|| '').toString().trim().toLowerCase();

          if (rowLab !== labFilter) continue;
          if (!rowName)             continue;
          if (seen[rowEmail])       continue;

          seen[rowEmail] = true;
          results.push({
            name:       rowName,
            email:      rowEmail,
            phone:      (row[iPhone]  || '').toString(),
            matric:     (row[iMatric] || '').toString(),
            category:   (row[iCat]   || '').toString(),
            supervisor: (row[iSup]   || '').toString()
          });
        }

        // Cache the full lab list for 30 minutes
        try { cache.put(cacheKey, JSON.stringify(results), 1800); } catch(e) {}

        if (!queryFilter) return results;
        return results.filter(function(u) {
          return u.name.toLowerCase().indexOf(queryFilter) !== -1;
        });
      }

      case 'GetIMenianPhoto': {
        // Optimisation: read ONLY the StudentID and Photo columns — not the entire sheet.
        // This avoids loading every other user's large base64 image into memory.
        var IMENIAN_ID2  = '16mqyApWABuMmYUumLXOLsAzVLwwu9V4MlQs73RYZgNY';
        var targetMatric = (payload.matric || '').toString().trim();

        // Check photo cache first (keyed by matric)
        var pCache    = CacheService.getScriptCache();
        var pCacheKey = 'imenian_photo_' + targetMatric;
        var pCached   = pCache.get(pCacheKey);
        if (pCached) return pCached === 'NONE' ? null : pCached;

        var pSS    = SpreadsheetApp.openById(IMENIAN_ID2);
        var pSheet = pSS.getSheets()[0];
        var pHdr   = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];

        var pMatricCol = pHdr.indexOf('StudentID') + 1; // getRange is 1-based
        var pPhotoCol  = pHdr.indexOf('Photo')     + 1;
        var pLastRow   = pSheet.getLastRow();

        // Read ONLY the two columns we need
        var matricVals = pSheet.getRange(2, pMatricCol, pLastRow - 1, 1).getValues();
        var photoVals  = pSheet.getRange(2, pPhotoCol,  pLastRow - 1, 1).getValues();

        for (var p = 0; p < matricVals.length; p++) {
          if ((matricVals[p][0] || '').toString().trim() === targetMatric) {
            var photo = (photoVals[p][0] || '').toString();
            if (photo.startsWith('data:image')) {
              // Cache photo for 30 min — skip if > 90 KB (CacheService limit)
              try {
                if (photo.length < 90000) pCache.put(pCacheKey, photo, 1800);
              } catch(e) {}
              return photo;
            }
            pCache.put(pCacheKey, 'NONE', 1800);
            return null;
          }
        }
        pCache.put(pCacheKey, 'NONE', 1800);
        return null;
      }

      case 'GetIMenianPhotoBatch': {
        // Returns {matric: base64DataUrl} for every member in the given lab.
        // One sheet open, one read — result cached per-matric so individual
        // GetIMenianPhoto calls also benefit from the warm cache.
        var IMENIAN_ID3 = '16mqyApWABuMmYUumLXOLsAzVLwwu9V4MlQs73RYZgNY';
        var bLabFilter  = (payload.lab || '').trim().toUpperCase();
        var bMatrics    = payload.matrics || []; // array of matric strings to include

        var bSS    = SpreadsheetApp.openById(IMENIAN_ID3);
        var bSheet = bSS.getSheets()[0];
        var bHdr   = bSheet.getRange(1, 1, 1, bSheet.getLastColumn()).getValues()[0];

        var bMatricCol = bHdr.indexOf('StudentID') + 1;
        var bPhotoCol  = bHdr.indexOf('Photo')     + 1;
        var bLabCol    = bHdr.indexOf('LabName')   + 1;
        var bLastRow   = bSheet.getLastRow();

        if (bMatricCol < 1 || bPhotoCol < 1 || bLastRow < 2) return {};

        var bMatricVals = bSheet.getRange(2, bMatricCol, bLastRow - 1, 1).getValues();
        var bPhotoVals  = bSheet.getRange(2, bPhotoCol,  bLastRow - 1, 1).getValues();
        var bLabVals    = bSheet.getRange(2, bLabCol,    bLastRow - 1, 1).getValues();

        var bMatricSet  = {};
        bMatrics.forEach(function(m) { if (m) bMatricSet[String(m).trim()] = true; });

        var bCache  = CacheService.getScriptCache();
        var bResult = {};

        for (var b = 0; b < bMatricVals.length; b++) {
          var bM   = String(bMatricVals[b][0] || '').trim();
          var bLab = String(bLabVals[b][0]    || '').trim().toUpperCase();
          var bPic = String(bPhotoVals[b][0]  || '').trim();

          if (bLab !== bLabFilter)          continue;
          if (!bM || !bMatricSet[bM])       continue;
          if (!bPic.startsWith('data:image')) continue;

          bResult[bM] = bPic;

          // Warm the individual photo cache so GetIMenianPhoto hits are instant too
          try {
            var bKey = 'imenian_photo_' + bM;
            if (!bCache.get(bKey) && bPic.length < 90000) {
              bCache.put(bKey, bPic, 1800);
            }
          } catch(e) {}
        }

        return bResult;
      }

      case 'GetAllIMenianUsers': {
        // Returns ALL users across every lab — no lab filter, no photos.
        // Used for cross-lab name search. Deduplicated by email.
        var allCacheKey = 'imenian_all_users';
        var allCache    = CacheService.getScriptCache();
        var allCached   = allCache.get(allCacheKey);
        if (allCached) return JSON.parse(allCached);

        var aSS    = SpreadsheetApp.openById('16mqyApWABuMmYUumLXOLsAzVLwwu9V4MlQs73RYZgNY');
        var aSheet = aSS.getSheets()[0];
        var aHdr   = aSheet.getRange(1, 1, 1, aSheet.getLastColumn()).getValues()[0];

        var aN = aHdr.indexOf('Name');
        var aE = aHdr.indexOf('Email');
        var aP = aHdr.indexOf('Phone');
        var aM = aHdr.indexOf('StudentID');
        var aC = aHdr.indexOf('Category');
        var aS = aHdr.indexOf('SupervisorName');
        var aL = aHdr.indexOf('LabName');

        var aLastRow = aSheet.getLastRow();
        var aLastCol = aSheet.getLastColumn();
        var aAll     = aSheet.getRange(2, 1, aLastRow - 1, aLastCol).getValues();

        var aSeen    = {};
        var aResults = [];

        for (var ai = 0; ai < aAll.length; ai++) {
          var aRow = aAll[ai];
          var aNm  = (aRow[aN] || '').toString().trim();
          var aEm  = (aRow[aE] || '').toString().trim().toLowerCase();
          if (!aNm || !aEm) continue;
          if (aSeen[aEm])   continue; // deduplicate by email
          aSeen[aEm] = true;
          aResults.push({
            name:       aNm,
            email:      aEm,
            phone:      (aRow[aP] || '').toString(),
            matric:     (aRow[aM] || '').toString(),
            category:   (aRow[aC] || '').toString(),
            supervisor: (aRow[aS] || '').toString(),
            lab:        (aRow[aL] || '').toString().trim()
          });
        }

        try { allCache.put(allCacheKey, JSON.stringify(aResults), 1800); } catch(e) {}
        return aResults;
      }

      case 'GetLabData': {
        // ONE spreadsheet open → returns BOTH members and photos for a lab.
        // This eliminates the second GAS call that was causing 4-6s total load time.
        var gLabFilter = (payload.lab || '').trim().toUpperCase();
        var gCacheKey  = 'imenian_' + gLabFilter.replace(/\W/g, '_');
        var gCache     = CacheService.getScriptCache();

        var gSS    = SpreadsheetApp.openById('16mqyApWABuMmYUumLXOLsAzVLwwu9V4MlQs73RYZgNY');
        var gSheet = gSS.getSheets()[0];
        var gHdr   = gSheet.getRange(1, 1, 1, gSheet.getLastColumn()).getValues()[0];

        var gName  = gHdr.indexOf('Name');
        var gEmail = gHdr.indexOf('Email');
        var gPhone = gHdr.indexOf('Phone');
        var gMatric= gHdr.indexOf('StudentID');
        var gCat   = gHdr.indexOf('Category');
        var gSup   = gHdr.indexOf('SupervisorName');
        var gLab   = gHdr.indexOf('LabName');
        var gPhoto = gHdr.indexOf('Photo');

        var gLastRow = gSheet.getLastRow();
        var gLastCol = gSheet.getLastColumn();
        var gAll     = gSheet.getRange(2, 1, gLastRow - 1, gLastCol).getValues();

        var gSeen    = {};
        var gMembers = [];
        var gPhotos  = {};  // {matric: base64}

        for (var g = 0; g < gAll.length; g++) {
          var gRow    = gAll[g];
          var gRowLab = (gRow[gLab]   || '').toString().trim().toUpperCase();
          var gRowNm  = (gRow[gName]  || '').toString().trim();
          var gRowEm  = (gRow[gEmail] || '').toString().trim().toLowerCase();
          var gRowMat = (gRow[gMatric]|| '').toString().trim();
          var gRowPic = gPhoto >= 0 ? (gRow[gPhoto] || '').toString().trim() : '';

          if (gRowLab !== gLabFilter) continue;
          if (!gRowNm)                continue;

          // Collect photos (before dedup — same matric may appear in multiple labs)
          if (gRowMat && gRowPic.startsWith('data:image') && !gPhotos[gRowMat]) {
            gPhotos[gRowMat] = gRowPic;
          }

          if (gSeen[gRowEm]) continue;
          gSeen[gRowEm] = true;

          gMembers.push({
            name:       gRowNm,
            email:      gRowEm,
            phone:      (gRow[gPhone] || '').toString(),
            matric:     gRowMat,
            category:   (gRow[gCat]  || '').toString(),
            supervisor: (gRow[gSup]  || '').toString()
          });
        }

        // Cache member list for SearchIMenian (per-keystroke fallback)
        try { gCache.put(gCacheKey, JSON.stringify(gMembers), 1800); } catch(e) {}

        return { members: gMembers, photos: gPhotos };
      }
      // ── End i-Menian Crossbridge ────────────────────────────────────────────

      default: return { success: false, error: "Unknown Action Type: " + actionType };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function fetchUpdates(lastClientTimestamp) {
  if (!lastClientTimestamp) return { newLogs: [], updatedBookings: [] };
  const lastTime = new Date(lastClientTimestamp).getTime();
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.LOGS);
  const recentLogs = getLastNRowsAsObjects(logSheet, 50);
  const newLogs = recentLogs.filter(l => new Date(l.timestamp).getTime() > lastTime);
  
  const bookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.BOOKINGS);
  const allBookings = getDataAsObjects(bookSheet);
  const updatedBookings = allBookings.filter(b => {
      const created = b.timestampCreated ? new Date(b.timestampCreated).getTime() : 0;
      const actioned = b.timestampActioned ? new Date(b.timestampActioned).getTime() : 0;
      return created > lastTime || actioned > lastTime;
  });
  
  return { newLogs, updatedBookings };
}

// ==========================================
// 2. DATA ACCESS LAYER
// ==========================================

function getSystemData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eqData = getDataAsObjects(ss.getSheetByName(SHEET_IDS.EQUIPMENT));
  eqData.forEach(eq => {
    try { eq.processCapabilities = JSON.parse(eq.processCapabilities || '[]'); } catch(e) { eq.processCapabilities = []; }
    try { eq.materialsOptions = JSON.parse(eq.materialsOptions || '[]'); } catch(e) { eq.materialsOptions = []; }
  });

  const bookData = getDataAsObjects(ss.getSheetByName(SHEET_IDS.BOOKINGS));
  const logData = getLastNRowsAsObjects(ss.getSheetByName(SHEET_IDS.LOGS), 5000);
  
  const configSheet = ss.getSheetByName(SHEET_IDS.CONFIG);
  const rawLabs = configSheet.getRange("A2:A").getValues().flat();
  const labs = rawLabs.filter(l => l && l.toString().trim() !== '');

  const coordJson = configSheet.getRange("C2").getValue();
  let coordinators = [];
  try { coordinators = JSON.parse(coordJson); } catch (e) { coordinators = []; }

  // NEW — Technical Staff (global, all labs)
  const techStaffJson = configSheet.getRange("D2").getValue();
  let techStaff = [];
  try { techStaff = JSON.parse(techStaffJson); } catch (e) { techStaff = []; }

  const rawSettings = configSheet.getRange("E2:F").getValues();
  const systemConfig = { officialEmail: 'imenmakmal@imen.ukm.edu.my', coordinators: coordinators };
  let lastArchive = null;

  rawSettings.forEach(row => {
    if (row[0] === 'officialEmail' && row[1]) systemConfig.officialEmail = row[1];
    if (row[0] === 'lastArchive' && row[1]) lastArchive = row[1];
  });

  // ⬇️ UPDATED: Uses the Hardcoded URL
  const appUrl = WEB_APP_URL; 

  return {
    equipment: eqData, bookings: bookData, logs: logData, labs: labs,
    config: {
      labs: labs,
      settings: systemConfig,
      coordinators: coordinators,
      techStaff: techStaff          // NEW
    },
    lastArchive: lastArchive,
    appUrl: appUrl
  };
}

// ==========================================
// 3. WRITERS & EMAIL LOGIC
// ==========================================

function saveLog(logObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.LOGS);
  // Auto-heal: ensure the userEmail header exists (col 16 / P). Added Jul 2026 —
  // the frontend always sent userEmail but it was never persisted, which weakened
  // no-show email-matching and distinct-user counts. Old rows simply stay blank.
  if (sheet.getRange(1, 16).getValue() === '') sheet.getRange(1, 16).setValue('userEmail');
  const row = [
    logObj.id, logObj.timestamp, logObj.equipmentId, logObj.equipmentName, logObj.lab,
    logObj.userName, logObj.affiliation, logObj.action, logObj.duration,
    logObj.samples, logObj.materials, logObj.issueDetails, logObj.sessionEnded || false,
    logObj.paymentType || '', logObj.paymentRef || '',
    logObj.userEmail || ''
  ];
  sheet.appendRow(row);
  
  // --- ISSUE REPORTING LOGIC ---
  if (logObj.action === 'Report' || logObj.action === 'Flagged') {
      
      // 1. Update Equipment Status in Inventory
      if (logObj.equipmentId) {
          const status = logObj.action === 'Report' ? 'Maintenance' : 'Attention';
          updateEquipmentStatus(logObj.equipmentId, status, logObj.issueDetails);
      }

      // 2. Extract Reporter Email (Format: "Name (email@domain.com)")
      let reporterEmail = null;
      try {
        const match = logObj.userName.match(/\(([^)]+)\)/);
        if (match) reporterEmail = match[1];
      } catch(e) {}

      // 3. Prepare Email Content
      const isCritical = logObj.action === 'Report';
      const color = isCritical ? '#dc2626' : '#d97706'; // Red or Amber
      const title = isCritical ? 'Equipment Breakdown Reported' : 'Equipment Issue Flagged';
      
      const emailBody = `
        <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
          <div style="background-color: ${color}; color: white; padding: 20px;">
            <h2 style="margin: 0;">${title}</h2>
            <p style="margin: 5px 0 0 0; opacity: 0.9;">Urgent Attention Required</p>
          </div>
          <div style="padding: 20px;">
            <p>An issue has been reported affecting the following equipment:</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666; width: 30%;"><strong>Equipment</strong></td><td style="padding: 10px; font-size: 16px;">${logObj.equipmentName}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Lab</strong></td><td style="padding: 10px;">${logObj.lab}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Reported By</strong></td><td style="padding: 10px;">${logObj.userName}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Time</strong></td><td style="padding: 10px;">${new Date(logObj.timestamp).toLocaleString()}</td></tr>
              <tr><td style="padding: 10px; color: #666; vertical-align: top;"><strong>Details</strong></td><td style="padding: 10px; color: ${color}; font-weight: bold;">${logObj.issueDetails}</td></tr>
            </table>
            
            <div style="margin-top: 20px; padding: 15px; background-color: #fff1f2; border-left: 4px solid ${color}; color: #881337; font-size: 13px;">
              <strong>System Action:</strong> Equipment status has been automatically updated to <em>${isCritical ? 'Maintenance' : 'Attention'}</em>.
            </div>
          </div>
          <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
            IMEN Lab Management System
          </div>
        </div>`;

      // 4. Send Email to Coordinator
      const coordEmail = getCoordinatorEmail(logObj.lab);
      if (coordEmail) {
          sendEmailSafe(coordEmail, `[IMEN ALERT] ${logObj.equipmentName} - ${logObj.action}`, emailBody);
      }

      // 5. Send Receipt to Student (Reporter)
      const allTechStaff = getTechStaffEmails();
      allTechStaff.forEach(function(techEmail) {
          // Avoid double-emailing if a tech staff happens to also be the PIC
          if (techEmail && techEmail !== coordEmail) {
              sendEmailSafe(techEmail, `[IMEN ALERT] ${logObj.equipmentName} - ${logObj.action}`, emailBody);
          }
      });

      // 6. Send Receipt to Student (Reporter)
      if (reporterEmail) {
          const studentBody = emailBody
              .replace('Urgent Attention Required', 'Receipt of Report')
              .replace('An issue has been reported', 'Thank you. We have received your report regarding:');
          sendEmailSafe(reporterEmail, `[IMEN] Report Received: ${logObj.equipmentName}`, studentBody);
      }
  }

  // --- USAGE / WALK-IN NOTIFICATION ---
  if (logObj.action === 'Usage' && logObj.userEmail) {
      const coordEmail = getCoordinatorEmail(logObj.lab);
      const usageTime  = new Date(logObj.timestamp).toLocaleString('en-MY', {
          timeZone: 'Asia/Kuala_Lumpur',
          dateStyle: 'full',
          timeStyle: 'short'
      });

      const usageBody = `
        <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #4f46e5; color: white; padding: 20px;">
            <h2 style="margin: 0;">Session Started</h2>
            <p style="margin: 5px 0 0 0; opacity: 0.9;">i-Nstrumen Activity Confirmation</p>
          </div>
          <div style="padding: 20px;">
            <p style="color:#555;">A usage session has been logged under your name. Details are as follows:</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666; width: 35%;"><strong>Equipment</strong></td><td style="padding: 10px; font-size: 15px; font-weight: bold;">${logObj.equipmentName}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Lab</strong></td><td style="padding: 10px;">${logObj.lab}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Name</strong></td><td style="padding: 10px;">${logObj.userName}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Duration</strong></td><td style="padding: 10px;">${logObj.duration}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Materials</strong></td><td style="padding: 10px;">${logObj.materials || '-'}</td></tr>
              <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Payment</strong></td><td style="padding: 10px;">${logObj.paymentType || '-'}${logObj.paymentRef ? ' &mdash; <span style="font-family:monospace;">' + logObj.paymentRef + '</span>' : ''}</td></tr>
              <tr><td style="padding: 10px; color: #666;"><strong>Time</strong></td><td style="padding: 10px;">${usageTime}</td></tr>
            </table>
            <div style="margin-top: 20px; padding: 15px; background-color: #fef9c3; border-left: 4px solid #ca8a04; color: #713f12; font-size: 13px;">
              <strong>&#9888; Not you?</strong> If you did not initiate this session, please contact your lab coordinator immediately${coordEmail ? ' at <a href="mailto:' + coordEmail + '" style="color:#4f46e5;">' + coordEmail + '</a>' : ''}.
            </div>
          </div>
          <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
            Generated automatically by i-Nstrumen &mdash; IMEN Lab Management System
          </div>
        </div>`;

      sendEmailSafe(logObj.userEmail, `[i-Nstrumen] Session Started: ${logObj.equipmentName}`, usageBody);
  }

  return { success: true };
}

function saveBooking(bookingObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.BOOKINGS);
  const existingBookings = getDataAsObjects(sheet);
  
  // --- SMART SLOT CONFLICT CHECK ---
  const hasConflict = existingBookings.some(b => {
      // 1. Basic Filters
      if (b.equipmentName !== bookingObj.equipmentName) return false;
      if (b.lab !== bookingObj.lab) return false; 
      if (b.status !== 'Approved') return false; 

      // 2. Normalize Dates (YYYY-MM-DD)
      // We compare purely based on Date Strings to avoid Timezone headaches
      const reqDate = bookingObj.date.split('T')[0];
      const bDate = b.date.split('T')[0];

      // Handle "2 Days" Logic
      let bDates = [bDate];
      if (b.duration === '2 Days') {
          const d = new Date(bDate); d.setDate(d.getDate() + 1);
          bDates.push(d.toISOString().split('T')[0]);
      }

      let reqDates = [reqDate];
      if (bookingObj.duration === '2 Days') {
          const d = new Date(reqDate); d.setDate(d.getDate() + 1);
          reqDates.push(d.toISOString().split('T')[0]);
      }

      // // 3. Do Dates Overlap?
      const dateMatch = reqDates.some(r => bDates.includes(r));
      if (!dateMatch) return false; // Different days, no conflict.

      // 4. IF Dates Match, Check Slots
      // A. "Day" hogs: If either is "1 Day" or "2 Days", it blocks everything
      if (bookingObj.duration.includes('Day') || b.duration.includes('Day')) {
          return true; 
      }

      // B. Sub-day slots: Only conflict if they are EXACTLY the same
      // e.g. "8am-1pm" vs "8am-1pm" = Conflict
      // e.g. "8am-1pm" vs "2pm-5pm" = OK
      if (bookingObj.duration === b.duration) {
          return true;
      }

      return false; // Different slots on same day -> ALLOWED
  });

  if (hasConflict) {
      return { success: false, error: "Booking Failed: Selected slot is not available." };
  }
  // ----------------------

  // ... (Rest of the save logic remains exactly the same) ...
  const row = [
      bookingObj.id, bookingObj.equipmentName, bookingObj.lab, bookingObj.userName,
      bookingObj.userEmail, "'" + bookingObj.userPhone, bookingObj.userId, bookingObj.affiliation,
      bookingObj.supervisor, bookingObj.date, bookingObj.duration, bookingObj.samples,
      bookingObj.materials, bookingObj.process, bookingObj.variant, bookingObj.status,
      bookingObj.remarks, bookingObj.timestampCreated, bookingObj.timestampActioned,
      bookingObj.paymentType || '',
      bookingObj.paymentRef  || ''
  ];
  sheet.appendRow(row);

  const coordEmail = getCoordinatorEmail(bookingObj.lab);
  const adminLink = `${WEB_APP_URL}?lab=${encodeURIComponent(bookingObj.lab)}&view=admin`;

  // --- HTML TEMPLATE GENERATOR ---
  const createTable = (title, color) => {
    // Build payment display string
    const payType = bookingObj.paymentType || 'Not specified';
    const payRef  = bookingObj.paymentRef  || '-';
    let payDisplay = payType;
    if (payType === 'Grant') payDisplay = `Grant &mdash; Ref: <strong>${payRef}</strong>`;
    if (payType === 'Cash')  payDisplay = `Cash &mdash; Receipt: <strong>${payRef}</strong>`;

    return `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${color}; color: white; padding: 20px;">
        <h2 style="margin: 0;">${title}</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9;">IMEN Lab Management System</p>
      </div>
      <div style="padding: 20px;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Equipment</strong></td><td style="padding: 10px;">${bookingObj.equipmentName}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Date &amp; Slot</strong></td><td style="padding: 10px;">${bookingObj.date} (${bookingObj.duration})</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>User Name</strong></td><td style="padding: 10px;">${bookingObj.userName} (${bookingObj.userId})</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Phone</strong></td><td style="padding: 10px;">${bookingObj.userPhone}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Affiliation</strong></td><td style="padding: 10px;">${bookingObj.affiliation}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Supervisor</strong></td><td style="padding: 10px;">${bookingObj.supervisor}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Process</strong></td><td style="padding: 10px;">${bookingObj.process || '-'} &rsaquo; ${bookingObj.variant || '-'}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Samples</strong></td><td style="padding: 10px;">${bookingObj.samples}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Materials</strong></td><td style="padding: 10px;">${bookingObj.materials}</td></tr>
          <tr style="background-color: #f0fdf4;">
            <td style="padding: 10px; color: #666;"><strong>&#128179; Payment</strong></td>
            <td style="padding: 10px; color: #166534;">${payDisplay}</td>
          </tr>
        </table>
      </div>
      <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
        Generated automatically by i-Nstrumen
      </div>
    </div>`;
  };

  // 1. Email to Coordinator
  if (coordEmail) {
      const body = createTable(`New Booking Request`, '#4f46e5') + 
        `<br><div style="text-align:center"><a href="${adminLink}" style="background-color:#4f46e5; color:white; padding:12px 24px; text-decoration:none; border-radius:5px; font-weight:bold; font-family: sans-serif;">Manage in Admin Panel</a></div>`;
      
      sendEmailSafe(coordEmail, `[IMEN] Action Required: ${bookingObj.equipmentName} (${bookingObj.userName})`, body);
  }
  
  // 2. Email to User (Receipt)
  if(bookingObj.userEmail) {
      const body = createTable(`Booking Received: Pending`, '#6b7280') + 
        `<p style="text-align:center; color:#666; font-family:sans-serif;">Your request is pending approval from the lab coordinator.</p>`;
        
      sendEmailSafe(bookingObj.userEmail, `[IMEN] Booking Received: ${bookingObj.equipmentName}`, body);
  }

  return { success: true };
}

function updateBooking(idOrObj, status) {
  let id = typeof idOrObj === 'object' ? idOrObj.id : idOrObj;
  let newStatus = typeof idOrObj === 'object' ? idOrObj.status : status;
  let remarks = typeof idOrObj === 'object' ? idOrObj.remarks : ''; 

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const searchId = String(id).trim();

  // --- CONFLICT CHECK ---
  if (newStatus === 'Approved') {
      const allBookings = getDataAsObjects(sheet);
      const targetBooking = allBookings.find(b => String(b.id).trim() === searchId);
      
      if (targetBooking) {
           const reqStart = new Date(targetBooking.date); reqStart.setHours(0,0,0,0);
           const reqEnd = new Date(reqStart); reqEnd.setHours(0,0,0,0);
           if (targetBooking.duration && targetBooking.duration.includes('2 Days')) {
               reqEnd.setDate(reqEnd.getDate() + 1);
           }
           
           const hasConflict = allBookings.some(b => {
              if (String(b.id).trim() === searchId) return false; 
              if (b.equipmentName !== targetBooking.equipmentName) return false;
              if (b.lab !== targetBooking.lab) return false; 
              if (b.status !== 'Approved') return false; 
              
              const bStart = new Date(b.date); bStart.setHours(0,0,0,0);
              const bEnd = new Date(bStart); bEnd.setHours(0,0,0,0);
              if (b.duration && b.duration.includes('2 Days')) {
                  bEnd.setDate(bEnd.getDate() + 1);
              }
              
              // 1. Check Date Overlap
              const dateOverlap = (reqStart <= bEnd && reqEnd >= bStart);
              if (!dateOverlap) return false;

              // 2. If Dates Overlap, Check Slots
              // If either is a "Day" booking (Full Day / 2 Days), it blocks everything
              if (targetBooking.duration.includes('Day') || b.duration.includes('Day')) return true;

              // If exact match (e.g. 8am-1pm vs 8am-1pm), it blocks
              if (targetBooking.duration === b.duration) return true;

              // Otherwise (e.g. 8am-1pm vs 2pm-5pm), ALLOW it
              return false;
           });
           
           if (hasConflict) {
               return { success: false, error: "Approval Failed: Overlap with another Approved booking." };
           }
      }
  }
  // -------------------------------------

  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();
    if (sheetId === searchId) { 
       sheet.getRange(i + 1, 16).setValue(newStatus); 
       if(remarks) sheet.getRange(i + 1, 17).setValue(remarks); 
       sheet.getRange(i + 1, 19).setValue(new Date().toISOString()); 
       
       const userEmail = data[i][4]; 
       const eqName = data[i][1];    
       const bookDateVal = data[i][9]; 
       
       if (userEmail) {
           let dateStr = bookDateVal;
           try { dateStr = new Date(bookDateVal).toDateString(); } catch(e) {}
           
           const isApproved = newStatus === 'Approved';
           const headerColor = isApproved ? '#059669' : '#dc2626'; // Green or Red
           const headerTitle = isApproved ? 'Booking Approved' : 'Booking Rejected';
           const userDetails = data[i][3]; // User Name
           const calUrl = isApproved ? _buildCalendarUrl(eqName, data[i][2], bookDateVal, data[i][10]) : null;

           const emailBody = `
            <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
              <div style="background-color: ${headerColor}; color: white; padding: 20px;">
                <h2 style="margin: 0;">${headerTitle}</h2>
                <p style="margin: 5px 0 0 0; opacity: 0.9;">${eqName}</p>
              </div>
              <div style="padding: 20px;">
                <p>Dear ${userDetails},</p>
                <p>Your booking request for <strong>${dateStr}</strong> has been updated.</p>
                
                ${remarks ? `<div style="background-color: #fff1f2; border-left: 4px solid #f43f5e; padding: 15px; margin: 15px 0; color: #881337;"><strong>Coordinator Remarks:</strong><br>${remarks}</div>` : ''}
                
                ${isApproved ? `<div style="background-color: #ecfdf5; border-left: 4px solid #10b981; padding: 15px; margin: 15px 0; color: #064e3b;"><strong>Next Steps:</strong><br>Please arrive on time and <strong>check in at the i-Nstrumen portal before using the equipment</strong>. Ensure you have the necessary safety gear and materials.</div>` : ''}

                ${calUrl ? `<div style="text-align:center; margin: 5px 0 15px;"><a href="${calUrl}" style="background-color:#059669; color:white; padding:10px 22px; text-decoration:none; border-radius:5px; font-weight:bold; font-family:sans-serif; font-size:13px;">&#128197; Add to Google Calendar</a></div>` : ''}

                <table style="width: 100%; margin-top: 20px; font-size: 13px; color: #555;">
                   <tr><td width="30%"><strong>Equipment:</strong></td><td>${eqName}</td></tr>
                   <tr><td><strong>Date:</strong></td><td>${dateStr}</td></tr>
                   <tr><td><strong>Status:</strong></td><td style="color:${headerColor}; font-weight:bold;">${newStatus}</td></tr>
                </table>
              </div>
              <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
                IMEN Lab Management System
              </div>
            </div>
           `;

           sendEmailSafe(userEmail, `[IMEN] Update: ${newStatus} - ${eqName}`, emailBody);
       }
       return { success: true };
    }
  }
  return { success: false, error: "ID not found" };
}

function cancelBooking(payload) {
  const id       = String(payload.bookingId || '').trim();
  const reason   = String(payload.reason   || 'No reason provided').trim();
  const cancelledAt = String(payload.cancelledTimestamp || new Date().toISOString()).trim();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.BOOKINGS);
  const data  = sheet.getDataRange().getValues();

  let targetRow = -1;
  let booking   = null;

  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();
    if (sheetId === id) {
      targetRow = i + 1;
      booking = {
        id:               String(data[i][0]  || '').trim(),
        equipmentName:    String(data[i][1]  || ''),
        lab:              String(data[i][2]  || ''),
        userName:         String(data[i][3]  || ''),
        userEmail:        String(data[i][4]  || ''),
        userPhone:        String(data[i][5]  || ''),
        userId:           String(data[i][6]  || ''),
        affiliation:      String(data[i][7]  || ''),
        supervisor:       String(data[i][8]  || ''),
        date:             String(data[i][9]  || ''),
        duration:         String(data[i][10] || ''),
        samples:          data[i][11],
        materials:        String(data[i][12] || ''),
        process:          String(data[i][13] || ''),
        variant:          String(data[i][14] || ''),
        status:           String(data[i][15] || ''),
        remarks:          String(data[i][16] || ''),
        timestampCreated: String(data[i][17] || ''),
        timestampActioned:String(data[i][18] || ''),
        paymentType:      String(data[i][19] || ''),
        paymentRef:       String(data[i][20] || '')
      };
      break;
    }
  }

  if (!booking) return { success: false, error: "Booking not found" };

  const currentStatus = booking.status;
  if (currentStatus !== 'Pending' && currentStatus !== 'Approved') {
    return { success: false, error: "Only Pending or Approved bookings can be cancelled. Current status: " + currentStatus };
  }

  // Validate slot hasn't ended (only for Approved bookings; Pending can always cancel)
  if (currentStatus === 'Approved') {
    const slotEnd = _getBookingSlotEnd(booking.date, booking.duration);
    const now = new Date();
    if (now > slotEnd) {
      return { success: false, error: "Cannot cancel — the booking slot has already ended." };
    }
  }

  // Update the booking row
  const cancelRemarks = 'Cancelled by user: ' + reason;
  sheet.getRange(targetRow, 16).setValue('Cancelled');
  sheet.getRange(targetRow, 17).setValue(cancelRemarks);
  sheet.getRange(targetRow, 19).setValue(cancelledAt);

  // Send notification email to PIC
  _sendCancelNotification(booking, reason, cancelledAt);

  return { success: true, booking: booking };
}

function _getBookingSlotEnd(dateStr, duration) {
  const raw = String(dateStr || '').split('T')[0];
  const [y, m, d] = raw.split('-').map(Number);
  if (!y || !m || !d) return new Date(0);
  const end = new Date(y, m - 1, d);
  if (duration === '8am-1pm')  end.setHours(13, 0, 0, 0);
  else if (duration === '2pm-5pm' || duration === '1pm-5pm') end.setHours(17, 0, 0, 0);
  else if (duration === '1 Day') end.setHours(17, 0, 0, 0);
  else if (duration === '2 Days') { end.setDate(end.getDate() + 1); end.setHours(17, 0, 0, 0); }
  else end.setHours(17, 0, 0, 0);
  return end;
}

function _sendCancelNotification(booking, reason, cancelledAt) {
  const coordEmail = getCoordinatorEmail(booking.lab);
  if (!coordEmail) return;

  const cancelTime = new Date(cancelledAt).toLocaleString('en-MY', {
    timeZone: 'Asia/Kuala_Lumpur',
    dateStyle: 'full',
    timeStyle: 'short'
  });

  const emailBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #6b7280; color: white; padding: 20px;">
        <h2 style="margin: 0;">Booking Cancelled by User</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9;">No action required — auto-approved cancellation</p>
      </div>
      <div style="padding: 20px;">
        <p>The following booking has been <strong>cancelled by the user</strong>. The slot is now free.</p>
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666; width: 30%;"><strong>Equipment</strong></td><td style="padding: 10px; font-size: 15px; font-weight: bold;">${booking.equipmentName}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Lab</strong></td><td style="padding: 10px;">${booking.lab}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>User</strong></td><td style="padding: 10px;">${booking.userName} (${booking.userId})</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Email</strong></td><td style="padding: 10px;">${booking.userEmail || '-'}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Phone</strong></td><td style="padding: 10px;">${booking.userPhone || '-'}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Date &amp; Slot</strong></td><td style="padding: 10px;">${booking.date} (${booking.duration})</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Process</strong></td><td style="padding: 10px;">${booking.process || '-'} ${booking.variant ? '> ' + booking.variant : ''}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Previous Status</strong></td><td style="padding: 10px;">${booking.status}</td></tr>
          <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Cancel Reason</strong></td><td style="padding: 10px; color: #dc2626; font-weight: bold;">${reason}</td></tr>
          <tr><td style="padding: 10px; color: #666;"><strong>Time</strong></td><td style="padding: 10px;">${cancelTime}</td></tr>
        </table>
        <div style="margin-top: 20px; padding: 15px; background-color: #f0fdf4; border-left: 4px solid #22c55e; border-radius: 4px; font-size: 13px; color: #166534;">
          <strong>Slot Released:</strong> The booking slot is now available for other researchers to book. No further action is needed from you.
        </div>
      </div>
      <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
        i-Nstrumen &mdash; IMEN Lab Management System
      </div>
    </div>`;

  const subject = `[i-Nstrumen] Booking Cancelled: ${booking.equipmentName} — ${booking.userName} (${booking.date})`;
  sendEmailSafe(coordEmail, subject, emailBody);
}

function findMyBookings(payload) {
  const phone = String(payload.phone || '').trim();

  const normalize = function(val) {
    if (!val) return '';
    var s = String(val).replace(/\D/g, '');
    if (s.indexOf('60') === 0) s = s.substring(2);
    if (s.indexOf('0')  === 0) s = s.substring(1);
    return s;
  };

  const searchPhone = normalize(phone);
  if (!searchPhone) return [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.BOOKINGS);
  const data  = sheet.getDataRange().getValues();
  const results = [];

  const now = new Date();

  for (var i = 1; i < data.length; i++) {
    const status = String(data[i][15] || '').trim();
    if (status !== 'Pending' && status !== 'Approved') continue;

    const rowPhone = normalize(String(data[i][5] || ''));
    if (rowPhone !== searchPhone) continue;

    // For Approved bookings, exclude if slot has already ended
    if (status === 'Approved') {
      const slotEnd = _getBookingSlotEnd(
        String(data[i][9] || ''),
        String(data[i][10] || '')
      );
      if (now > slotEnd) continue;
    }

    results.push({
      id:               String(data[i][0]  || '').trim(),
      equipmentName:    String(data[i][1]  || ''),
      lab:              String(data[i][2]  || ''),
      userName:         String(data[i][3]  || ''),
      userEmail:        String(data[i][4]  || ''),
      userPhone:        String(data[i][5]  || ''),
      userId:           String(data[i][6]  || ''),
      affiliation:      String(data[i][7]  || ''),
      supervisor:       String(data[i][8]  || ''),
      date:             String(data[i][9]  || ''),
      duration:         String(data[i][10] || ''),
      samples:          data[i][11],
      materials:        String(data[i][12] || ''),
      process:          String(data[i][13] || ''),
      variant:          String(data[i][14] || ''),
      status:           status,
      remarks:          String(data[i][16] || ''),
      timestampCreated: String(data[i][17] || ''),
      timestampActioned:String(data[i][18] || ''),
      paymentType:      String(data[i][19] || ''),
      paymentRef:       String(data[i][20] || '')
    });
  }

  return results;
}

// ==========================================
// AI SMART MATCH — multi-provider failover LLM router
// Providers: Gemini (primary) → Mistral → Groq (fallback)
// Modes: restricted (inventory only) | advisory (knowledge + parameters)
// Keys stored in Script Properties: GEMINI_KEY, MISTRAL_KEY, GROQ_KEY
// Stateless: no chat history stored server-side. Frontend sends last 6 msgs.
// ==========================================

function getSmartMatchEquipment(payload) {
  // ── Global kill switch: Config sheet setting "smartMatchAI: off" disables all AI calls ──
  if (_isAIDisabled()) {
    return { error: true, reply: 'AI assistant is currently disabled by the administrator.', matches: [] };
  }

  const rawPrompt  = String((payload && payload.userPrompt) || '').trim();
  const userPrompt = _sanitizeForAI(rawPrompt);
  const history    = Array.isArray(payload && payload.conversationHistory)
                     ? payload.conversationHistory.slice(-6).map(function(m) {
                         return { role: m.role, text: _sanitizeForAI(String(m.text || '')) };
                       })
                     : [];
  const mode       = String((payload && payload.mode) || 'restricted').trim();
  const isAdvisory = mode === 'advisory';

  if (!userPrompt) return { error: true, reply: 'Empty message.' };

  const context = _buildEquipmentContext(isAdvisory);
  if (!context) return { error: true, reply: 'Equipment database is empty.' };

  const systemInstruction = isAdvisory ? _ADVISORY_SYS(context) : _RESTRICTED_SYS(context);

  // ── Failover chain ──
  const providers = [_callGemini, _callMistral, _callGroq];
  const names     = ['Gemini', 'Mistral', 'Groq'];

  for (let i = 0; i < providers.length; i++) {
    try {
      const raw = providers[i](systemInstruction, userPrompt, history);
      if (raw) {
        const parsed = _parseMatchResponse(raw);
        parsed.provider      = names[i];
        parsed.fallbackUsed  = i > 0;
        parsed.mode          = mode;
        console.log('[SmartMatch] ' + names[i] + ' succeeded (' + mode + ').');
        return parsed;
      }
    } catch (e) {
      console.log('[SmartMatch] ' + names[i] + ' failed: ' + e.toString());
    }
  }

  return {
    error: true,
    reply: 'All smart search channels are currently at capacity. Please try again shortly.',
    matches: [],
    mode: mode
  };
}

// ── System instructions ──

// ── Kill switch check — reads "smartMatchAI" from Config sheet E2:F.
//    Set value to "off"/"disabled"/"false" to disable. Anything else (or missing) = enabled.
//    Cached 5 minutes so it doesn't slow down every chat request.
function _isAIDisabled() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('smart_match_ai_disabled');
    if (cached !== null) return cached === 'true';

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
    var data = sheet.getRange('E2:F').getValues();
    var disabled = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === 'smartmatchai') {
        var v = String(data[i][1] || '').trim().toLowerCase();
        disabled = (v === 'off' || v === 'disabled' || v === 'false' || v === '0');
        break;
      }
    }
    try { cache.put('smart_match_ai_disabled', String(disabled), 300); } catch(e) {}
    return disabled;
  } catch(e) {
    return false; // fail-open: if we can't read config, AI stays enabled
  }
}

// ── Strip PII before sending to external AI: emails, phone numbers, matric IDs ──
function _sanitizeForAI(text) {
  if (!text) return '';
  var s = String(text);
  s = s.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '[email]');           // emails
  s = s.replace(/(\+?6?0?1)[0-9][0-9]{7,8}\b/g, '[phone]');                               // Malaysian mobile
  s = s.replace(/\+?\d[\d\s().-]{8,}\d/g, '[phone]');                                     // other long numbers
  s = s.replace(/\b[A-Za-z]?\d{6,}[A-Za-z]?\b/g, '[id]');                                 // matric-style IDs
  return s;
}

function _RESTRICTED_SYS(context) {
  return 'CRITICAL — RESPOND ONLY WITH VALID JSON. No markdown, no tables, no code fences, no text outside the JSON object.\n' +
    'You are the i-Nstrumen Lab Assistant for IMEN (Institute of Microengineering and Nanoelectronics, UKM). ' +
    'Match the user\'s requested process or fabrication need ONLY to equipment listed in the EQUIPMENT INVENTORY below.\n\n' +
    'RULES:\n' +
    '1. Never invent or suggest equipment not in the list.\n' +
    '2. If nothing matches, say so politely and suggest the closest alternative from the list if any.\n' +
    '3. For multi-turn conversations, remember context from earlier messages.\n' +
    '4. Be concise and helpful.\n' +
    '5. OUTPUT FORMAT — You MUST respond with exactly this JSON shape, nothing else:\n' +
    '{"reply":"your conversational text here","matches":[{"name":"exact equipment name from list","lab":"lab name","reason":"why it matches"}]}\n' +
    'If no matches, use an empty matches array: []\n\n' +
    'EQUIPMENT INVENTORY:\n' + context;
}

function _ADVISORY_SYS(context) {
  return 'CRITICAL — RESPOND ONLY WITH VALID JSON. No markdown, no tables, no code fences, no text outside the JSON object.\n' +
    'You are the i-Nstrumen Lab Assistant with deep expertise in semiconductor fabrication, ' +
    'thin film deposition, material characterization, and nano/micro-engineering equipment. ' +
    'You help researchers at IMEN UKM find the right lab tools and understand how to use them.\n\n' +
    'YOUR CAPABILITIES:\n' +
    '1. MATCH: First check the EQUIPMENT INVENTORY below and recommend the best tools.\n' +
    '2. EXPLAIN: Describe how each matched equipment works, its typical applications, and why it fits their need.\n' +
    '3. PARAMETERS: If asked, suggest general operating parameters (voltages, temperatures, pressures, wavelengths, deposition rates) — these are REFERENCE GUIDELINES only, not exact SOPs.\n' +
    '4. WORKFLOWS: For multi-step processes (e.g., lithography → etching → deposition → characterization), outline the full equipment chain.\n' +
    '5. HONESTY: If a capability does not exist in our labs, state it clearly but mention if such tools are common in similar research facilities.\n' +
    '6. DISCLAIMER: End your reply text with: "Advisory: These suggestions and parameters are for reference only. Always verify procedures, settings, and safety protocols with your lab PIC or equipment manual before operating any instrument."\n\n' +
    'OUTPUT FORMAT — You MUST respond with exactly this JSON shape, nothing else. Put ALL explanation inside the "reply" field as plain text (use \\n for line breaks):\n' +
    '{"reply":"your full response including explanation, parameters if relevant, and disclaimer","matches":[{"name":"exact equipment name from list","lab":"lab name","reason":"why it matches the request"}]}\n' +
    'If no matches, use an empty matches array: []\n\n' +
    'EQUIPMENT INVENTORY:\n' + context;
}

function _buildEquipmentContext(isAdvisory) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.EQUIPMENT);
  if (!sheet) return '';
  const eqList = getDataAsObjects(sheet);
  if (!eqList.length) return '';

  const lines = [];
  eqList.forEach(function(eq) {
    if (!eq.name) return;
    const status = String(eq.status || 'Active');
    if (status === 'Maintenance') return;

    let caps = '';
    try {
      const parsed = typeof eq.processCapabilities === 'string'
        ? JSON.parse(eq.processCapabilities || '[]')
        : (eq.processCapabilities || []);
      if (Array.isArray(parsed) && parsed.length) {
        caps = parsed.map(function(c) {
          let s = c.name || '';
          if (Array.isArray(c.options) && c.options.length) {
            s += ' (' + c.options.map(function(o) {
              return typeof o === 'string' ? o : (o.name || '');
            }).join(', ') + ')';
          }
          return s;
        }).join('; ');
      }
    } catch(e) { caps = ''; }

    const desc = String(eq.description || '').substring(0, isAdvisory ? 250 : 120);
    let line = '- ' + eq.name + ' | Lab: ' + (eq.lab || '?') +
      ' | Status: ' + status +
      (caps ? ' | Capabilities: ' + caps : '') +
      (desc ? ' | Notes: ' + desc : '');

    if (isAdvisory) {
      line += ' | Access: ' + (eq.accessMode || 'both') + ' | Tracking: ' + (eq.trackingUnit || 'Hours');
    }

    lines.push(line);
  });

  return lines.join('\n');
}

function _callGemini(systemInstruction, userPrompt, history) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) throw new Error('GEMINI_KEY not configured');

  const contents = [];
  // Convert history to Gemini format
  history.forEach(function(msg) {
    contents.push({
      role: msg.role === 'user' ? 'user' : 'model',
      parts: [{ text: msg.text }]
    });
  });
  contents.push({
    role: 'user',
    parts: [{ text: userPrompt }]
  });

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + key;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: systemInstruction }] },
      contents: contents,
      generationConfig: { temperature: 0.3, maxOutputTokens: 2048, responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error('Gemini HTTP ' + res.getResponseCode());
  const data = JSON.parse(res.getContentText());
  return data.candidates[0].content.parts[0].text;
}

function _callMistral(systemInstruction, userPrompt, history) {
  const key = PropertiesService.getScriptProperties().getProperty('MISTRAL_KEY');
  if (!key) throw new Error('MISTRAL_KEY not configured');

  const messages = [{ role: 'system', content: systemInstruction }];
  history.forEach(function(msg) {
    messages.push({ role: msg.role === 'user' ? 'user' : 'assistant', content: msg.text });
  });
  messages.push({ role: 'user', content: userPrompt });

  const res = UrlFetchApp.fetch('https://api.mistral.ai/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({
      model: 'mistral-small-latest',
      messages: messages,
      temperature: 0.3,
      max_tokens: 2048,
      response_format: { type: 'json_object' }
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error('Mistral HTTP ' + res.getResponseCode());
  const data = JSON.parse(res.getContentText());
  return data.choices[0].message.content;
}

function _callGroq(systemInstruction, userPrompt, history) {
  const key = PropertiesService.getScriptProperties().getProperty('GROQ_KEY');
  if (!key) throw new Error('GROQ_KEY not configured');

  const messages = [{ role: 'system', content: systemInstruction }];
  history.forEach(function(msg) {
    messages.push({ role: msg.role === 'user' ? 'user' : 'assistant', content: msg.text });
  });
  messages.push({ role: 'user', content: userPrompt });

  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: messages,
      temperature: 0.3,
      max_tokens: 2048,
      response_format: { type: 'json_object' }
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error('Groq HTTP ' + res.getResponseCode());
  const data = JSON.parse(res.getContentText());
  return data.choices[0].message.content;
}

// ── DIAGNOSTIC: run from Apps Script editor → select "debugSmartMatch" → Run → View → Logs ──
// Tests each provider individually and shows exactly which keys are set and what errors occur.
function debugSmartMatch() {
  const props = PropertiesService.getScriptProperties();
  const geminiKey  = props.getProperty('GEMINI_KEY');
  const mistralKey = props.getProperty('MISTRAL_KEY');
  const groqKey    = props.getProperty('GROQ_KEY');

  Logger.log('=== Script Properties Status ===');
  Logger.log('GEMINI_KEY:  ' + (geminiKey  ? '✅ SET (' + geminiKey.substring(0, 8) + '…)' : '❌ MISSING'));
  Logger.log('MISTRAL_KEY: ' + (mistralKey ? '✅ SET (' + mistralKey.substring(0, 8) + '…)' : '❌ MISSING'));
  Logger.log('GROQ_KEY:    ' + (groqKey    ? '✅ SET (' + groqKey.substring(0, 8) + '…)' : '❌ MISSING'));

  // Test equipment context
  Logger.log('');
  Logger.log('=== Equipment Context ===');
  try {
    const ctx = _buildEquipmentContext(false);
    const lines = ctx.split('\n').filter(function(l) { return l.trim(); });
    Logger.log('Equipment items: ' + lines.length);
    if (lines.length > 0) Logger.log('First: ' + lines[0].substring(0, 100));
  } catch(e) {
    Logger.log('❌ Context error: ' + e.toString());
  }

  // Test each provider with a simple prompt
  const testSys = 'You are a test assistant. Reply with exactly: {"reply":"test ok","matches":[]}';
  const testPrompt = 'Say test ok';
  const testHist = [];

  Logger.log('');
  Logger.log('=== Provider Tests ===');

  // Gemini
  if (geminiKey) {
    try {
      Logger.log('Testing Gemini...');
      const r = _callGemini(testSys, testPrompt, testHist);
      Logger.log('✅ Gemini OK: ' + r.substring(0, 80));
    } catch(e) {
      Logger.log('❌ Gemini FAILED: ' + e.toString());
    }
  } else {
    Logger.log('⏭️  Gemini skipped (no key)');
  }

  // Mistral
  if (mistralKey) {
    try {
      Logger.log('Testing Mistral...');
      const r = _callMistral(testSys, testPrompt, testHist);
      Logger.log('✅ Mistral OK: ' + r.substring(0, 80));
    } catch(e) {
      Logger.log('❌ Mistral FAILED: ' + e.toString());
    }
  } else {
    Logger.log('⏭️  Mistral skipped (no key)');
  }

  // Groq
  if (groqKey) {
    try {
      Logger.log('Testing Groq...');
      const r = _callGroq(testSys, testPrompt, testHist);
      Logger.log('✅ Groq OK: ' + r.substring(0, 80));
    } catch(e) {
      Logger.log('❌ Groq FAILED: ' + e.toString());
    }
  } else {
    Logger.log('⏭️  Groq skipped (no key)');
  }

  Logger.log('');
  Logger.log('=== Full getSmartMatchEquipment test (restricted) ===');
  try {
    const result = getSmartMatchEquipment({ userPrompt: 'I need SEM imaging', conversationHistory: [], mode: 'restricted' });
    Logger.log('Restricted: ' + JSON.stringify(result).substring(0, 300));
  } catch(e) {
    Logger.log('❌ Restricted FAILED: ' + e.toString());
  }

  Logger.log('');
  Logger.log('=== Full getSmartMatchEquipment test (advisory) ===');
  try {
    const result = getSmartMatchEquipment({ userPrompt: 'Explain how thermal evaporation works and what settings I should use', conversationHistory: [], mode: 'advisory' });
    Logger.log('Advisory: ' + JSON.stringify(result).substring(0, 300));
  } catch(e) {
    Logger.log('❌ Advisory FAILED: ' + e.toString());
  }
}

function _parseMatchResponse(raw) {
  let text = String(raw || '').trim();
  // Strip markdown code fences if present
  text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');

  // If the response doesn't look like JSON at all, wrap as plain text reply
  if (text.indexOf('{') !== 0) {
    return { error: false, reply: text, matches: [] };
  }

  try {
    const obj = JSON.parse(text);
    if (obj && typeof obj.reply === 'string') {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.EQUIPMENT);
      const eqList = sheet ? getDataAsObjects(sheet) : [];
      const validNames = {};
      eqList.forEach(function(e) { validNames[String(e.name || '').toLowerCase()] = true; });

      const matches = Array.isArray(obj.matches) ? obj.matches.filter(function(m) {
        return m && m.name && validNames[String(m.name).toLowerCase()];
      }) : [];

      return { error: false, reply: obj.reply, matches: matches };
    }
  } catch(e) {
    // JSON parse failed — try extracting reply field via regex
    const replyMatch = text.match(/"reply"\s*:\s*"((?:[^"\\]|\\.)*)"/);
    if (replyMatch) {
      return { error: false, reply: replyMatch[1].replace(/\\n/g, '\n').replace(/\\"/g, '"'), matches: [] };
    }
    // Fallback: return raw text as reply
    return { error: false, reply: text, matches: [] };
  }
  return { error: false, reply: text, matches: [] };
}

function saveEquipment(eqObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.EQUIPMENT);
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for(let i=1; i<data.length; i++) { if(data[i][0] == eqObj.id) { rowIdx = i + 1; break; } }
  const procStr = typeof eqObj.processCapabilities === 'object' ? JSON.stringify(eqObj.processCapabilities) : eqObj.processCapabilities;
  const matStr = typeof eqObj.materialsOptions === 'object' ? JSON.stringify(eqObj.materialsOptions) : eqObj.materialsOptions;
  const rowData = [ eqObj.id, eqObj.assetId, eqObj.name, eqObj.lab, eqObj.description, eqObj.imageUrl, eqObj.status, eqObj.maintenanceReason, eqObj.accessMode, eqObj.trackingUnit, eqObj.calibrationDate, eqObj.picEmail, procStr, matStr ];
  if (rowIdx > -1) { sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]); } else { sheet.appendRow(rowData); }
  return { success: true };
}

function deleteEquipment(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.EQUIPMENT);
  const data = sheet.getDataRange().getValues();
  const targetId = String(id).trim(); 
  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();
    if (sheetId === targetId) {
      sheet.deleteRow(i + 1); 
      return { success: true };
    }
  }
  return { success: false, error: "ID not found" };
}

function addLab(labName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
  const rawLabs = sheet.getRange("A2:A").getValues().flat();
  const labs = rawLabs.filter(l => l && l.toString().trim() !== '');
  if (labs.some(l => l.toString().toLowerCase() === labName.toLowerCase())) return { success: false, error: 'Exists' };
  labs.push(labName); labs.sort();
  const lastRow = sheet.getMaxRows();
  sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  sheet.getRange(2, 1, labs.length, 1).setValues(labs.map(l => [l]));
  return { success: true };
}

function deleteLab(labName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
  const rawLabs = sheet.getRange("A2:A").getValues().flat();
  let labs = rawLabs.filter(l => l && l.toString().trim() !== '');
  const initialLen = labs.length;
  labs = labs.filter(l => l !== labName);
  if (labs.length === initialLen) return { success: false };
  const lastRow = sheet.getMaxRows();
  sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  if (labs.length > 0) sheet.getRange(2, 1, labs.length, 1).setValues(labs.map(l => [l]));
  return { success: true };
}

function renameLab(oldName, newName) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const configSheet = ss.getSheetByName(SHEET_IDS.CONFIG);
   const rawLabs = configSheet.getRange("A2:A").getValues().flat();
   const labs = rawLabs.filter(l => l && l.toString().trim() !== '').map(l => l === oldName ? newName : l);
   const lastRow = configSheet.getMaxRows();
   configSheet.getRange(2, 1, lastRow - 1, 1).clearContent();
   if (labs.length > 0) configSheet.getRange(2, 1, labs.length, 1).setValues(labs.map(l => [l]));
   updateSheetColumn(ss, SHEET_IDS.EQUIPMENT, 'lab', oldName, newName);
   updateSheetColumn(ss, SHEET_IDS.BOOKINGS, 'lab', oldName, newName);
   updateSheetColumn(ss, SHEET_IDS.LOGS, 'lab', oldName, newName);
   return { success: true };
}

function saveConfig(configObj) {
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
   if (configObj.coordinators) sheet.getRange("C2").setValue(JSON.stringify(configObj.coordinators));
   if (configObj.techStaff !== undefined) sheet.getRange("D2").setValue(JSON.stringify(configObj.techStaff));

   if (configObj.officialEmail) {
       const data = sheet.getRange("E2:F").getValues();
       let found = false;
       for(let i=0; i<data.length; i++) {
          if(data[i][0] === 'officialEmail') { sheet.getRange(i+2, 6).setValue(configObj.officialEmail); found = true; break; }
       }
       if (!found) {
           let rowToUpdate = data.length + 2; 
           for(let i=0; i<data.length; i++) { if(!data[i][0]) { rowToUpdate = i + 2; break; } }
           sheet.getRange(rowToUpdate, 5).setValue('officialEmail'); 
           sheet.getRange(rowToUpdate, 6).setValue(configObj.officialEmail); 
       }
   }
   return { success: true };
}

function archiveSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_IDS.LOGS);
  const configSheet = ss.getSheetByName(SHEET_IDS.CONFIG);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  logSheet.copyTo(ss).setName(`Logs_Archive_${timestamp}`);
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).clearContent();
  
  const data = configSheet.getRange("E2:F").getValues();
  let found = false;
  for(let i=0; i<data.length; i++) {
      if(data[i][0] === 'lastArchive') { configSheet.getRange(i+2, 6).setValue(timestamp); found = true; break; }
  }
  if(!found) {
      let rowToUpdate = data.length + 2;
      for(let i=0; i<data.length; i++) { if(!data[i][0]) { rowToUpdate = i + 2; break; } }
      configSheet.getRange(rowToUpdate, 5).setValue('lastArchive');
      configSheet.getRange(rowToUpdate, 6).setValue(timestamp);
  }
  return { success: true };
}

// --- EMAIL HELPERS ---

function getCoordinatorEmail(labName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
    const json = sheet.getRange("C2").getValue();
    const coords = JSON.parse(json);
    const coord = coords.find(c => c.lab === labName);
    return coord ? coord.email : null;
  } catch (e) { return null; }
}

function getTechStaffEmails() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.CONFIG);
    const json = sheet.getRange("D2").getValue();
    const staff = JSON.parse(json);
    return staff.map(function(s) { return s.email; }).filter(Boolean);
  } catch (e) { return []; }
}

// Google Calendar prefill URL for an approved booking ("Add to Google Calendar"
// button in the approval email). The student adds the event to their OWN
// calendar, so no invite management or event cleanup is needed on our side.
// Times are written as floating local times + ctz, so the event lands at the
// correct Malaysian wall-clock time regardless of the student's device timezone.
function _buildCalendarUrl(eqName, lab, dateVal, duration) {
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return null;
  const SLOT_TIMES = {                       // [startHHmmss, endHHmmss, extraDays]
    '8am-1pm': ['080000', '130000', 0],
    '2pm-5pm': ['140000', '170000', 0],
    '1pm-5pm': ['130000', '170000', 0],
    '1 Day':   ['080000', '170000', 0],
    '2 Days':  ['080000', '170000', 1]       // spans start date + next day
  };
  const t = SLOT_TIMES[(duration || '').toString().trim()];
  if (!t) return null;                       // unknown slot -> button is simply omitted
  const tz   = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const day1 = Utilities.formatDate(d, tz, 'yyyyMMdd');
  const dEnd = new Date(d); dEnd.setDate(dEnd.getDate() + t[2]);
  const day2 = Utilities.formatDate(dEnd, tz, 'yyyyMMdd');
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text='    + encodeURIComponent('Equipment booking: ' + eqName + ' — check in on arrival')
    + '&dates='   + day1 + 'T' + t[0] + '/' + day2 + 'T' + t[1]
    + '&ctz='     + encodeURIComponent(tz)
    + '&details=' + encodeURIComponent('Check in at the i-Nstrumen portal BEFORE using the equipment '
        + '(blue Check-in button on the equipment card; verify with the phone number used at booking). '
        + PORTAL_URL + '?lab=' + encodeURIComponent(lab || ''))
    + '&location=' + encodeURIComponent((lab || '') + ', IMEN UKM');
}

function sendEmailSafe(to, subject, htmlBody) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: htmlBody,
      name: 'i-Nstrumen'
    });
    console.log("Email sent to: " + to);
    return true;
  } catch (e) {
    console.log("Email failed to: " + to + ". Reason: " + e.toString());
    return false;
  }
}

// ==========================================
// WEEKLY NO-SHOW REMINDER  (runs every Friday via time-driven trigger)
// ==========================================

// Shared computation used by the Friday trigger AND the manual PIC-summary run.
// Returns { noShows, systemEmail, sentWeek } for the last 7 days.
function _getWeeklyNoShows() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName(SHEET_IDS.BOOKINGS);
  const logSheet  = ss.getSheetByName(SHEET_IDS.LOGS);

  const allBookings = getDataAsObjects(bookSheet);
  const allLogs     = getDataAsObjects(logSheet);

  // ── Read system email for CC ──
  const configSheet = ss.getSheetByName(SHEET_IDS.CONFIG);
  const rawSettings = configSheet.getRange('E2:F').getValues();
  let systemEmail   = 'imenmakmal@gmail.com';
  rawSettings.forEach(function(row) {
    if (row[0] === 'officialEmail' && row[1]) systemEmail = String(row[1]).trim();
  });

  // ── Lookback window: approved bookings whose date has already passed (last 7 days) ──
  const today   = new Date(); today.setHours(0, 0, 0, 0);
  const weekAgo = new Date(today); weekAgo.setDate(weekAgo.getDate() - 7);

  const pastBookings = allBookings.filter(function(b) {
    if ((b.status || '').toString().trim() !== 'Approved') return false;
    // NOTE: bookings without a userEmail are kept — they can't get a student
    // reminder (grouping skips them) but MUST still appear in PIC summaries.
    const bDate = new Date(b.date); bDate.setHours(0, 0, 0, 0);
    return bDate >= weekAgo && bDate < today;
  });

  const sentWeek = today.toLocaleDateString('en-MY', { day: 'numeric', month: 'long', year: 'numeric' });

  if (!pastBookings.length) {
    Logger.log('[i-Nstrumen] Weekly reminder: no past bookings this week.');
    return { noShows: [], systemEmail: systemEmail, sentWeek: sentWeek };
  }

  // ── Normalizers ──
  // CRITICAL: never use toISOString() for day-matching. It converts to UTC,
  // which shifts a Malaysian (UTC+8) local-midnight booking date to the
  // PREVIOUS day — so no booking key ever matched its check-in log key and
  // every student was falsely emailed as a no-show. Format in the sheet's
  // own timezone instead.
  const tz = ss.getSpreadsheetTimeZone();
  function localDay(v) {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return String(v).split('T')[0];
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  // Trim + lowercase so stray spaces / casing differences can't break matching.
  function norm(s) { return (s || '').toString().trim().toLowerCase(); }

  // ── Group every Usage log by equipment + day → list of {name, email} ──
  // A booking counts as attended if ANY check-in on the same machine that day
  // matches EITHER the booked email OR the booked name. Email is the stronger
  // signal (stable across name-spelling differences); name is the fallback for
  // older logs that have no email stored.
  const logsByEquipDay = {};
  allLogs.forEach(function(l) {
    if ((l.action || '') !== 'Usage') return;
    const d = localDay(l.timestamp);
    if (!d) return;
    const k = norm(l.lab) + '|' + norm(l.equipmentName) + '|' + d;
    (logsByEquipDay[k] = logsByEquipDay[k] || []).push({ name: norm(l.userName), email: norm(l.userEmail) });
  });

  function hasCheckIn(b) {
    const entries = logsByEquipDay[norm(b.lab) + '|' + norm(b.equipmentName) + '|' + localDay(b.date)];
    if (!entries) return false;
    const bName = norm(b.userName), bEmail = norm(b.userEmail);
    return entries.some(function(e) {
      return (bEmail && e.email && e.email === bEmail) || (bName && e.name && e.name === bName);
    });
  }

  // ── Identify no-shows ──
  const noShows = pastBookings.filter(function(b) { return !hasCheckIn(b); });
  return { noShows: noShows, systemEmail: systemEmail, sentWeek: sentWeek };
}

// ==========================================
// FRIDAY 5 PM RUN: student reminders + per-lab PIC summaries
// ==========================================
function sendWeeklyNoShowReminder() {
  const wk          = _getWeeklyNoShows();
  const noShows     = wk.noShows;
  const systemEmail = wk.systemEmail;
  const sentWeek    = wk.sentWeek;

  if (!noShows.length) {
    Logger.log('[i-Nstrumen] Weekly reminder: every booking had a check-in. Nothing to send.');
    return;
  }

  // ── Group by userEmail ──
  const byUser = {};
  noShows.forEach(function(b) {
    const em = (b.userEmail || '').toString().trim().toLowerCase();
    if (!em) return;
    if (!byUser[em]) byUser[em] = { name: b.userName || 'Researcher', bookings: [] };
    byUser[em].bookings.push(b);
  });

  // ── Send one summary email per user ──
  Object.keys(byUser).forEach(function(userEmail) {
    const u    = byUser[userEmail];
    const rows = u.bookings.map(function(b) {
      const bDate = new Date(b.date).toLocaleDateString('en-MY', {
        weekday: 'short', day: 'numeric', month: 'short', year: 'numeric'
      });
      return '<tr style="border-bottom:1px solid #f1f5f9;">' +
        '<td style="padding:10px 12px;font-size:13px;">' + (b.equipmentName || '-') + '</td>' +
        '<td style="padding:10px 12px;font-size:13px;color:#4f46e5;">' + (b.lab || '-') + '</td>' +
        '<td style="padding:10px 12px;font-size:13px;">' + bDate + '</td>' +
        '<td style="padding:10px 12px;font-size:13px;color:#64748b;">' + (b.duration || '-') + '</td>' +
        '</tr>';
    }).join('');

    const body =
      '<div style="font-family:Arial,sans-serif;color:#1e293b;max-width:620px;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">' +

        '<div style="background:linear-gradient(135deg,#4f46e5,#7c3aed);color:white;padding:24px 28px;">' +
          '<h2 style="margin:0;font-size:20px;">Weekly Check-in Summary</h2>' +
          '<p style="margin:6px 0 0;opacity:0.85;font-size:13px;">Week ending ' + sentWeek + '</p>' +
        '</div>' +

        '<div style="padding:24px 28px;">' +
          '<p style="margin:0 0 14px;">Hi <strong>' + u.name + '</strong>,</p>' +
          '<p style="margin:0 0 18px;color:#475569;font-size:14px;">This is a friendly reminder from the i-Nstrumen Lab Management System. The following approved booking(s) this week did <strong style="color:#dc2626;">not have a check-in recorded</strong>:</p>' +

          '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">' +
            '<thead>' +
              '<tr style="background:#f8fafc;">' +
                '<th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;">Equipment</th>' +
                '<th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;">Lab</th>' +
                '<th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;">Booked Date</th>' +
                '<th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;">Duration</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + rows + '</tbody>' +
          '</table>' +

          '<div style="margin:20px 0 0;padding:14px 16px;background:#f0fdf4;border-left:4px solid #22c55e;border-radius:4px;font-size:13px;color:#166534;">' +
            '<strong>&#128276; Why does this matter?</strong><br>' +
            'Proper check-ins keep usage records accurate, ensure fair equipment access, and help the lab run smoothly for every researcher.' +
          '</div>' +

          '<p style="margin:18px 0 0;font-size:13px;color:#64748b;">' +
            'If you <em>did</em> use the equipment but forgot to check in, please contact your lab coordinator to update the record. ' +
            'If you were unable to attend, no action is needed — this is a soft reminder only.' +
          '</p>' +
        '</div>' +

        '<div style="background:#f8fafc;padding:14px 28px;text-align:center;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;">' +
          'Automated weekly summary &mdash; i-Nstrumen &middot; IMEN Lab Management System' +
        '</div>' +
      '</div>';

    try {
      MailApp.sendEmail({
        to:       userEmail,
        cc:       systemEmail,
        subject:  '[i-Nstrumen] Friendly Reminder: Missed Check-in(s) — Week of ' + sentWeek,
        htmlBody: body,
        name:     'i-Nstrumen'
      });
      Logger.log('[i-Nstrumen] Reminder sent → ' + userEmail + ' (' + u.bookings.length + ' booking(s))');
    } catch(e) {
      Logger.log('[i-Nstrumen] Failed to send reminder to ' + userEmail + ': ' + e.toString());
    }
  });

  Logger.log('[i-Nstrumen] Weekly run complete — ' + noShows.length + ' no-show(s) across ' + Object.keys(byUser).length + ' user(s).');

  // ── PIC summaries: one email per lab, same Friday schedule ──
  _sendPicSummaries(noShows, systemEmail, sentWeek);
}

// ==========================================
// PIC LAB SUMMARIES — one email per lab listing that lab's missed check-ins,
// sent to the lab's coordinator(s) so they can remind their own users.
// ==========================================
function _sendPicSummaries(noShows, systemEmail, sentWeek) {
  if (!noShows.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // All lab coordinators (PICs) from Config
  let coords = [];
  try { coords = JSON.parse(ss.getSheetByName(SHEET_IDS.CONFIG).getRange('C2').getValue()) || []; } catch (e) { coords = []; }

  // Group no-shows by lab
  const byLab = {};
  noShows.forEach(function(b) {
    const lab = (b.lab || 'Unknown Lab').toString().trim();
    (byLab[lab] = byLab[lab] || []).push(b);
  });

  Object.keys(byLab).forEach(function(lab) {
    const list = byLab[lab];

    // Every coordinator assigned to this lab (a lab can have more than one PIC)
    const picEmails = coords
      .filter(function(c) { return c.lab && c.lab.toString().trim() === lab && c.email; })
      .map(function(c) { return c.email.toString().trim(); });

    // No PIC configured for this lab → deliver to the system mailbox instead
    const to = picEmails.length ? picEmails.join(',') : systemEmail;

    const rows = list.map(function(b, i) {
      const bDate = new Date(b.date).toLocaleDateString('en-MY', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
      return '<tr style="border-bottom:1px solid #f1f5f9;">' +
        '<td style="padding:9px 12px;font-size:12px;color:#94a3b8;">' + (i + 1) + '</td>' +
        '<td style="padding:9px 12px;font-size:13px;"><strong>' + (b.userName || '-') + '</strong></td>' +
        '<td style="padding:9px 12px;font-size:12px;color:#4f46e5;">' + (b.userEmail || '-') + '</td>' +
        '<td style="padding:9px 12px;font-size:13px;">' + (b.equipmentName || '-') + '</td>' +
        '<td style="padding:9px 12px;font-size:12px;">' + bDate + '</td>' +
        '<td style="padding:9px 12px;font-size:12px;color:#64748b;">' + (b.duration || '-') + '</td>' +
        '</tr>';
    }).join('');

    const body =
      '<div style="font-family:Arial,sans-serif;color:#1e293b;max-width:680px;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">' +

        '<div style="background:linear-gradient(135deg,#b45309,#d97706);color:white;padding:24px 28px;">' +
          '<h2 style="margin:0;font-size:20px;">PIC Weekly Summary — Missed Check-ins</h2>' +
          '<p style="margin:6px 0 0;opacity:0.9;font-size:13px;">' + lab + ' &middot; Week ending ' + sentWeek + '</p>' +
        '</div>' +

        '<div style="padding:24px 28px;">' +
          '<p style="margin:0 0 6px;">Dear PIC of <strong>' + lab + '</strong>,</p>' +
          '<p style="margin:0 0 18px;color:#475569;font-size:14px;">' +
            'The following <strong style="color:#b45309;">' + list.length + ' approved booking(s)</strong> in your lab this week had ' +
            '<strong style="color:#dc2626;">no check-in recorded</strong>. ' +
            'Each user has also received an individual soft reminder.' +
          '</p>' +

          '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">' +
            '<thead>' +
              '<tr style="background:#fffbeb;">' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">#</th>' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">User</th>' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">Email</th>' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">Equipment</th>' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">Booked Date</th>' +
                '<th style="padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:.05em;">Slot</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + rows + '</tbody>' +
          '</table>' +

          '<div style="margin:20px 0 0;padding:14px 16px;background:#eff6ff;border-left:4px solid #3b82f6;border-radius:4px;font-size:13px;color:#1e40af;">' +
            '<strong>&#128161; Suggested action:</strong> a quick word with these users helps keep booking slots fair and usage records accurate. ' +
            'If any of them actually used the equipment but forgot to check in, you can update the record from the Control panel.' +
          '</div>' +
        '</div>' +

        '<div style="background:#f8fafc;padding:14px 28px;text-align:center;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;">' +
          'Automated weekly PIC summary &mdash; i-Nstrumen &middot; IMEN Lab Management System' +
        '</div>' +
      '</div>';

    // CC the system mailbox — unless it is already the recipient (fallback case)
    const mailOpts = {
      to:       to,
      subject:  '[i-Nstrumen] PIC Summary: ' + list.length + ' missed check-in(s) — ' + lab + ' — Week of ' + sentWeek,
      htmlBody: body,
      name:     'i-Nstrumen'
    };
    if (to.toLowerCase().indexOf(systemEmail.toLowerCase()) === -1) mailOpts.cc = systemEmail;

    try {
      MailApp.sendEmail(mailOpts);
      Logger.log('[i-Nstrumen] PIC summary sent → ' + to + '  (' + lab + ', ' + list.length + ' no-show(s))');
    } catch (e) {
      Logger.log('[i-Nstrumen] PIC summary FAILED for ' + lab + ': ' + e.toString());
    }
  });
}

// ── ONE-TIME MANUAL RUN — for the first rollout (e.g. on a Saturday). ──
// Sends ONLY the per-lab PIC summaries for the past 7 days, CC to the system
// mailbox. Students are NOT emailed again. Run it from the Apps Script editor:
// select "sendPicSummaryToday" and press Run.
function sendPicSummaryToday() {
  const wk = _getWeeklyNoShows();
  if (!wk.noShows.length) {
    Logger.log('[i-Nstrumen] No missed check-ins in the last 7 days — no PIC summary needed.');
    return;
  }
  _sendPicSummaries(wk.noShows, wk.systemEmail, wk.sentWeek);
  Logger.log('[i-Nstrumen] One-time PIC summary run complete — ' + wk.noShows.length + ' no-show(s) sent to PIC(s).');
}

// ── Run this ONCE from Apps Script editor to register the Friday trigger ──
function installWeeklyTrigger() {
  // Remove any existing trigger for this function (prevents duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyNoShowReminder') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Removed old trigger.');
    }
  });
  // Create new time-driven trigger: every Friday at 5–6 PM (script owner timezone)
  ScriptApp.newTrigger('sendWeeklyNoShowReminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();
  Logger.log('[i-Nstrumen] ✅ Weekly Friday trigger installed — fires every Friday 5–6 PM.');
}

// ── DIAGNOSTIC: run from the Apps Script editor, then View → Logs. ──
// Shows every approved booking from the last 7 days and whether a matching
// check-in log was found — WITHOUT sending any email. Use this to verify the
// no-show matching against real database rows before Friday's automatic run.
function debugNoShowCheck() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const tz        = ss.getSpreadsheetTimeZone();
  const bookings  = getDataAsObjects(ss.getSheetByName(SHEET_IDS.BOOKINGS));
  const logs      = getDataAsObjects(ss.getSheetByName(SHEET_IDS.LOGS));

  function localDay(v) {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return String(v).split('T')[0];
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  function norm(s) { return (s || '').toString().trim().toLowerCase(); }

  const today   = new Date(); today.setHours(0, 0, 0, 0);
  const weekAgo = new Date(today); weekAgo.setDate(weekAgo.getDate() - 7);

  // Group Usage logs by equipment+day → [{name,email}]
  const logsByEquipDay = {};
  logs.forEach(function(l) {
    if ((l.action || '') !== 'Usage') return;
    const d = localDay(l.timestamp);
    if (!d) return;
    const k = norm(l.lab) + '|' + norm(l.equipmentName) + '|' + d;
    (logsByEquipDay[k] = logsByEquipDay[k] || []).push({ name: norm(l.userName), email: norm(l.userEmail) });
  });

  Logger.log('Spreadsheet timezone: ' + tz);
  let flagged = 0, matched = 0;
  bookings.forEach(function(b) {
    if ((b.status || '').toString().trim() !== 'Approved') return;
    const bDate = new Date(b.date); bDate.setHours(0, 0, 0, 0);
    if (bDate < weekAgo || bDate >= today) return;

    const eqDay   = norm(b.lab) + '|' + norm(b.equipmentName) + '|' + localDay(b.date);
    const entries = logsByEquipDay[eqDay] || [];
    const bName = norm(b.userName), bEmail = norm(b.userEmail);
    const hit = entries.some(function(e) {
      return (bEmail && e.email && e.email === bEmail) || (bName && e.name && e.name === bName);
    });
    if (hit) matched++; else flagged++;

    Logger.log((hit ? '✅ CHECKED-IN ' : '⚠️ NO-SHOW    ')
      + b.equipmentName + ' | ' + b.userName + ' <' + (b.userEmail || 'no-email') + '> | ' + localDay(b.date));
    // For a flagged no-show, show who DID check in on that machine that day (the evidence).
    if (!hit) {
      if (entries.length === 0) {
        Logger.log('        └─ no check-in logged on this machine that day → genuine no-show.');
      } else {
        entries.forEach(function(e) {
          Logger.log('        └─ someone else checked in: ' + (e.name || '?') + ' <' + (e.email || 'no-email') + '>');
        });
      }
    }
  });
  Logger.log('Summary: ' + matched + ' checked-in, ' + flagged + ' would receive the reminder.');
}

// ==========================================
// PRE-SLOT CHECK-IN REMINDER  (daily triggers at 7 AM and 1 PM)
// Emails students with an Approved booking TODAY ~1 hour before their slot
// starts, walking them through the check-in procedure (the step students
// always miss). Dedupe/audit: Bookings col V "checkInReminderSent" gets a
// timestamp after a successful send, so a booking is never reminded twice.
// "2 Days" bookings are reminded on the start date only.
// ==========================================

const CHECKIN_SLOTS_MORNING   = ['8am-1pm', '1 Day', '2 Days']; // slots starting 8 AM
const CHECKIN_SLOTS_AFTERNOON = ['2pm-5pm', '1pm-5pm'];         // 2 PM + legacy slot

function _slotLabel(duration) {
  switch ((duration || '').toString().trim()) {
    case '8am-1pm': return '8:00 AM – 1:00 PM';
    case '2pm-5pm': return '2:00 PM – 5:00 PM';
    case '1pm-5pm': return '1:00 PM – 5:00 PM';
    case '1 Day':   return '8:00 AM – 5:00 PM (full day)';
    case '2 Days':  return '8:00 AM start (2-day booking)';
    default:        return duration || '-';
  }
}

// Shared READ-ONLY computation used by the trigger handler and the debug
// dry-run. Returns { candidates: [{booking, row}], todayStr, tz }.
function _getPendingCheckInReminders(slotSet) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName(SHEET_IDS.BOOKINGS);
  const logSheet  = ss.getSheetByName(SHEET_IDS.LOGS);
  const tz        = ss.getSpreadsheetTimeZone();

  // Same normalizers as _getWeeklyNoShows — NEVER toISOString() for day
  // matching (it shifts UTC+8 local dates to the previous day).
  function localDay(v) {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return String(v).split('T')[0];
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  function norm(s) { return (s || '').toString().trim().toLowerCase(); }

  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const bookings = getDataAsObjects(bookSheet);
  const logs     = getDataAsObjects(logSheet);

  // id -> sheet row (col A read once). Bookings rows are append-only (never
  // deleted programmatically), so the mapping stays valid during this run.
  const rowById = {};
  const lastRow = bookSheet.getLastRow();
  if (lastRow > 1) {
    bookSheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r, i) {
      const id = String(r[0]).trim();
      if (id && !(id in rowById)) rowById[id] = i + 2;
    });
  }

  // Group Usage logs by lab|equipment|localDay → [{name,email}] (identical
  // structure to the weekly no-show job) to skip students already checked in.
  const logsByEquipDay = {};
  logs.forEach(function(l) {
    if ((l.action || '') !== 'Usage') return;
    const d = localDay(l.timestamp);
    if (!d) return;
    const k = norm(l.lab) + '|' + norm(l.equipmentName) + '|' + d;
    (logsByEquipDay[k] = logsByEquipDay[k] || []).push({ name: norm(l.userName), email: norm(l.userEmail) });
  });
  function hasCheckIn(b) {
    const entries = logsByEquipDay[norm(b.lab) + '|' + norm(b.equipmentName) + '|' + localDay(b.date)];
    if (!entries) return false;
    const bName = norm(b.userName), bEmail = norm(b.userEmail);
    return entries.some(function(e) {
      return (bEmail && e.email && e.email === bEmail) || (bName && e.name && e.name === bName);
    });
  }

  const candidates = [];
  bookings.forEach(function(b) {
    if ((b.status || '').toString().trim() !== 'Approved') return;
    if (localDay(b.date) !== todayStr) return;                            // today, LOCAL day
    if (slotSet.indexOf((b.duration || '').toString().trim()) === -1) return;
    if (b.checkInReminderSent) return;   // already reminded (col V)
    if (hasCheckIn(b)) return;           // already checked in (early check-in edge)
    if (!(b.userEmail || '').toString().trim()) {
      Logger.log('[i-Nstrumen] Check-in reminder skipped (no email): ' + b.userName + ' / ' + b.equipmentName);
      return;
    }
    candidates.push({ booking: b, row: rowById[String(b.id).trim()] || -1 });
  });

  return { candidates: candidates, todayStr: todayStr, tz: tz };
}

// TRIGGER HANDLER — runs daily at 7 AM (morning slots) and 1 PM (afternoon
// slots). Installed once via installCheckInReminderTriggers().
function sendCheckInReminders() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const bookSheet = ss.getSheetByName(SHEET_IDS.BOOKINGS);
  const tz        = ss.getSpreadsheetTimeZone();

  // Auto-heal: make sure the col V header exists (same pattern as saveLog).
  if (bookSheet.getRange(1, 22).getValue() === '') {
    bookSheet.getRange(1, 22).setValue('checkInReminderSent');
  }

  // Branch on the hour IN THE SPREADSHEET TIMEZONE so it can never disagree
  // with the localDay() date matching used to pick candidates.
  const hourNow   = Number(Utilities.formatDate(new Date(), tz, 'H'));
  const isMorning = hourNow < 12;
  const slotSet   = isMorning ? CHECKIN_SLOTS_MORNING : CHECKIN_SLOTS_AFTERNOON;
  const runLabel  = isMorning ? 'morning' : 'afternoon';

  const pending = _getPendingCheckInReminders(slotSet);
  if (!pending.candidates.length) {
    Logger.log('[i-Nstrumen] Check-in reminders (' + runLabel + '): nothing to send.');
    return;
  }

  let sent = 0;
  pending.candidates.forEach(function(c) {
    const b          = c.booking;
    const slotLabel  = _slotLabel(b.duration);
    const portalLink = PORTAL_URL + '?lab=' + encodeURIComponent(b.lab || '');
    const dObj       = new Date(b.date);
    const prettyDate = isNaN(dObj.getTime()) ? String(b.date) : Utilities.formatDate(dObj, tz, 'EEEE, d MMMM yyyy');

    const emailBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0284c7; color: white; padding: 20px;">
          <h2 style="margin: 0;">Your Session Starts Soon</h2>
          <p style="margin: 5px 0 0 0; opacity: 0.9;">Check-in required before use</p>
        </div>
        <div style="padding: 20px;">
          <p>Dear ${b.userName},</p>
          <p style="color:#555;">A reminder for your <strong>approved</strong> equipment booking <strong>today</strong>:</p>
          <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666; width: 30%;"><strong>Equipment</strong></td><td style="padding: 10px; font-size: 15px; font-weight: bold;">${b.equipmentName}</td></tr>
            <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Lab</strong></td><td style="padding: 10px;">${b.lab}</td></tr>
            <tr style="border-bottom: 1px solid #eee;"><td style="padding: 10px; color: #666;"><strong>Date</strong></td><td style="padding: 10px;">${prettyDate}</td></tr>
            <tr><td style="padding: 10px; color: #666;"><strong>Slot</strong></td><td style="padding: 10px;">${slotLabel}</td></tr>
          </table>

          <div style="margin-top: 20px; padding: 15px; background-color: #eff6ff; border-left: 4px solid #0284c7; border-radius: 4px;">
            <strong style="color:#0c4a6e;">How to check in (takes under a minute):</strong>
            <ol style="margin: 10px 0 0; padding-left: 20px; color: #1e3a5f; font-size: 13px; line-height: 1.7;">
              <li>Open the <strong>i-Nstrumen portal</strong> (button below) and select your lab (${b.lab}).</li>
              <li>Find the <strong>${b.equipmentName}</strong> card and press the blue <strong>Check-in</strong> button.</li>
              <li>Verify your identity with the <strong>same phone number you used when booking</strong>.</li>
              <li>Confirm the session details and press <strong>Confirm Check-in</strong>.</li>
              <li>You will receive a &ldquo;Session Started&rdquo; email &mdash; that is your proof of check-in.</li>
            </ol>
          </div>

          <div style="margin-top: 15px; padding: 15px; background-color: #fef9c3; border-left: 4px solid #ca8a04; color: #713f12; font-size: 13px;">
            <strong>&#9888; Important:</strong> Check-in must be completed <strong>before you start using the equipment</strong>.
            Bookings with no check-in are flagged in the weekly summary sent to your lab PIC, and attempting to check in
            <strong>after your slot has ended</strong> automatically marks the booking as &ldquo;Not Attend&rdquo;.
          </div>

          <div style="text-align:center; margin: 25px 0 5px;">
            <a href="${portalLink}" style="background-color:#0284c7; color:white; padding:12px 28px; text-decoration:none; border-radius:5px; font-weight:bold; font-family: sans-serif;">Open i-Nstrumen &amp; Check In</a>
          </div>
        </div>
        <div style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #9ca3af;">
          Generated automatically by i-Nstrumen &mdash; IMEN Lab Management System
        </div>
      </div>
    `;

    const ok = sendEmailSafe(b.userEmail, `[i-Nstrumen] Check-in Reminder: ${b.equipmentName} — Today, ${slotLabel}`, emailBody);
    if (ok && c.row > 0) {
      bookSheet.getRange(c.row, 22).setValue(new Date());   // dedupe + audit stamp
      sent++;
    }
  });
  Logger.log('[i-Nstrumen] Check-in reminders (' + runLabel + '): ' + sent + '/' + pending.candidates.length + ' sent and stamped.');
}

// ── Run this ONCE from the Apps Script editor to register both daily triggers ──
function installCheckInReminderTriggers() {
  // Remove any existing triggers for this handler (prevents duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendCheckInReminders') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Removed old check-in reminder trigger.');
    }
  });
  ScriptApp.newTrigger('sendCheckInReminders').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('sendCheckInReminders').timeBased().everyDays(1).atHour(13).create();
  Logger.log('[i-Nstrumen] ✅ Check-in reminder triggers installed — daily 7–8 AM (morning slots) and 1–2 PM (afternoon slots).');
}

// ── DIAGNOSTIC: run from the Apps Script editor, then View → Logs. ──
// Shows who WOULD receive a check-in reminder for BOTH runs — WITHOUT sending
// any email or writing anything. Use this to verify against real rows.
function debugCheckInReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Script TZ: ' + Session.getScriptTimeZone() + ' | Spreadsheet TZ: ' + ss.getSpreadsheetTimeZone());
  [['MORNING 7 AM run', CHECKIN_SLOTS_MORNING], ['AFTERNOON 1 PM run', CHECKIN_SLOTS_AFTERNOON]]
  .forEach(function(run) {
    const p = _getPendingCheckInReminders(run[1]);
    Logger.log('-- ' + run[0] + ' (' + run[1].join(', ') + ') | today=' + p.todayStr
      + ' | ' + p.candidates.length + ' would be emailed:');
    p.candidates.forEach(function(c) {
      const b = c.booking;
      Logger.log('   WOULD EMAIL: ' + b.equipmentName + ' | ' + b.lab + ' | ' + b.userName
        + ' <' + b.userEmail + '> | ' + b.duration + ' | sheet row ' + c.row);
    });
  });
}

// Run this ONCE from the Apps Script editor to grant MailApp permission
function testEmailNow() {
  var recipient = Session.getActiveUser().getEmail();
  sendEmailSafe(
    recipient,
    '[i-Nstrumen] Email Test',
    '<div style="font-family:Arial,sans-serif;padding:20px;"><h2 style="color:#4f46e5;">Email is working!</h2><p>If you received this, <strong>MailApp is authorized</strong> and all booking/walk-in notifications will now deliver correctly.</p><p style="color:#888;font-size:12px;">Sent from i-Nstrumen Apps Script</p></div>'
  );
  Logger.log('Test email dispatched to: ' + recipient);
}

// --- GENERAL HELPERS ---

function updateEquipmentStatus(eqId, status, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IDS.EQUIPMENT);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == eqId) {
      sheet.getRange(i + 1, 7).setValue(status); 
      sheet.getRange(i + 1, 8).setValue(reason || ''); 
      break;
    }
  }
}

function updateSheetColumn(ss, sheetName, colName, oldVal, newVal) {
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  const values = range.getValues();
  const headers = values[0];
  const colIdx = headers.indexOf(colName);
  if (colIdx === -1) return;
  let dirty = false;
  for(let i=1; i<values.length; i++) {
    if(values[i][colIdx] === oldVal) { values[i][colIdx] = newVal; dirty = true; }
  }
  if(dirty) range.setValues(values);
}

// --- ROBUST DATA FETCHING HELPERS ---

function getDataAsObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0];
  const objects = [];
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === "")) continue;
    
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      if (val && typeof val === 'object' && val.toString() === 'Exception') {
         val = ""; 
      }
      
      if (headers[j] === 'userPhone' || headers[j] === 'userId' || headers[j] === 'id') {
          val = String(val);
      }
      
      obj[headers[j]] = val;
    }
    objects.push(obj);
  }
  return objects;
}

function getLastNRowsAsObjects(sheet, n) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startRow = Math.max(2, lastRow - n + 1);
  const numRows = lastRow - startRow + 1;
  
  if (numRows < 1) return [];

  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues().reverse(); 
  const objects = [];
  const tz = sheet.getParent().getSpreadsheetTimeZone();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === "")) continue;

    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = row[j];
      
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      if (headers[j] === 'userPhone' || headers[j] === 'userId' || headers[j] === 'id') {
          val = String(val);
      }
      
      obj[headers[j]] = val;
    }
    objects.push(obj);
  }
  return objects;
}