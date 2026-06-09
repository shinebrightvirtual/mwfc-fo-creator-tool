// ============================================================
// MWFC Fragrance Oil Listing Tool — Google Apps Script
// Deploy as Web App: Execute as Me, Access: Anyone
// ============================================================

const SHEET_NAME = "Fragrance Oils";
const FILL_LINES_FOLDER_ID = "1mzHh_1FW2yKGbkTlf-vMv9gG4UUvECqU";
const SOAP_TESTING_FOLDER_ID = "1xRp_MhxRovGyyEFkDszfa4JX6zp7QAV_";

// Friendly display names — must match sheet header row exactly
const COLUMNS = [
  "Oil Name","Midwest Maker Signature Scent?","Description","Top Notes","Middle Notes","Base Notes","Alt Naming Ideas",
  "Vanillin %","Ethyl Vanillin %","Flashpoint (°F)","Phthalate Free?","Contains EOs?",
  "IFRA Cat 1","IFRA Cat 2","IFRA Cat 3","IFRA Cat 4","IFRA Cat 5A","IFRA Cat 5B","IFRA Cat 5C","IFRA Cat 5D",
  "IFRA Cat 6","IFRA Cat 7A","IFRA Cat 7B","IFRA Cat 8","IFRA Cat 9","IFRA Cat 10A","IFRA Cat 10B",
  "IFRA Cat 11A","IFRA Cat 11B","IFRA Cat 12",
  "SDS Link","IFRA Cert Link","EU Allergen Link",
  "Candle FO Load","Wax Used","Burn Space","Candle Size","Burn Time","Wick(s)","Cold Throw","Hot Throw","Candle Notes","Fill Line Photo",
  "Soap FO Load","Lye Concentration","Water:Lye Ratio","Superfat %","Oil Temp","Vitamin E %","Soap Recipe",
  "Acceleration","Ricing","Discoloration","Trace","Scent Retention","Design Notes","Soap Notes","Soap Test Photo"
];

// Map from tool field keys to friendly sheet column names
const KEY_TO_COL = {
  "name":"Oil Name","original":"Midwest Maker Signature Scent?","desc":"Description",
  "top":"Top Notes","mid":"Middle Notes","base":"Base Notes","altnames":"Alt Naming Ideas",
  "vanillin":"Vanillin %","ethylvanillin":"Ethyl Vanillin %","flash":"Flashpoint (°F)",
  "phthalate":"Phthalate Free?","eo":"Contains EOs?",
  "ifra_1":"IFRA Cat 1","ifra_2":"IFRA Cat 2","ifra_3":"IFRA Cat 3","ifra_4":"IFRA Cat 4",
  "ifra_5A":"IFRA Cat 5A","ifra_5B":"IFRA Cat 5B","ifra_5C":"IFRA Cat 5C","ifra_5D":"IFRA Cat 5D",
  "ifra_6":"IFRA Cat 6","ifra_7A":"IFRA Cat 7A","ifra_7B":"IFRA Cat 7B","ifra_8":"IFRA Cat 8",
  "ifra_9":"IFRA Cat 9","ifra_10A":"IFRA Cat 10A","ifra_10B":"IFRA Cat 10B",
  "ifra_11A":"IFRA Cat 11A","ifra_11B":"IFRA Cat 11B","ifra_12":"IFRA Cat 12",
  "sds":"SDS Link","ifracert":"IFRA Cert Link","eu":"EU Allergen Link",
  "c_load":"Candle FO Load","c_wax":"Wax Used","c_space":"Burn Space","c_size":"Candle Size",
  "c_burntime":"Burn Time","c_wick":"Wick(s)","c_cold":"Cold Throw","c_hot":"Hot Throw",
  "c_notes":"Candle Notes","fill_line_photo":"Fill Line Photo",
  "s_load":"Soap FO Load","s_lye":"Lye Concentration","s_waterratio":"Water:Lye Ratio",
  "s_sfat":"Superfat %","s_oiltemp":"Oil Temp","s_vite":"Vitamin E %","s_recipe":"Soap Recipe",
  "s_accel":"Acceleration","s_ricing":"Ricing","s_discolor":"Discoloration","s_trace":"Trace",
  "s_scent":"Scent Retention","s_design":"Design Notes","s_notes":"Soap Notes",
  "soap_test_photo":"Soap Test Photo"
};

// Reverse map: friendly col name → tool field key
const COL_TO_KEY = {};
Object.keys(KEY_TO_COL).forEach(function(k) { COL_TO_KEY[KEY_TO_COL[k]] = k; });

// Cache key prefix for chunked saves
const CACHE_PREFIX = "mwfc_chunk1_";

function doGet(e) {
  var result;
  try {
    var action = e && e.parameter && e.parameter.action ? e.parameter.action : null;
    if      (action === "getAll")     result = getAllOils();
    else if (action === "getOne")     result = getOneOil(e.parameter.name);
    else if (action === "search")     result = searchOils(e.parameter.q);
    else if (action === "saveChunk")  result = saveChunk(e.parameter);
    else if (action === "save")       result = saveOilDirect(e.parameter);
    else if (action === "delete")     result = deleteOil(e.parameter.name);
    else if (action === "upload")     result = uploadPhoto(
      e.parameter.oilName, e.parameter.photoType,
      e.parameter.fileName, e.parameter.fileData, e.parameter.mimeType
    );
    else if (action === "debug") result = debugSheet();
    else result = { ok: true, message: "MWFC FO Tool API" };
  } catch(err) {
    result = { error: err.toString() };
  }

  var json = JSON.stringify(result);
  var cb = e && e.parameter && e.parameter.callback ? e.parameter.callback : null;
  if (cb) {
    return ContentService
      .createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { return doGet(e); }

// ── Chunked save ─────────────────────────────────────────────
// Chunk 1: store params in cache, return ok
// Chunk 2: merge with cached chunk 1, write to sheet
function saveChunk(params) {
  var cache = CacheService.getScriptCache();
  var name = params.name;
  if (!name) return { error: "Missing name" };

  if (params.chunk === "1") {
    // Store chunk 1 in cache (expires in 10 mins)
    var data = {};
    Object.keys(params).forEach(function(k) {
      if (k !== "action" && k !== "chunk" && k !== "callback") {
        data[k] = params[k];
      }
    });
    cache.put(CACHE_PREFIX + name, JSON.stringify(data), 600);
    return { ok: true };
  }

  if (params.chunk === "2") {
    // Get chunk 1 from cache
    var cached = cache.get(CACHE_PREFIX + name);
    var merged = cached ? JSON.parse(cached) : {};
    // Merge chunk 2
    Object.keys(params).forEach(function(k) {
      if (k !== "action" && k !== "chunk" && k !== "callback") {
        merged[k] = params[k];
      }
    });
    cache.remove(CACHE_PREFIX + name);
    return writeOilToSheet(merged);
  }

  return { error: "Invalid chunk number" };
}

// ── Direct save (fallback, for small payloads) ───────────────
function saveOilDirect(params) {
  var data = {};
  Object.keys(params).forEach(function(k) {
    if (k !== "action" && k !== "callback") data[k] = params[k];
  });
  return writeOilToSheet(data);
}

// ── Write oil to sheet ───────────────────────────────────────
function writeOilToSheet(data) {
  if (!data.name) return { error: "Missing oil name" };
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var allData = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  var existingRow = -1;
  for (var i = 0; i < allData.length; i++) {
    if (allData[i][nameIdx] && String(allData[i][nameIdx]).toLowerCase() === String(data.name).toLowerCase()) {
      existingRow = DATA_START_ROW + i;
      break;
    }
  }
  var existingObj = {};
  if (existingRow > 0) {
    headers.forEach(function(h, i) { existingObj[h] = String(allData[existingRow - DATA_START_ROW][i] || ""); });
  }
  var row = COLUMNS.map(function(col) {
    var key = COL_TO_KEY[col] || col;
    var val = data[key] !== undefined && data[key] !== null && String(data[key]) !== "" ? data[key] : "";
    if (val !== "") return val;
    if (existingObj[col] !== undefined && existingObj[col] !== "") return existingObj[col];
    return "";
  });
  var targetRow;
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    targetRow = existingRow;
  } else {
    // Find next empty row starting from DATA_START_ROW
    var lastRow = sheet.getLastRow();
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  }
  // Apply checkbox validation to boolean columns
  var boolCols = ["Midwest Maker Signature Scent?", "Phthalate Free?", "Contains EOs?"];
  var sheetHeaders = getSheetHeaders(sheet);
  boolCols.forEach(function(colName) {
    var colIdx = sheetHeaders.indexOf(colName);
    if (colIdx !== -1) {
      var cell = sheet.getRange(targetRow, colIdx + 1);
      var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      cell.setDataValidation(rule);
    }
  });
  return { success: true, action: existingRow > 0 ? "updated" : "created" };
}

// ── Sheet helpers ────────────────────────────────────────────
function debugSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  return {
    sheetName: sheet.getName(),
    totalDataRows: data.length,
    headers: headers.slice(0, 5),
    firstDataRow: data.length > 0 ? data[0].slice(0, 5) : [],
    nameColIndex: headers.indexOf("Oil Name"),
    headerRow: HEADER_ROW,
    dataStartRow: DATA_START_ROW
  };
}

// Row offsets for the styled sheet
var HEADER_ROW = 3;  // Row 3 has the actual column headers
var DATA_START_ROW = 4;  // Data starts at row 4

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(HEADER_ROW, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.getRange(HEADER_ROW, 1, 1, COLUMNS.length).setFontWeight("bold");
    sheet.setFrozenRows(HEADER_ROW);
  }
  return sheet;
}

function getSheetHeaders(sheet) {
  return sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  return sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, sheet.getLastColumn()).getValues();
}

function getAllOils() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  var oils = data
    .filter(function(row) { return row[nameIdx] && String(row[nameIdx]).trim() !== ""; })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        var key = COL_TO_KEY[h] || h;
        obj[key] = row[i] !== undefined ? String(row[i]) : "";
      });
      return obj;
    });
  return { oils: oils, sheetUrl: ss.getUrl() };
}

function getOneOil(name) {
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  for (var i = 0; i < data.length; i++) {
    if (data[i][nameIdx] && String(data[i][nameIdx]).toLowerCase() === String(name).toLowerCase()) {
      var obj = {};
      headers.forEach(function(h, idx) {
        var key = COL_TO_KEY[h] || h;
        obj[key] = data[i][idx] !== undefined ? String(data[i][idx]) : "";
      });
      return { oil: obj };
    }
  }
  return { oil: null };
}

function searchOils(q) {
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  var lower = q.toLowerCase();
  var names = data
    .map(function(row) { return String(row[nameIdx]); })
    .filter(function(n) { return n && n.trim() !== "" && n.toLowerCase().indexOf(lower) !== -1; });
  return { names: names };
}

function deleteOil(name) {
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  for (var i = 0; i < data.length; i++) {
    if (data[i][nameIdx] && String(data[i][nameIdx]).toLowerCase() === String(name).toLowerCase()) {
      sheet.deleteRow(DATA_START_ROW + i);
      return { success: true };
    }
  }
  return { error: "Not found" };
}

function uploadPhoto(oilName, photoType, fileName, fileData, mimeType) {
  var folderId = photoType === "fill_line" ? FILL_LINES_FOLDER_ID : SOAP_TESTING_FOLDER_ID;
  var parentFolder = DriveApp.getFolderById(folderId);
  var subName = oilName.replace(/[^a-zA-Z0-9 +\-]/g, "").trim();
  var subFolder;
  var existing = parentFolder.getFoldersByName(subName);
  subFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(subName);
  var decoded = Utilities.base64Decode(fileData);
  var blob = Utilities.newBlob(decoded, mimeType, fileName);
  var file = subFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = "https://drive.google.com/file/d/" + file.getId() + "/view";
  var sheet = getOrCreateSheet();
  var headers = getSheetHeaders(sheet);
  var data = getSheetData(sheet);
  var nameIdx = headers.indexOf("Oil Name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  var colKey = photoType === "fill_line" ? "Fill Line Photo" : "Soap Test Photo";
  var colIdx = headers.indexOf(colKey);
  for (var i = 0; i < data.length; i++) {
    if (data[i][nameIdx] && String(data[i][nameIdx]).toLowerCase() === String(oilName).toLowerCase()) {
      sheet.getRange(DATA_START_ROW + i, colIdx + 1).setValue(url);
      break;
    }
  }
  return { success: true, url: url };
}
