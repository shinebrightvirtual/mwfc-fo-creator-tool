// ============================================================
// MWFC Fragrance Oil Listing Tool — Google Apps Script
// Serves as data API only. UI lives on Vercel.
// Deploy as Web App: Execute as Me, Access: Anyone
// ============================================================

const SHEET_NAME = "Fragrance Oils";
const FILL_LINES_FOLDER_ID = "1mzHh_1FW2yKGbkTlf-vMv9gG4UUvECqU";
const SOAP_TESTING_FOLDER_ID = "1xRp_MhxRovGyyEFkDszfa4JX6zp7QAV_";

const COLUMNS = [
  "name","original","desc","top","mid","base","altnames",
  "vanillin","ethylvanillin","flash","gravity","color","phthalate","eo",
  "ifra_1","ifra_2","ifra_3","ifra_4","ifra_5A","ifra_5B","ifra_5C","ifra_5D",
  "ifra_6","ifra_7A","ifra_7B","ifra_8A","ifra_8B","ifra_9","ifra_10A","ifra_10B",
  "ifra_11A","ifra_11B",
  "sds","ifracert","eu",
  "c_load","c_wax","c_space","c_size","c_burntime","c_wick","c_cold","c_hot","c_notes",
  "s_load","s_lye","s_waterratio","s_sfat","s_oiltemp","s_vite","s_recipe",
  "s_accel","s_ricing","s_discolor","s_trace","s_scent","s_design","s_notes",
  "fill_line_photo","soap_test_photo"
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var output;
  try {
    var action = e && e.parameter && e.parameter.action ? e.parameter.action : null;
    var result;

    if (action === "getAll")        result = getAllOils();
    else if (action === "getOne")   result = getOneOil(e.parameter.name);
    else if (action === "search")   result = searchOils(e.parameter.q);
    else if (action === "save")     result = saveOil(JSON.parse(e.parameter.data));
    else if (action === "delete")   result = deleteOil(e.parameter.name);
    else if (action === "upload")   result = uploadPhoto(
      e.parameter.oilName, e.parameter.photoType,
      e.parameter.fileName, e.parameter.fileData, e.parameter.mimeType
    );
    else result = { error: "Unknown action" };

    output = ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    output = ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllOils() {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { oils: [] };
  var headers = data[0];
  var oils = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? String(row[i]) : ""; });
    return obj;
  });
  return { oils: oils };
}

function getOneOil(name) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var nameIdx = headers.indexOf("name");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]).toLowerCase() === String(name).toLowerCase()) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx] !== undefined ? String(data[i][idx]) : ""; });
      return { oil: obj };
    }
  }
  return { oil: null };
}

function searchOils(q) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { names: [] };
  var headers = data[0];
  var nameIdx = headers.indexOf("name");
  var lower = q.toLowerCase();
  var names = data.slice(1)
    .map(function(row) { return String(row[nameIdx]); })
    .filter(function(n) { return n.toLowerCase().indexOf(lower) !== -1; });
  return { names: names };
}

function saveOil(data) {
  var sheet = getOrCreateSheet();
  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var nameIdx = headers.indexOf("name");
  var existingRow = -1;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][nameIdx]).toLowerCase() === String(data.name).toLowerCase()) {
      existingRow = i + 1;
      break;
    }
  }
  var existingObj = {};
  if (existingRow > 0) {
    headers.forEach(function(h, i) { existingObj[h] = String(allData[existingRow - 1][i] || ""); });
  }
  var row = COLUMNS.map(function(col) {
    if (data[col] !== undefined && data[col] !== null && data[col] !== "") return data[col];
    if (existingObj[col]) return existingObj[col];
    return "";
  });
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, action: existingRow > 0 ? "updated" : "created" };
}

function deleteOil(name) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var nameIdx = headers.indexOf("name");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]).toLowerCase() === String(name).toLowerCase()) {
      sheet.deleteRow(i + 1);
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
  // Write URL back to sheet
  var sheet = getOrCreateSheet();
  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var nameIdx = headers.indexOf("name");
  var colKey = photoType === "fill_line" ? "fill_line_photo" : "soap_test_photo";
  var colIdx = headers.indexOf(colKey);
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][nameIdx]).toLowerCase() === String(oilName).toLowerCase()) {
      sheet.getRange(i + 1, colIdx + 1).setValue(url);
      break;
    }
  }
  return { success: true, url: url };
}
