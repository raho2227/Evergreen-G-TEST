// ✅ Evergreen Industries - Smart Data Sync Google Apps Script
// Ye code automatically column headers ko match karega aur uske niche sahi data daal dega.

function doGet(e) {
  var action = e.parameter.action;
  var sheetId = e.parameter.sheetId; 
  var ss;
  
  // Security Check: Enforce API Key
  var API_SECRET = 'EG_SECURE_2026_X9';
  if (e.parameter.secret !== API_SECRET) {
    return ContentService.createTextOutput("Error: 403 Forbidden - Invalid API Secret Key.");
  }
  
  try {
    if (sheetId) {
      ss = SpreadsheetApp.openById(sheetId);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
  } catch (err) {
    return ContentService.createTextOutput("Error: Cannot open Spreadsheet. Ensure the Spreadsheet ID inside your app is correct.");
  }
  
  if (action === 'save_client') {
    return dynamicAppend(ss, "Clients", e.parameter, {
      "Client ID": "id",
      "Name": "name",
      "Mobile": "mobile",
      "State": "state",
      "City": "city",
      "Sector": "sector",
      "Address": "address",
      "Email": "email",
      "Trade Name": "tradeName",
      "Payment Terms": "paymentTerms",
      "Created By": "createdBy",
      "Created At": "createdAt",
      "Location": "location"
    });
  }
  
  if (action === 'confirmed_orders') {
    return dynamicAppend(ss, "Orders", e.parameter, {
      "Date": "date",
      "Order ID": "orderId",
      "Booked By": "bookedBy",
      "Status": "status",
      "Client": "client",
      "Priority": "priority",
      "Bag Type": "bagType",
      "Width": "width",
      "Length": "length",
      "Gusset": "gusset",
      "Qty": "qty",
      "Unit": "unit",
      "Rate": "rate",
      "Bag Color": "bagColor",
      "Loop Color": "loopColor",
      "Print Desc": "printDesc",
      "GSM": "gsm",
      "Stereo": "stereo",
      "Bill Type": "billType",
      "Material": "bagMaterial",
      "Quality": "bagQuality",
      "Fare": "fare",
      "Dispatch Condition": "dispatchCondition"
    });
  }

  if (action === 'update_order_status') {
    return updateRowStatus(ss, "Orders", "Order ID", e.parameter.orderId, "Status", e.parameter.status);
  }

  if (action === 'get_order_statuses') {
    return getStatuses(ss, "Orders", "Order ID", "Status");
  }
  
  if (action === 'log') {
    return dynamicAppend(ss, "Marketing Log", e.parameter, {
      "Timestamp": "ts",
      "Date": "date",
      "Marketing Person": "createdBy",
      "Area": "area",
      "Client Name": "client",
      "Sector": "sector",
      "Contact": "contact",
      "Discussion": "discussion",
      "Order Chance": "orderChance",
      "Followup Date": "followup",
      "Location": "location"
    });
  }
  
  if (action === 'attendance') {
    return dynamicAppend(ss, "Attendance", e.parameter, {
      "Date": "date",
      "Marketing Person": "user",
      "Activity": "status",
      "Time": "time",
      "Location": "location"
    });
  }
  
  return ContentService.createTextOutput("Unknown action");
}

function normalizeHeader(str) {
  // Removes spaces and special characters for a matching logic (e.g. "Bag Color" == "bagcolor")
  return String(str).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function dynamicAppend(ss, sheetName, params, fieldMapping) {
  var sheet = ss.getSheetByName(sheetName);
  var defaultHeaders = Object.keys(fieldMapping);
  
  // If the sheet doesn't exist, create it and create the bold headings
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(defaultHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, defaultHeaders.length).setBackground("#f3f3f3");
  }
  
  // Read the current headers that user has arranged in the first row
  var lastCol = sheet.getLastColumn() || 1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Give standard headers if sheet is entirely blank
  if (headers.length === 1 && String(headers[0]).trim() === "") {
    headers = defaultHeaders;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#f3f3f3");
  }
  
  // Setup mapping rule
  var normFieldMapping = {};
  for(var k in fieldMapping) {
    normFieldMapping[normalizeHeader(k)] = fieldMapping[k];
  }
  
  var rowData = [];
  var paramsFound = 0;
  
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]).trim();
    if (header === "") {
      rowData.push("");
      continue;
    }
    
    var normHeader = normalizeHeader(header);
    var paramKey = normFieldMapping[normHeader];
    
    // Put proper param value directly beneath this header column!
    if (paramKey && params[paramKey] !== undefined) {
      rowData.push(params[paramKey]);
      paramsFound++;
    } else {
      rowData.push(""); 
    }
  }
  
  sheet.appendRow(rowData);
  return ContentService.createTextOutput("Success: Saved " + paramsFound + " fields");
}

function updateRowStatus(ss, sheetName, idHeader, orderId, statusHeader, newStatus) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return ContentService.createTextOutput("Error: Sheet not found");

  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var idColIdx = -1;
  var statusColIdx = -1;

  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === normalizeHeader(idHeader)) idColIdx = i + 1;
    if (normalizeHeader(headers[i]) === normalizeHeader(statusHeader)) statusColIdx = i + 1;
  }

  if (idColIdx === -1 || statusColIdx === -1) {
    return ContentService.createTextOutput("Error: Required columns (ID or Status) not found");
  }

  var data = sheet.getRange(1, idColIdx, lastRow, 1).getValues();
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][0]) === String(orderId)) {
      sheet.getRange(j + 1, statusColIdx).setValue(newStatus);
      return ContentService.createTextOutput("Success: Status updated to " + newStatus);
    }
  }

  return ContentService.createTextOutput("Error: Order ID not found in sheet");
}

function getStatuses(ss, sheetName, idHeader, statusHeader) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = -1;
  var statusIdx = -1;

  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === normalizeHeader(idHeader)) idIdx = i;
    if (normalizeHeader(headers[i]) === normalizeHeader(statusHeader)) statusIdx = i;
  }

  var results = {};
  if (idIdx !== -1 && statusIdx !== -1) {
    for (var j = 1; j < data.length; j++) {
      var id = data[j][idIdx];
      var status = data[j][statusIdx];
      if (id) results[id] = status;
    }
  }

  return ContentService.createTextOutput(JSON.stringify(results)).setMimeType(ContentService.MimeType.JSON);
}