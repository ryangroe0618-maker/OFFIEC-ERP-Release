function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var payload = parsePayload_(e);

    var spreadsheetId = toText_(payload.spreadsheetId);
    var sheetName = toText_(payload.sheetName);
    var values = Array.isArray(payload.values) ? payload.values : [];
    var append = Boolean(payload.append);
    var clear = Boolean(payload.clear);

    if (!spreadsheetId) {
      throw new Error("spreadsheetId is required");
    }
    if (!sheetName) {
      throw new Error("sheetName is required");
    }
    if (!Array.isArray(values) || values.length === 0) {
      throw new Error("values is required");
    }

    var normalized = normalize2dArray_(values);
    if (normalized.length === 0 || normalized[0].length === 0) {
      throw new Error("values is empty");
    }

    lock.waitLock(30000);
    locked = true;

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    if (append) {
      if (sheet.getLastRow() > 0 && hasCompatibleHeader_(sheet, normalized[0])) {
        sheet.getRange(1, 1, 1, normalized[0].length).setValues([normalized[0]]);
        normalized = normalized.slice(1);
      }

      if (normalized.length === 0) {
        return jsonResponse_({
          ok: true,
          mode: "append",
          spreadsheetId: spreadsheetId,
          sheetName: sheetName,
          appendedRows: 0,
          totalRows: sheet.getLastRow(),
          cols: 0
        });
      }
    } else if (clear) {
      sheet.clearContents();
    }

    var startRow = append ? sheet.getLastRow() + 1 : 1;
    var startCol = 1;
    var rowCount = normalized.length;
    var colCount = normalized[0].length;

    sheet.getRange(startRow, startCol, rowCount, colCount).setValues(normalized);

    return jsonResponse_({
      ok: true,
      mode: append ? "append" : (clear ? "replace" : "write"),
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      rows: rowCount,
      appendedRows: append ? rowCount : 0,
      totalRows: sheet.getLastRow(),
      cols: colCount,
      append: append,
      clear: clear
    });
  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}


function doGet() {
  return jsonResponse_({
    ok: true,
    message: "brand_out_uploader is running"
  });
}


function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("No POST body");
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error("Invalid JSON payload");
  }
}


function normalize2dArray_(values) {
  var maxCols = 0;
  for (var i = 0; i < values.length; i += 1) {
    var row = Array.isArray(values[i]) ? values[i] : [values[i]];
    if (row.length > maxCols) {
      maxCols = row.length;
    }
  }

  return values.map(function(row) {
    var current = Array.isArray(row) ? row.slice() : [row];
    while (current.length < maxCols) {
      current.push("");
    }
    return current.map(function(cell) {
      return cell == null ? "" : String(cell);
    });
  });
}


function hasSameHeader_(sheet, header) {
  if (!Array.isArray(header) || header.length === 0) {
    return false;
  }
  var existing = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  for (var i = 0; i < header.length; i += 1) {
    if (toText_(existing[i]) !== toText_(header[i])) {
      return false;
    }
  }
  return true;
}


function hasCompatibleHeader_(sheet, header) {
  if (!Array.isArray(header) || header.length === 0) {
    return false;
  }
  var existingColCount = Math.min(sheet.getLastColumn(), header.length);
  if (existingColCount === 0) {
    return false;
  }
  var existing = sheet.getRange(1, 1, 1, existingColCount).getValues()[0];
  for (var i = 0; i < existingColCount; i += 1) {
    var existingText = toText_(existing[i]);
    if (existingText && existingText !== toText_(header[i])) {
      return false;
    }
  }
  return true;
}


function toText_(value) {
  return value == null ? "" : String(value).trim();
}


function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
