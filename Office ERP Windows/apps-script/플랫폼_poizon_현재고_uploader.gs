function doGet() {
  return jsonOutput({
    ok: true,
    message: "web app is running"
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const spreadsheetId = body.spreadsheetId;
    const sheetName = body.sheetName;
    const values = body.values;
    const clear = body.clear === true;
    const clearRangeOnly = body.clearRangeOnly === true;

    if (!spreadsheetId || !sheetName || !Array.isArray(values) || values.length === 0) {
      return jsonOutput({
        ok: false,
        message: "spreadsheetId, sheetName, values가 필요합니다."
      });
    }

    const rowCount = values.length;
    const colCount = Array.isArray(values[0]) ? values[0].length : 0;
    if (colCount <= 0) {
      return jsonOutput({
        ok: false,
        message: "업로드할 열 개수가 0입니다."
      });
    }

    lock.waitLock(30000);
    locked = true;

    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const startRow = Math.max(toPositiveInt(body.startRow, 1), 1);
    const startCol = Math.max(toPositiveInt(body.startCol, 1), 1);

    if (clear) {
      if (clearRangeOnly) {
        const clearCols = Math.max(toPositiveInt(body.clearCols, colCount), colCount);
        const existingRows = Math.max(sheet.getLastRow() - startRow + 1, 0);
        const clearRows = Math.max(toPositiveInt(body.clearRows, 0), existingRows, rowCount);
        if (clearRows > 0 && clearCols > 0) {
          sheet.getRange(startRow, startCol, clearRows, clearCols).clearContent();
        }
      } else {
        sheet.clearContents();
      }
    }

    sheet.getRange(startRow, startCol, rowCount, colCount).setValues(values);

    return jsonOutput({
      ok: true,
      message: "업로드 완료",
      rows: rowCount,
      cols: colCount,
      clear: clear,
      clearRangeOnly: clearRangeOnly,
      startRow: startRow,
      startCol: startCol
    });

  } catch (err) {
    return jsonOutput({
      ok: false,
      message: String(err)
    });
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

function toPositiveInt(value, fallback) {
  const number = parseInt(value, 10);
  if (!isFinite(number) || number < 0) {
    return fallback;
  }
  return number;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
