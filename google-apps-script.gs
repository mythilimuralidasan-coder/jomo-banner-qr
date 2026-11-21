var HEADERS = ["Timestamp","Lead ID","Name","Phone","Email","Property Type","Configuration","Requirements","Budget","Timeline","Details Completed","Source"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var d = JSON.parse(e.postData.contents);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);

    if (d.action === "update") {
      var last = sh.getLastRow();
      var ids = last > 1 ? sh.getRange(2, 2, last - 1, 1).getValues() : [];
      var row = -1;
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(d.id)) { row = i + 2; break; }
      }
      if (row < 0) throw new Error("Lead not found");
      sh.getRange(row, 5, 1, 6).setValues([[d.email || "", d.property_type || "", d.configuration || "", d.requirements || "", d.budget || "", d.timeline || ""]]);
      sh.getRange(row, 11).setValue("Yes");
    } else {
      sh.appendRow([d.timestamp || new Date(), d.id || "", d.name || "", "", "", "", "", "", "", "", "No", d.source || ""]);
      sh.getRange(sh.getLastRow(), 4).setNumberFormat("@").setValue(d.phone || "");
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
