function checkMRPChangesAndRun() {
  var props = PropertiesService.getScriptProperties();
  var prevHashes = JSON.parse(props.getProperty("MRP_HASHES") || "{}");
  var currHashes = {};
  var changed = false;

  var files = DriveApp.getFolderById(MRP_SOURCE_FOLDER_ID).getFiles();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var fileId = file.getId();

    try {
      var spreadsheet = SpreadsheetApp.openById(fileId);
      var mrpReqSheet = spreadsheet.getSheetByName(MRP_SOURCE_TAB_NAME);
      if (!mrpReqSheet) continue;

      var data = mrpReqSheet.getDataRange().getValues();
      var hash = computeHash(data);
      currHashes[fileId] = hash;

      if (prevHashes[fileId] !== hash) {
        changed = true;
      }
    } catch (e) {
      Logger.log("Error hashing file: " + file.getName() + " - " + e.message);
      continue;
    }
  }

  // Detect deleted files: any fileId in prevHashes but not in currHashes
  for (var oldId in prevHashes) {
    if (!(oldId in currHashes)) {
      changed = true;
    }
  }

  if (!changed) {
    Logger.log("No changes detected. Skipping consolidation.");
    return;
  }

  Logger.log("Changes detected. Running consolidation.");

  try {
    consolidateMRPReqSheet();
    // Only persist new hash state after a successful run,
    // so a failed run gets retried on the next check.
    props.setProperty("MRP_HASHES", JSON.stringify(currHashes));
  } catch (e) {
    Logger.log("Consolidation failed, hashes not updated: " + e.message);
  }
}

function computeHash(data) {
  var str = JSON.stringify(data);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  return digest.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function installMRPHashTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkMRPChangesAndRun") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("checkMRPChangesAndRun")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Trigger installed: checkMRPChangesAndRun every 5 minutes.");
}
