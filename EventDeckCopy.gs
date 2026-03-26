/**
 * EVENT SYNC TOOL - V2.5 (FINAL MASTER WITH ERROR REPORTING)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('▶️ Event Sync')
    .addItem('Step 1: Backup Decks to Drive', 'runEventBackup')
    .addSeparator()
    .addItem('Clear All Backup Links (Fresh Start)', 'clearBackupLinks')
    .addItem('Help / Instructions', 'showInstructions')
    .addToUi();
}

function runEventBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('Root Folder Required', 'Paste the Folder URL or ID:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  let input = response.getResponseText().trim();
  let rootFolderId = input;
  if (input.includes("folders/")) {
    rootFolderId = input.split("folders/")[1].split("?")[0].split("/")[0];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const totalRows = data.length - 1;

  const colSummary = headers.indexOf("Summary");
  const colTime = headers.indexOf("Time");
  const colRoom = headers.indexOf("Room");
  const colLink = headers.indexOf("Session Final Deck Link");
  let colBackup = headers.indexOf("Local Backup Link");
  
  if (colBackup === -1) {
    colBackup = headers.length;
    sheet.getRange(1, colBackup + 1).setValue("Local Backup Link");
  }

  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(rootFolderId);
  } catch (e) {
    ui.alert('Error: Access Denied to Root Folder.');
    return;
  }

  let successCount = 0;
  let errorLog = []; // To store specific details for the final report

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;
    const summary = row[colSummary] || "Unknown Title";
    
    // Progress Toast
    let progress = Math.round((i / totalRows) * 100);
    ss.toast(`Progress: ${progress}%`, `Processing Row ${rowNum}`, 2);

    const deckUrl = row[colLink];
    const roomName = row[colRoom] ? row[colRoom].toString().trim() : "";
    const existingBackup = row[colBackup];

    if (!deckUrl || existingBackup || !roomName) continue;

    try {
      const subFolders = rootFolder.getFoldersByName(roomName);
      let targetFolder = null;
      while (subFolders.hasNext()) {
        let f = subFolders.next();
        if (f.getName() === roomName) { targetFolder = f; break; }
      }

      if (!targetFolder) {
        throw new Error("Room folder '" + roomName + "' not found");
      }

      const fileIdMatch = deckUrl.match(/[-\w]{25,}/);
      if (!fileIdMatch) throw new Error("Invalid URL Format");
      
      const fileId = fileIdMatch[0];
      const originalFile = DriveApp.getFileById(fileId);
      const originalName = originalFile.getName();
      let extension = originalName.includes('.') ? originalName.substring(originalName.lastIndexOf('.')) : "";

      let timeStr = row[colTime].toString().trim();
      while (timeStr.length < 4) timeStr = "0" + timeStr;
      const cleanSummary = summary.toString().replace(/[\\\/\:\*\?\"\<\>\|]/g, "");
      const newFileName = timeStr + "_" + cleanSummary + extension;

      const copy = originalFile.makeCopy(newFileName, targetFolder);
      sheet.getRange(rowNum, colBackup + 1).setValue(copy.getUrl());
      
      successCount++;
      SpreadsheetApp.flush(); 

    } catch (e) {
      errorLog.push(`Row ${rowNum} (${summary}): ${e.message}`);
      sheet.getRange(rowNum, colBackup + 1).setValue("⚠️ Error: " + e.message);
    }
  }

  // FINAL REPORT
  let finalMessage = `Sync Complete.\n✅ Success: ${successCount}\n❌ Failed: ${errorLog.length}`;
  if (errorLog.length > 0) {
    finalMessage += `\n\nERRORS FOUND:\n------------------\n${errorLog.slice(0, 10).join('\n')}`;
    if (errorLog.length > 10) finalMessage += `\n...and ${errorLog.length - 10} more.`;
  }
  
  ui.alert('Final Report', finalMessage, ui.ButtonSet.OK);
}

function clearBackupLinks() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('🚨 CRITICAL CONFIRMATION', 
    'This will wipe all links in the "Local Backup Link" column. This cannot be undone. \n\nAre you sure you want to start fresh?', 
    ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getDataRange().getValues()[0];
    const colBackup = headers.indexOf("Local Backup Link");
    if (colBackup !== -1) {
      sheet.getRange(2, colBackup + 1, sheet.getLastRow(), 1).clearContent();
      ui.alert('Links Cleared. You can now run a fresh sync.');
    }
  }
}

function showInstructions() {
  SpreadsheetApp.getUi().alert('Instructions', 
    '1. Ensure Room folders match Column G exactly.\n2. Paste Root URL when prompted.\n3. Errors will be reported by Row Number at the end.', 
    SpreadsheetApp.getUi().ButtonSet.OK);
}
