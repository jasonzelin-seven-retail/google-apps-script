// /*******************************
//  * KONFIGURASI GLOBAL
//  *******************************/
// const SPREADSHEET_ID = '1R_-hi6Ncp1QwQyfc3aRbAzXuJO1SdoqzCnGWK8YSBYE';

// const SOURCE_SHEET_NAME = 'Schedule';
// const DEST_SHEET_NAME   = 'Folder Structure';

// const NPS_ROOT_ID = '0ANco5d7U6OmtUk9PVA';

// const BATCH_SIZE = 50;
// const MAX_EXECUTION_TIME = 330000;
// const LIST_BATCH_LIMIT = 500;
// const DAILY_UPDATE_HOUR = 12;



// /*******************************
//  * MENU
//  *******************************/
// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('📁 NPS Folder Manager')
//     .addItem('🆕 Create Folders (Auto-Continue)', 'showCreateFoldersDialog')
//     .addItem('🔄 Update Folders (Append Missing)', 'showUpdateFoldersDialog')
//     .addItem('📊 Update File Count (Refresh All)', 'updateAllFileCounts')
//     .addItem('▶️ Resume Update Count', 'resumeUpdateCount')
//     .addItem('⏰ Setup Daily Auto-Update', 'setupDailyTrigger')
//     .addItem('🛑 Remove Daily Auto-Update', 'removeDailyTrigger')
//     .addSeparator()
//     .addItem('▶️ Resume Processing', 'resumeProcessing')
//     .addItem('🔄 Reset Progress', 'resetProgress')
//     .addItem('📊 List All Structure', 'startListAllStructure')
//     .addItem('🔍 Debug: Check Data', 'debugCheckSourceData')
//     .addToUi();
// }



// /*******************************
//  * SETUP DAILY TRIGGER
//  *******************************/
// function setupDailyTrigger() {
//   removeDailyTrigger();
  
//   ScriptApp.newTrigger('dailyUpdateFileCounts')
//     .timeBased()
//     .atHour(DAILY_UPDATE_HOUR)
//     .everyDays(1)
//     .create();
  
//   SpreadsheetApp.getUi().alert(
//     `✅ Daily Auto-Update berhasil diaktifkan!\n\n` +
//     `Script akan update file count setiap hari jam ${DAILY_UPDATE_HOUR}:00 WIB.`
//   );
// }

// function removeDailyTrigger() {
//   const triggers = ScriptApp.getProjectTriggers();
//   triggers.forEach(trigger => {
//     if (trigger.getHandlerFunction() === 'dailyUpdateFileCounts') {
//       ScriptApp.deleteTrigger(trigger);
//     }
//   });
  
//   SpreadsheetApp.getUi().alert('✅ Daily Auto-Update berhasil dihapus.');
// }

// function dailyUpdateFileCounts() {
//   updateAllFileCounts();
// }



// /*******************************
//  * UPDATE FILE COUNTS - BATCH VERSION (REVISED)
//  *******************************/
// function updateAllFileCounts() {
//   const props = PropertiesService.getScriptProperties();
//   props.deleteProperty('UPDATE_LAST_ROW');
//   props.deleteProperty('UPDATE_TOTAL_UPDATED');
  
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const destSheet = ss.getSheetByName(DEST_SHEET_NAME);
  
//   if (!destSheet) {
//     SpreadsheetApp.getUi().alert('❌ Sheet tidak ditemukan!');
//     return;
//   }
  
//   const lastRow = destSheet.getLastRow();
//   if (lastRow <= 1) {
//     SpreadsheetApp.getUi().alert('⚠️ Tidak ada data untuk diupdate.');
//     return;
//   }
  
//   props.setProperty('UPDATE_LAST_ROW', '1');
//   props.setProperty('UPDATE_TOTAL_UPDATED', '0');
  
//   const result = continueUpdateFileCounts();
  
//   if (!result.startsWith('✅ File Count Update Selesai')) {
//     createWorkerTrigger_('continueUpdateFileCounts');
//     SpreadsheetApp.getUi().alert(result + '\n\n✅ Trigger otomatis sudah diaktifkan.');
//   } else {
//     deleteWorkerTrigger_('continueUpdateFileCounts');
//     SpreadsheetApp.getUi().alert(result);
//   }
// }


// function continueUpdateFileCounts() {
//   const startTime = new Date().getTime();
//   const props = PropertiesService.getScriptProperties();
  
//   try {
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     const destSheet = ss.getSheetByName(DEST_SHEET_NAME);
    
//     const totalRows = destSheet.getLastRow();
//     if (totalRows <= 1) {
//       return '⚠️ Tidak ada data untuk diupdate.';
//     }
    
//     let lastProcessedRow = Number(props.getProperty('UPDATE_LAST_ROW')) || 1;
//     let totalUpdated = Number(props.getProperty('UPDATE_TOTAL_UPDATED')) || 0;
    
//     while (lastProcessedRow < totalRows) {
//       // Check timeout - sisakan 30 detik untuk safety
//       if (new Date().getTime() - startTime > (MAX_EXECUTION_TIME - 30000)) {
//         props.setProperty('UPDATE_LAST_ROW', String(lastProcessedRow));
//         props.setProperty('UPDATE_TOTAL_UPDATED', String(totalUpdated));
        
//         return `⏱️ UPDATE IN PROGRESS (PAUSE)\n` +
//                `Processed: ${lastProcessedRow} / ${totalRows - 1}\n` +
//                `Total Updated: ${totalUpdated}\n\n` +
//                `➡️ Worker trigger akan melanjutkan otomatis dalam 1 menit.`;
//       }
      
//       const startRow = lastProcessedRow + 1;
//       const endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows);
//       const numRows = endRow - startRow + 1;
      
//       const data = destSheet.getRange(startRow, 1, numRows, 4).getValues();
//       const updateBuffer = [];
      
//       for (let i = 0; i < data.length; i++) {
//         const folderUrl = data[i][3]; // Column D: Folder URL
        
//         if (!folderUrl) {
//           updateBuffer.push([0, 'No URL', 0, 'No URL']);
//           continue;
//         }
        
//         try {
//           const folderId = extractFolderIdFromUrl(folderUrl);
//           const mainFolder = DriveApp.getFolderById(folderId);
          
//           const personalChatData = getFilesDataFromSubfolder(mainFolder, 'Personal Chat');
//           const groupChatData = getFilesDataFromSubfolder(mainFolder, 'Group Chat');
          
//           updateBuffer.push([
//             personalChatData.count,
//             personalChatData.links,
//             groupChatData.count,
//             groupChatData.links
//           ]);
          
//           totalUpdated++;
          
//         } catch (e) {
//           updateBuffer.push([0, 'Error: ' + e.message, 0, 'Error: ' + e.message]);
//           Logger.log(`Error processing row ${startRow + i}: ${e.message}`);
//         }
//       }
      
//       // Write batch to sheet
//       if (updateBuffer.length > 0) {
//         destSheet.getRange(startRow, 5, updateBuffer.length, 4).setValues(updateBuffer);
//       }
      
//       lastProcessedRow = endRow;
//       props.setProperty('UPDATE_LAST_ROW', String(lastProcessedRow));
//       props.setProperty('UPDATE_TOTAL_UPDATED', String(totalUpdated));
//     }
    
//     // Selesai semua
//     const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
//     destSheet.getRange(1, 9).setValue(`Last Updated: ${timestamp}`);
    
//     props.deleteProperty('UPDATE_LAST_ROW');
//     props.deleteProperty('UPDATE_TOTAL_UPDATED');
//     deleteWorkerTrigger_('continueUpdateFileCounts');
    
//     return `✅ File Count Update Selesai!\n\n` +
//            `Total Rows Updated: ${totalUpdated}\n` +
//            `Timestamp: ${timestamp}`;
    
//   } catch (err) {
//     deleteWorkerTrigger_('continueUpdateFileCounts');
//     return `❌ ERROR: ${err.message}\n\n${err.stack}`;
//   }
// }


// function resumeUpdateCount() {
//   const props = PropertiesService.getScriptProperties();
//   const lastRow = props.getProperty('UPDATE_LAST_ROW');
  
//   if (!lastRow) {
//     SpreadsheetApp.getUi().alert('⚠️ Tidak ada proses update yang sedang berjalan.');
//     return;
//   }
  
//   const result = continueUpdateFileCounts();
//   SpreadsheetApp.getUi().alert(result);
// }


// function extractFolderIdFromUrl(url) {
//   const match = url.match(/[-\w]{25,}/);
//   return match ? match[0] : null;
// }


// function getFilesDataFromSubfolder(parentFolder, subfolderName) {
//   try {
//     const subfolders = parentFolder.getFoldersByName(subfolderName);
//     if (!subfolders.hasNext()) {
//       return { count: 0, links: 'Folder tidak ditemukan' };
//     }
    
//     const subfolder = subfolders.next();
//     const filesFolder = subfolder.getFoldersByName('Files');
    
//     if (!filesFolder.hasNext()) {
//       return { count: 0, links: 'Folder Files tidak ditemukan' };
//     }
    
//     const files = filesFolder.next().getFiles();
//     const fileLinks = [];
//     let count = 0;
    
//     // Batasi iterasi untuk menghindari timeout pada folder besar
//     let maxIterations = 500;
    
//     while (files.hasNext() && count < maxIterations) {
//       try {
//         const file = files.next();
//         fileLinks.push(`${file.getName()} (${file.getUrl()})`);
//         count++;
//       } catch (e) {
//         Logger.log(`Error accessing file: ${e.message}`);
//         break;
//       }
//     }
    
//     return {
//       count: count,
//       links: count > 0 ? fileLinks.join('\n') : 'Tidak ada file'
//     };
    
//   } catch (e) {
//     return { count: 0, links: `Error: ${e.message}` };
//   }
// }



// /*******************************
//  * DIALOG UPDATE FOLDERS
//  *******************************/
// function showUpdateFoldersDialog() {
//   const html = HtmlService.createHtmlOutput(`
//     <style>
//       body { font-family: Arial; padding: 20px; }
//       input { width: 100%; padding: 8px; margin: 10px 0; box-sizing: border-box; }
//       button { background: #34a853; color: white; padding: 10px 20px; border: none; cursor: pointer; width: 100%; margin-top: 10px; }
//       button:hover { background: #2d8e47; }
//       label { font-weight: bold; display: block; margin-top: 10px; }
//       .tip { background: #e8f4f8; padding: 10px; border-radius: 4px; font-size: 12px; margin-top: 10px; }
//     </style>
    
//     <h3>🔄 Update Folders (Append Missing)</h3>
    
//     <div class="tip">
//       📥 Fitur ini akan:<br>
//       1. Membaca semua folder yang ada di Drive<br>
//       2. Membandingkan dengan data di sheet<br>
//       3. Menambahkan folder yang belum tercatat ke baris terakhir<br><br>
//       ⚠️ Tidak akan membuat folder baru, hanya mencatat yang sudah ada.
//     </div>
    
//     <label>Center Filter (optional):</label>
//     <input type="text" id="centerUpdate" placeholder="Kosongkan untuk semua">
    
//     <button onclick="startUpdate()">🔄 Start Update</button>
//     <div id="statusUpdate" style="margin-top: 15px; padding: 10px; display: none; white-space: pre-wrap; font-size: 12px;"></div>
    
//     <script>
//       function startUpdate() {
//         const center = document.getElementById('centerUpdate').value.trim();
//         const status = document.getElementById('statusUpdate');
//         status.style.display = 'block';
//         status.style.background = '#fff3cd';
//         status.innerHTML = '⏳ Processing... Please wait...';
        
//         google.script.run
//           .withSuccessHandler(function(result) {
//             status.style.background = result.includes('❌') ? '#f8d7da' : '#d4edda';
//             status.innerHTML = result;
//           })
//           .withFailureHandler(function(error) {
//             status.style.background = '#f8d7da';
//             status.innerHTML = '❌ Error: ' + error;
//           })
//           .startUpdateFolders(center);
//       }
//     </script>
//   `).setWidth(500).setHeight(600);
  
//   SpreadsheetApp.getUi().showModalDialog(html, 'Update Folders');
// }



// /*******************************
//  * UPDATE FOLDERS - CORE FUNCTION
//  *******************************/
// function startUpdateFolders(centerFilter = '') {
//   try {
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     let destSheet = ss.getSheetByName(DEST_SHEET_NAME);
    
//     if (!destSheet) {
//       destSheet = ss.insertSheet(DEST_SHEET_NAME);
//       setupSheetHeaders(destSheet);
//     }
    
//     const lastRow = destSheet.getLastRow();
//     const existingData = lastRow > 1 ? destSheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];
//     const existingUrls = new Set(existingData.map(row => row[3]));
    
//     const npsFolder = DriveApp.getFolderById(NPS_ROOT_ID);
//     const outputBuffer = [];
//     let totalChecked = 0;
//     let totalAdded = 0;
    
//     const centerFolders = npsFolder.getFolders();
//     while (centerFolders.hasNext()) {
//       const centerFolder = centerFolders.next();
//       const centerName = centerFolder.getName();
      
//       if (centerFilter && !centerName.includes(centerFilter)) continue;
      
//       const weekFolders = centerFolder.getFolders();
//       while (weekFolders.hasNext()) {
//         const weekFolder = weekFolders.next();
//         const weekName = weekFolder.getName();
        
//         const saFolders = weekFolder.getFolders();
//         while (saFolders.hasNext()) {
//           const saFolder = saFolders.next();
//           const saName = saFolder.getName();
//           const folderUrl = saFolder.getUrl();
          
//           totalChecked++;
          
//           if (!existingUrls.has(folderUrl)) {
//             const personalChatData = getFilesDataFromSubfolder(saFolder, 'Personal Chat');
//             const groupChatData = getFilesDataFromSubfolder(saFolder, 'Group Chat');
            
//             outputBuffer.push([
//               centerName,
//               weekName,
//               saName,
//               folderUrl,
//               personalChatData.count,
//               personalChatData.links,
//               groupChatData.count,
//               groupChatData.links
//             ]);
//             totalAdded++;
//           }
//         }
//       }
//     }
    
//     if (outputBuffer.length > 0) {
//       const nextRow = destSheet.getLastRow() + 1;
//       destSheet.getRange(nextRow, 1, outputBuffer.length, 8).setValues(outputBuffer);
//     }
    
//     return `✅ UPDATE SELESAI!\n\n` +
//            `Total Folders Checked: ${totalChecked}\n` +
//            `Total Added (Missing): ${totalAdded}\n` +
//            `Total Already Existed: ${totalChecked - totalAdded}`;
    
//   } catch (err) {
//     return `❌ ERROR: ${err.message}\n\n${err.stack}`;
//   }
// }



// /*******************************
//  * DIALOG CREATE FOLDERS
//  *******************************/
// function showCreateFoldersDialog() {
//   const html = HtmlService.createHtmlOutput(`
//     <style>
//       body { font-family: Arial; padding: 20px; }
//       input, select { width: 100%; padding: 8px; margin: 10px 0; box-sizing: border-box; }
//       button { background: #4285f4; color: white; padding: 10px 20px; border: none; cursor: pointer; width: 100%; margin-top: 10px; }
//       button:hover { background: #357ae8; }
//       label { font-weight: bold; display: block; margin-top: 10px; }
//       .tip { background: #e8f4f8; padding: 10px; border-radius: 4px; font-size: 12px; margin-top: 10px; }
//       .warn { background: #fff3cd; padding: 10px; border-radius: 4px; font-size: 12px; margin-top: 10px; }
//     </style>
    
//     <h3>📁 Auto Create Folders</h3>
    
//     <div class="tip">
//       🚀 <b>Auto-Continue Mode</b><br>
//       Script akan otomatis proses semua batch sampai selesai.<br>
//       Per batch: ${BATCH_SIZE} rows.<br><br>
//       Struktur: Center / Year-Week / SA Name - Class ID / {Group Chat, Personal Chat} / Files
//     </div>
    
//     <label>Year:</label>
//     <input type="number" id="year" value="2026" min="2024" max="2030">
    
//     <label>Week Range:</label>
//     <div style="display: flex; gap: 10px;">
//       <div style="flex: 1;">
//         <label style="font-size: 12px;">From Week:</label>
//         <input type="number" id="weekFrom" value="5" min="1" max="52">
//       </div>
//       <div style="flex: 1;">
//         <label style="font-size: 12px;">To Week:</label>
//         <input type="number" id="weekTo" value="28" min="1" max="52">
//       </div>
//     </div>
    
//     <label>Center (optional):</label>
//     <input type="text" id="center" placeholder="Kosongkan untuk semua">
    
//     <label>Mode:</label>
//     <select id="mode">
//       <option value="append">Append (Tambah ke existing data)</option>
//       <option value="fresh">Fresh Start (Clear semua data)</option>
//     </select>
    
//     <div class="warn">
//       ⚠️ Mode "Append" akan menambahkan data baru ke baris terakhir.<br>
//       Mode "Fresh Start" akan menghapus semua data existing terlebih dahulu.<br><br>
//       ✅ Duplikasi otomatis dicegah berdasarkan Folder URL.
//     </div>
    
//     <button onclick="startProcessing()">🚀 Start Processing</button>
//     <div id="status" style="margin-top: 15px; padding: 10px; display: none; white-space: pre-wrap; font-size: 12px;"></div>
    
//     <script>
//       function startProcessing() {
//         const year = document.getElementById('year').value;
//         const weekFrom = parseInt(document.getElementById('weekFrom').value);
//         const weekTo = parseInt(document.getElementById('weekTo').value);
//         const center = document.getElementById('center').value.trim();
//         const mode = document.getElementById('mode').value;
        
//         if (weekFrom > weekTo) {
//           alert('Week From harus ≤ Week To!');
//           return;
//         }
        
//         const status = document.getElementById('status');
//         status.style.display = 'block';
//         status.style.background = '#fff3cd';
//         status.innerHTML = '⏳ Initializing... Please wait...';
        
//         google.script.run
//           .withSuccessHandler(function(result) {
//             status.style.background = result.includes('❌') ? '#f8d7da' : '#d4edda';
//             status.innerHTML = result;
//           })
//           .withFailureHandler(function(error) {
//             status.style.background = '#f8d7da';
//             status.innerHTML = '❌ Error: ' + error;
//           })
//           .startBatchProcessing(year, weekFrom, weekTo, center, mode);
//       }
//     </script>
//   `).setWidth(500).setHeight(850);
  
//   SpreadsheetApp.getUi().showModalDialog(html, 'Create Folders - Auto Continue');
// }



// /*******************************
//  * SETUP SHEET HEADERS
//  *******************************/
// function setupSheetHeaders(sheet) {
//   sheet.getRange(1, 1, 1, 9).setValues([[
//     'Center', 
//     'Week Folder', 
//     'SA Name - Class Code', 
//     'Folder URL',
//     'Personal Chat Files',
//     'Personal Chat Links',
//     'Group Chat Files',
//     'Group Chat Links',
//     'Last Updated'
//   ]]);
//   sheet.getRange(1, 1, 1, 9)
//     .setBackground('#4a86e8')
//     .setFontColor('#ffffff')
//     .setFontWeight('bold')
//     .setWrap(true);
  
//   sheet.setColumnWidth(1, 100);
//   sheet.setColumnWidth(2, 120);
//   sheet.setColumnWidth(3, 180);
//   sheet.setColumnWidth(4, 250);
//   sheet.setColumnWidth(5, 80);
//   sheet.setColumnWidth(6, 300);
//   sheet.setColumnWidth(7, 80);
//   sheet.setColumnWidth(8, 300);
//   sheet.setColumnWidth(9, 150);
// }



// /*******************************
//  * CREATE FOLDERS - START BATCH
//  *******************************/
// function startBatchProcessing(year, weekFrom, weekTo, centerFilter = '', mode = 'append') {
//   const props = PropertiesService.getScriptProperties();
//   props.deleteAllProperties();
  
//   props.setProperty('year', String(year));
//   props.setProperty('weekFrom', String(weekFrom));
//   props.setProperty('weekTo', String(weekTo));
//   props.setProperty('centerFilter', centerFilter);
//   props.setProperty('lastProcessedRow', '1');
//   props.setProperty('totalProcessed', '0');
//   props.setProperty('totalCreated', '0');
//   props.setProperty('totalSkipped', '0');
//   props.setProperty('totalErrors', '0');
  
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   let destSheet = ss.getSheetByName(DEST_SHEET_NAME);
  
//   if (!destSheet) {
//     destSheet = ss.insertSheet(DEST_SHEET_NAME);
//   }
  
//   if (mode === 'fresh') {
//     destSheet.clearContents();
//   }
  
//   if (destSheet.getLastRow() === 0 || mode === 'fresh') {
//     setupSheetHeaders(destSheet);
//   }
  
//   const result = processAllBatches();

//   if (!result.startsWith('✅ SELESAI SEMUA')) {
//     createWorkerTrigger_('processAllBatches');
//   } else {
//     deleteWorkerTrigger_('processAllBatches');
//   }

//   return result;
// }



// /*******************************
//  * CREATE FOLDERS - CORE AUTO BATCH
//  *******************************/
// function processAllBatches() {
//   const startTime = new Date().getTime();
//   const props = PropertiesService.getScriptProperties();
  
//   const year = Number(props.getProperty('year'));
//   const weekFrom = Number(props.getProperty('weekFrom'));
//   const weekTo = Number(props.getProperty('weekTo'));
//   const centerFilter = props.getProperty('centerFilter') || '';
  
//   try {
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
//     if (!sourceSheet) return `❌ Sheet "${SOURCE_SHEET_NAME}" tidak ditemukan`;
    
//     const destSheet = ss.getSheetByName(DEST_SHEET_NAME);
//     if (!destSheet) return `❌ Sheet "${DEST_SHEET_NAME}" tidak ditemukan`;
    
//     const lastRow = destSheet.getLastRow();
//     const existingUrls = new Set();
    
//     if (lastRow > 1) {
//       const existingData = destSheet.getRange(2, 1, lastRow - 1, 4).getValues();
//       existingData.forEach(row => {
//         if (row[3]) existingUrls.add(row[3]);
//       });
//     }
    
//     const totalRows = sourceSheet.getLastRow();
//     const lastCol = Math.max(sourceSheet.getLastColumn(), 15);
//     const npsFolder = DriveApp.getFolderById(NPS_ROOT_ID);
//     const regex = /^(\d{4})\s*-\s*Week\s*(\d+)$/i;
    
//     let lastProcessedRow = Number(props.getProperty('lastProcessedRow'));
//     let totalProcessed = Number(props.getProperty('totalProcessed'));
//     let totalCreated = Number(props.getProperty('totalCreated'));
//     let totalSkipped = Number(props.getProperty('totalSkipped'));
//     let totalErrors = Number(props.getProperty('totalErrors'));
//     let totalDuplicate = 0;
    
//     while (lastProcessedRow < totalRows) {
//       if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
//         props.setProperty('lastProcessedRow', String(lastProcessedRow));
//         props.setProperty('totalProcessed', String(totalProcessed));
//         props.setProperty('totalCreated', String(totalCreated));
//         props.setProperty('totalSkipped', String(totalSkipped));
//         props.setProperty('totalErrors', String(totalErrors));
//         return `⏱️ TIMEOUT - Progress disimpan (worker trigger akan lanjut otomatis)\n` +
//                `Processed: ${lastProcessedRow} / ${totalRows}\n` +
//                `Created: ${totalCreated} | Skipped: ${totalSkipped} | Duplicate: ${totalDuplicate} | Errors: ${totalErrors}`;
//       }
      
//       const startRow = lastProcessedRow + 1;
//       const endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows);
//       const numRows = endRow - startRow + 1;
      
//       const data = sourceSheet.getRange(startRow, 1, numRows, lastCol).getValues();
//       const outputBuffer = [];
      
//       for (let i = 0; i < data.length; i++) {
//         const centerCode = data[i][0];
//         const saName = data[i][1];
//         const classId = data[i][2];
//         const weekBlasting = data[i][11];
        
//         if (!centerCode || !saName || !classId || !weekBlasting) continue;
        
//         const saNameClassCode = `${saName} - ${classId}`;
        
//         const normalized = weekBlasting.toString().trim().replace(/\s+/g, ' ');
//         const match = normalized.match(regex);
//         if (!match) continue;
        
//         const dataYear = Number(match[1]);
//         const dataWeek = Number(match[2]);
        
//         if (centerFilter && centerCode.trim() !== centerFilter.trim()) continue;
//         if (dataYear !== year) continue;
//         if (dataWeek < weekFrom || dataWeek > weekTo) continue;
        
//         totalProcessed++;
        
//         try {
//           const result = createNestedFolderStructure(
//             npsFolder,
//             centerCode,
//             dataYear,
//             dataWeek,
//             saNameClassCode
//           );
          
//           if (result.created) totalCreated++; else totalSkipped++;
          
//           const folderUrl = result.folder.getUrl();
          
//           if (!existingUrls.has(folderUrl)) {
//             const personalChatData = getFilesDataFromSubfolder(result.folder, 'Personal Chat');
//             const groupChatData = getFilesDataFromSubfolder(result.folder, 'Group Chat');
            
//             outputBuffer.push([
//               centerCode,
//               `${dataYear}-Week ${dataWeek}`,
//               saNameClassCode,
//               folderUrl,
//               personalChatData.count,
//               personalChatData.links,
//               groupChatData.count,
//               groupChatData.links
//             ]);
            
//             existingUrls.add(folderUrl);
//           } else {
//             totalDuplicate++;
//             Logger.log(`Duplicate skipped: ${folderUrl}`);
//           }
          
//         } catch (e) {
//           totalErrors++;
//           Logger.log(`Error row ${startRow + i}: ${e.message}`);
//         }
//       }
      
//       if (outputBuffer.length > 0) {
//         const outStartRow = destSheet.getLastRow() + 1;
//         destSheet.getRange(outStartRow, 1, outputBuffer.length, 8).setValues(outputBuffer);
//       }
      
//       lastProcessedRow = endRow;
      
//       props.setProperty('lastProcessedRow', String(lastProcessedRow));
//       props.setProperty('totalProcessed', String(totalProcessed));
//       props.setProperty('totalCreated', String(totalCreated));
//       props.setProperty('totalSkipped', String(totalSkipped));
//       props.setProperty('totalErrors', String(totalErrors));
//     }
    
//     props.deleteAllProperties();
//     deleteWorkerTrigger_('processAllBatches');
    
//     let summary = `✅ SELESAI SEMUA!\n\n` +
//                   `Total Processed: ${totalProcessed}\n` +
//                   `Created: ${totalCreated}\n` +
//                   `Skipped: ${totalSkipped}\n` +
//                   `Errors: ${totalErrors}`;
    
//     if (totalDuplicate > 0) {
//       summary += `\n\n⚠️ Duplicate Prevented: ${totalDuplicate}`;
//     }
    
//     return summary;
    
//   } catch (err) {
//     deleteWorkerTrigger_('processAllBatches');
//     return `❌ ERROR: ${err.message}\n\n${err.stack}`;
//   }
// }



// /*******************************
//  * RESUME, RESET, DEBUG
//  *******************************/
// function resumeProcessing() {
//   const props = PropertiesService.getScriptProperties();
//   const lastRow = props.getProperty('lastProcessedRow');
  
//   if (!lastRow) {
//     SpreadsheetApp.getUi().alert('⚠️ Tidak ada proses yang sedang berjalan.');
//     return;
//   }
  
//   const result = processAllBatches();
//   SpreadsheetApp.getUi().alert(result);
// }

// function resetProgress() {
//   PropertiesService.getScriptProperties().deleteAllProperties();
//   deleteWorkerTrigger_('processAllBatches');
//   deleteWorkerTrigger_('continueListAllStructure');
//   deleteWorkerTrigger_('continueUpdateFileCounts');
//   SpreadsheetApp.getUi().alert('✅ Progress & trigger direset.');
// }

// function debugCheckSourceData() {
//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  
//   if (!sheet) {
//     SpreadsheetApp.getUi().alert(`❌ Sheet "${SOURCE_SHEET_NAME}" tidak ditemukan`);
//     return;
//   }

//   const lastRow = sheet.getLastRow();
//   const sampleData = sheet.getRange(1, 1, Math.min(11, lastRow), 15).getValues();
//   const regex = /^(\d{4})\s*-\s*Week\s*(\d+)$/i;

//   let msg = `📊 DATA CHECK\n\nTotal Rows: ${lastRow}\n\n`;
//   msg += `=== SAMPLE (Rows 2-6) ===\n`;

//   for (let i = 1; i < Math.min(6, sampleData.length); i++) {
//     const center = sampleData[i][0];
//     const saName = sampleData[i][1];
//     const classId = sampleData[i][2];
//     const week = (sampleData[i][11] || '').toString().trim().replace(/\s+/g, ' ');
//     const match = week.match(regex);
    
//     const saNameClassCode = `${saName} - ${classId}`;
    
//     msg += `Row ${i+1}: ${center} | ${saNameClassCode} | "${week}"\n`;
//     msg += `  Match: ${match ? `Year=${match[1]}, Week=${match[2]}` : 'NO'}\n`;
//   }

//   SpreadsheetApp.getUi().alert(msg);
// }



// /*******************************
//  * FOLDER & FILE UTILITIES
//  *******************************/
// function createNestedFolderStructure(parentFolder, centerCode, year, week, saNameClassCode) {
//   const centerFolderName = CENTER_FOLDER_MAP[centerCode] || centerCode;
//   const centerFolder = getOrCreateFolder(parentFolder, centerFolderName);
//   const weekFolderName = `${year}-Week ${week}`;
//   const weekFolder = getOrCreateFolder(centerFolder, weekFolderName);
//   const saFolder = getOrCreateFolder(weekFolder, saNameClassCode, true);
//   const groupChatFolder = getOrCreateFolder(saFolder.folder, 'Group Chat');
//   const groupFilesFolder = getOrCreateFolder(groupChatFolder, 'Files');
//   const personalChatFolder = getOrCreateFolder(saFolder.folder, 'Personal Chat');
//   const personalFilesFolder = getOrCreateFolder(personalChatFolder, 'Files');

//   return {
//     folder: saFolder.folder,
//     created: saFolder.created
//   };
// }

// function getOrCreateFolder(parentFolder, folderName, returnStatus = false) {
//   const existing = parentFolder.getFoldersByName(folderName);
//   if (existing.hasNext()) {
//     const folder = existing.next();
//     return returnStatus ? { folder, created: false } : folder;
//   }
//   const newFolder = parentFolder.createFolder(folderName);
//   return returnStatus ? { folder: newFolder, created: true } : newFolder;
// }

// function getFilesFromFolder(folder) {
//   const files = folder.getFiles();
//   const items = [];
//   while (files.hasNext()) {
//     const file = files.next();
//     items.push(`${file.getName()} (${file.getUrl()})`);
//   }
//   return items.length ? items.join(', ') : 'Tidak ada file';
// }



// /*******************************
//  * LIST ALL STRUCTURE
//  *******************************/
// function startListAllStructure() {
//   const props = PropertiesService.getScriptProperties();
//   props.deleteProperty('LIST_STATE');
//   props.deleteProperty('LIST_DONE');

//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   let sheet = ss.getSheetByName(DEST_SHEET_NAME);
//   if (!sheet) sheet = ss.insertSheet(DEST_SHEET_NAME);
//   sheet.clear();
//   setupSheetHeaders(sheet);

//   const msg = continueListAllStructure();

//   if (!msg.startsWith('✅ List All Structure SELESAI')) {
//     createWorkerTrigger_('continueListAllStructure');
//   } else {
//     deleteWorkerTrigger_('continueListAllStructure');
//   }

//   SpreadsheetApp.getUi().alert(msg);
// }

// function continueListAllStructure() {
//   const startTime = new Date().getTime();
//   const props = PropertiesService.getScriptProperties();

//   let state = props.getProperty('LIST_STATE');
//   if (state) {
//     state = JSON.parse(state);
//   } else {
//     state = {
//       centerIndex: 0,
//       weekIndex: 0,
//       saIndex: 0,
//       rowWritten: 1,
//       totalFolders: 0
//     };
//   }

//   const npsFolder = DriveApp.getFolderById(NPS_ROOT_ID);
//   const centerFolders = npsFolder.getFolders();
//   const centers = [];
//   while (centerFolders.hasNext()) centers.push(centerFolders.next());

//   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//   const sheet = ss.getSheetByName(DEST_SHEET_NAME);

//   let writtenThisRun = 0;

//   for (let c = state.centerIndex; c < centers.length; c++) {
//     const centerFolder = centers[c];
//     const centerName = centerFolder.getName();

//     const weekFoldersIter = centerFolder.getFolders();
//     const weeks = [];
//     while (weekFoldersIter.hasNext()) weeks.push(weekFoldersIter.next());

//     for (let w = (c === state.centerIndex ? state.weekIndex : 0); w < weeks.length; w++) {
//       const weekFolder = weeks[w];
//       const weekName = weekFolder.getName();

//       const saFoldersIter = weekFolder.getFolders();
//       const saFolders = [];
//       while (saFoldersIter.hasNext()) saFolders.push(saFoldersIter.next());

//       for (let s = (c === state.centerIndex && w === state.weekIndex ? state.saIndex : 0); s < saFolders.length; s++) {
//         const saFolder = saFolders[s];

//         if (new Date().getTime() - startTime > MAX_EXECUTION_TIME ||
//             writtenThisRun >= LIST_BATCH_LIMIT) {

//           state.centerIndex = c;
//           state.weekIndex = w;
//           state.saIndex = s;
//           props.setProperty('LIST_STATE', JSON.stringify(state));

//           return `⏱️ List All Structure (PAUSE)\n` +
//                  `Center index: ${c + 1}/${centers.length}\n` +
//                  `Total folder tercatat: ${state.totalFolders}\n\n` +
//                  `➡️ Worker trigger akan melanjutkan otomatis.`;
//         }

//         const url = saFolder.getUrl();
//         const personalChatData = getFilesDataFromSubfolder(saFolder, 'Personal Chat');
//         const groupChatData = getFilesDataFromSubfolder(saFolder, 'Group Chat');

//         state.rowWritten++;
//         state.totalFolders++;
//         writtenThisRun++;

//         sheet.getRange(state.rowWritten, 1, 1, 8).setValues([[
//           centerName,
//           weekName,
//           saFolder.getName(),
//           url,
//           personalChatData.count,
//           personalChatData.links,
//           groupChatData.count,
//           groupChatData.links
//         ]]);
//       }

//       state.saIndex = 0;
//     }

//     state.weekIndex = 0;
//   }

//   props.deleteProperty('LIST_STATE');
//   props.setProperty('LIST_DONE', 'true');
//   sheet.autoResizeColumns(1, 8);
//   deleteWorkerTrigger_('continueListAllStructure');

//   return `✅ List All Structure SELESAI\nTotal folder tercatat: ${state.totalFolders}`;
// }



// /*******************************
//  * TIME-BASED TRIGGERS HELPERS
//  *******************************/
// function createWorkerTrigger_(handlerFunctionName) {
//   const triggers = ScriptApp.getProjectTriggers();
//   triggers.forEach(t => {
//     if (t.getHandlerFunction() === handlerFunctionName) {
//       ScriptApp.deleteTrigger(t);
//     }
//   });

//   ScriptApp.newTrigger(handlerFunctionName)
//     .timeBased()
//     .everyMinutes(1)
//     .create();
// }

// function deleteWorkerTrigger_(handlerFunctionName) {
//   const triggers = ScriptApp.getProjectTriggers();
//   triggers.forEach(t => {
//     if (t.getHandlerFunction() === handlerFunctionName) {
//       ScriptApp.deleteTrigger(t);
//     }
//   });
// }
