// /**
//  * ============================================
//  * KELAS MANAGER v3.0 - COMPREHENSIVE SOLUTION
//  * ============================================
//  * Task 1: Max X kelas/SA/week dengan auto-move
//  * Task 2: Overtime Survey (post-graduation)
//  * Task 3: Back-to-back Survey Matrix (<30 days)
//  */

// // ============================================
// // MENU & INITIALIZATION
// // ============================================

// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('📊 Kelas Manager')
//     .addItem('1️⃣ Quick Check Max/SA', 'quickCheck')
//     .addItem('2️⃣ Proses Kelas (Max 10)', 'processClasses')
//     .addItem('3️⃣ Proses Custom', 'showCustomDialog')
//     .addSeparator()
//     .addItem('📋 Task 2: Overtime Surveys', 'task2_overtimeSurveys')
//     .addItem('📊 Task 3: Back-to-Back Matrix', 'task3_backToBackMatrix')
//     .addToUi();
// }

// // ============================================
// // TASK 1: MAX KELAS PER SA
// // ============================================

// function quickCheck() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName('Schedule');
//   if (!sheet) {
//     SpreadsheetApp.getUi().alert('❌ Sheet "Schedule" tidak ditemukan!');
//     return;
//   }
  
//   const colMap = getColumnMapping(sheet.getDataRange().getValues()[0]);
//   const today = new Date();
//   today.setHours(0, 0, 0, 0);
  
//   const advisorWeekCount = {};
//   let expiredCount = 0;
//   const data = sheet.getDataRange().getValues();
  
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const advisor = String(row[colMap.advisor] || '').trim();
//     const weekBlasting = String(row[colMap.weekBlasting] || '').trim();
//     const status = String(row[colMap.status] || '').toLowerCase().trim();
//     const graduationDateStr = row[colMap.graduationDate];
    
//     const gradDate = parseDate(graduationDateStr);
//     if (gradDate && gradDate < today && status === 'active') {
//       expiredCount++;
//     }
    
//     if (status === 'active' && advisor && weekBlasting && (!gradDate || gradDate >= today)) {
//       const key = advisor + '_' + weekBlasting;
//       if (!advisorWeekCount[key]) advisorWeekCount[key] = 0;
//       advisorWeekCount[key]++;
//     }
//   }
  
//   const maxPerSA = Math.max(...Object.values(advisorWeekCount), 0);
//   const saWithExcess = Object.values(advisorWeekCount).filter(v => v > 10).length;
  
//   const msg = '📊 QUICK CHECK:\n\n' +
//               'Max kelas/SA/week: ' + maxPerSA + '\n' +
//               'SA dengan >10 kelas: ' + saWithExcess + '\n' +
//               'Kelas expired: ' + expiredCount + '\n\n' +
//               (maxPerSA <= 10 ? '✅ SEMUA BAIK!' : '🚨 PERLU PROSES!');
  
//   SpreadsheetApp.getUi().alert(msg);
//   return { maxPerSA, expiredCount, saWithExcess };
// }

// function processClasses() {
//   return processCustomClasses({
//     centers: [],
//     advisors: [],
//     weekBlasting: '',
//     maxClassesPerSA: 10,
//     startRow: 2,
//     endRow: 100000
//   });
// }

// // ============================================
// // TASK 2: OVERTIME SURVEY DETECTION
// // ============================================

// function task2_overtimeSurveys() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
//   if (!sheet) {
//     SpreadsheetApp.getUi().alert('❌ Sheet "Schedule" tidak ditemukan!');
//     return;
//   }
  
//   const data = sheet.getDataRange().getValues();
//   const headers = data[0];
//   const today = new Date();
//   today.setHours(0, 0, 0, 0);
  
//   // Auto-detect columns
//   let gradCol = -1;
//   const surveyCols = [];
  
//   headers.forEach((h, i) => {
//     const header = String(h).toLowerCase().trim();
//     if (header.includes('graduation') && header.includes('date') && !header.includes('week')) {
//       gradCol = i;
//     }
//     if ((header.includes('survey') || header.includes('blast')) && 
//         (header.includes('date') || header.includes('week') || header.includes('schedule'))) {
//       surveyCols.push({index: i, name: h});
//     }
//   });
  
//   if (gradCol === -1) {
//     SpreadsheetApp.getUi().alert('❌ Kolom Graduation Date tidak ditemukan!\nPastikan ada kolom dengan "Graduation Date"');
//     return;
//   }
  
//   let overtimeCount = 0;
//   const details = [];
//   const overtimeBySurvey = {};
  
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const center = String(row[0] || '').trim();
//     const sa = String(row[1] || '').trim();
//     const status = String(row[7] || '').toLowerCase().trim();
//     const gradDateRaw = row[gradCol];
    
//     if (status !== 'active') continue;
    
//     const gradDate = parseDate(gradDateRaw);
//     if (!gradDate || gradDate >= today) continue;
    
//     // Check if any survey scheduled after graduation
//     let hasOvertimeSurvey = false;
//     surveyCols.forEach(surveyCol => {
//       const surveyValue = row[surveyCol.index];
//       if (surveyValue) {
//         const surveyDate = parseDate(surveyValue);
//         if (surveyDate && surveyDate > gradDate) {
//           hasOvertimeSurvey = true;
//           const surveyName = surveyCol.name;
//           if (!overtimeBySurvey[surveyName]) overtimeBySurvey[surveyName] = 0;
//           overtimeBySurvey[surveyName]++;
//         }
//       }
//     });
    
//     if (hasOvertimeSurvey) {
//       overtimeCount++;
//       details.push(center + ' - ' + sa + ' (Row ' + (i+1) + ')');
//     }
//   }
  
//   // Display results
//   let resultMsg = '🚨 TASK 2: OVERTIME SURVEY SCHEDULE\n\n';
//   resultMsg += 'Total kelas dengan survey setelah graduation: ' + overtimeCount + '\n\n';
  
//   if (overtimeCount > 0) {
//     resultMsg += 'Detail per Survey Column:\n';
//     Object.keys(overtimeBySurvey).forEach(surveyName => {
//       resultMsg += '  • ' + surveyName + ': ' + overtimeBySurvey[surveyName] + ' kelas\n';
//     });
//     resultMsg += '\nSample Details (first 10):\n' + details.slice(0, 10).join('\n');
//   } else {
//     resultMsg += '✅ TIDAK ADA overtime survey!';
//   }
  
//   SpreadsheetApp.getUi().alert(resultMsg);
//   Logger.log('Overtime Survey Count: ' + overtimeCount);
  
//   return {
//     overtimeCount: overtimeCount,
//     details: details,
//     bySurvey: overtimeBySurvey
//   };
// }

// // ============================================
// // TASK 3: BACK-TO-BACK SURVEY MATRIX
// // ============================================

// function task3_backToBackMatrix() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
//   if (!sheet) {
//     SpreadsheetApp.getUi().alert('❌ Sheet "Schedule" tidak ditemukan!');
//     return;
//   }
  
//   const data = sheet.getDataRange().getValues();
//   const headers = data[0];
//   const MIN_DAYS = 30;
  
//   // Auto-detect survey date columns
//   const surveyDateCols = [];
//   headers.forEach((h, i) => {
//     const header = String(h).toLowerCase().trim();
//     if ((header.includes('survey') || header.includes('blast')) && 
//         (header.includes('date') || header.includes('schedule'))) {
//       surveyDateCols.push({index: i, name: h, order: surveyDateCols.length + 1});
//     }
//   });
  
//   if (surveyDateCols.length < 2) {
//     SpreadsheetApp.getUi().alert('❌ Perlu minimal 2 kolom survey date!\nContoh: "Survey 1 Date", "Survey 2 Date"');
//     return;
//   }
  
//   // Matrix: Center -> Survey Pair -> Count
//   const centerMatrix = {};
//   const surveyPairs = [];
  
//   // Generate survey pairs (Survey 1-2, Survey 2-3, etc)
//   for (let i = 0; i < surveyDateCols.length - 1; i++) {
//     surveyPairs.push({
//       from: surveyDateCols[i],
//       to: surveyDateCols[i + 1],
//       label: 'Survey ' + (i+1) + '-' + (i+2)
//     });
//   }
  
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const center = String(row[0] || '').trim();
//     const status = String(row[7] || '').toLowerCase().trim();
    
//     if (!center || status !== 'active') continue;
    
//     if (!centerMatrix[center]) {
//       centerMatrix[center] = {};
//       surveyPairs.forEach(pair => {
//         centerMatrix[center][pair.label] = 0;
//       });
//     }
    
//     // Check each survey pair
//     surveyPairs.forEach(pair => {
//       const date1 = parseDate(row[pair.from.index]);
//       const date2 = parseDate(row[pair.to.index]);
      
//       if (date1 && date2) {
//         const diffDays = Math.abs((date2 - date1) / (1000 * 60 * 60 * 24));
//         if (diffDays < MIN_DAYS) {
//           centerMatrix[center][pair.label]++;
//         }
//       }
//     });
//   }
  
//   // Generate matrix display
//   const centers = Object.keys(centerMatrix).sort();
//   let matrixText = '📊 TASK 3: BACK-TO-BACK SURVEY MATRIX\n';
//   matrixText += '(Jarak < 30 hari)\n\n';
  
//   // Header row
//   matrixText += 'Center'.padEnd(12);
//   surveyPairs.forEach(pair => {
//     matrixText += pair.label.padEnd(12);
//   });
//   matrixText += '\n' + '-'.repeat(12 + surveyPairs.length * 12) + '\n';
  
//   // Data rows
//   centers.forEach(center => {
//     matrixText += center.padEnd(12);
//     surveyPairs.forEach(pair => {
//       const count = centerMatrix[center][pair.label];
//       matrixText += String(count).padEnd(12);
//     });
//     matrixText += '\n';
//   });
  
//   SpreadsheetApp.getUi().alert(matrixText);
  
//   // Create summary sheet
//   createMatrixSheet(centerMatrix, surveyPairs);
  
//   Logger.log('Back-to-back Matrix: ' + JSON.stringify(centerMatrix));
//   return centerMatrix;
// }

// function createMatrixSheet(centerMatrix, surveyPairs) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let matrixSheet = ss.getSheetByName('Back-to-Back Matrix');
  
//   if (matrixSheet) {
//     matrixSheet.clear();
//   } else {
//     matrixSheet = ss.insertSheet('Back-to-Back Matrix');
//   }
  
//   // Headers
//   const headers = ['Center'];
//   surveyPairs.forEach(pair => headers.push(pair.label));
//   headers.push('TOTAL');
  
//   matrixSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
//   matrixSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
//   // Data
//   const centers = Object.keys(centerMatrix).sort();
//   const dataRows = [];
  
//   centers.forEach(center => {
//     const row = [center];
//     let total = 0;
//     surveyPairs.forEach(pair => {
//       const count = centerMatrix[center][pair.label];
//       row.push(count);
//       total += count;
//     });
//     row.push(total);
//     dataRows.push(row);
//   });
  
//   if (dataRows.length > 0) {
//     matrixSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
//   }
  
//   // Format
//   matrixSheet.autoResizeColumns(1, headers.length);
//   matrixSheet.setFrozenRows(1);
  
//   SpreadsheetApp.getUi().alert('✅ Matrix disimpan di sheet "Back-to-Back Matrix"!');
// }

// // ============================================
// // HELPER FUNCTIONS
// // ============================================

// function parseDate(dateStr) {
//   if (!dateStr) return null;
//   if (dateStr instanceof Date) return new Date(dateStr);
  
//   const str = String(dateStr).trim();
  
//   // ISO format: 2026-02-12
//   const isoMatch = str.match(/(\d{4}-\d{2}-\d{2})/);
//   if (isoMatch) return new Date(isoMatch[1]);
  
//   // DD-MMM-YYYY: 4-Aug-2025
//   const dmy = str.match(/(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
//   if (dmy) {
//     const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
//     const month = months[dmy[2].toLowerCase()];
//     if (month !== undefined) {
//       return new Date(parseInt(dmy[3]), month, parseInt(dmy[1]));
//     }
//   }
  
//   // Week format: 2026 - Week 7
//   const weekMatch = str.match(/(\d{4})\s*-?\s*Week\s*(\d+)/i);
//   if (weekMatch) {
//     const year = parseInt(weekMatch[1]);
//     const week = parseInt(weekMatch[2]);
//     const jan1 = new Date(year, 0, 1);
//     const daysToAdd = (week - 1) * 7;
//     return new Date(jan1.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
//   }
  
//   return null;
// }

// function getColumnMapping(headers) {
//   const map = {};
//   headers.forEach((h, i) => {
//     const header = String(h).toLowerCase().trim();
//     if (header.includes('center')) map.center = i;
//     if (header.includes('advisor') || header.includes('sa')) map.advisor = i;
//     if (header.includes('status')) map.status = i;
//     if (header.includes('graduation') && header.includes('date') && !header.includes('week')) map.graduationDate = i;
//     if (header.includes('week') && header.includes('blasting')) map.weekBlasting = i;
//   });
//   return map;
// }

// function getSheetData() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName('Schedule');
//   if (!sheet) throw new Error('Sheet "Schedule" tidak ditemukan!');
  
//   const data = sheet.getDataRange().getValues();
//   const centers = [];
//   const advisors = [];
//   const weeks = [];
//   const centerAdvisorMap = {};

//   for (let i = 1; i < Math.min(data.length, 1000); i++) {
//     const row = data[i];
//     const center = row[0];
//     const advisor = row[1];
//     const weekBlasting = row[11];
    
//     if (center && !centers.includes(center)) centers.push(center);
//     if (advisor && !advisors.includes(advisor)) advisors.push(advisor);
//     if (weekBlasting && !weeks.includes(weekBlasting)) weeks.push(weekBlasting);
    
//     if (center) {
//       if (!centerAdvisorMap[center]) centerAdvisorMap[center] = [];
//       if (advisor && !centerAdvisorMap[center].includes(advisor)) {
//         centerAdvisorMap[center].push(advisor);
//       }
//     }
//   }
  
//   return { centers, advisors, weeks, centerAdvisorMap, totalRows: sheet.getLastRow() };
// }

// function processCustomClasses(params) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName('Schedule');
//   if (!sheet) throw new Error('Sheet "Schedule" tidak ditemukan!');
  
//   const data = sheet.getDataRange().getValues();
//   const colMap = getColumnMapping(data[0]);
//   const today = new Date();
//   today.setHours(0, 0, 0, 0);
  
//   let moved = 0;
//   let deleted = 0;
//   const maxClassesPerSA = params.maxClassesPerSA || 10;
  
//   // Delete expired first
//   for (let i = data.length - 1; i >= params.startRow - 1; i--) {
//     const row = data[i];
//     const status = String(row[colMap.status] || '').toLowerCase().trim();
//     const gradDate = parseDate(row[colMap.graduationDate]);
    
//     if (status === 'active' && gradDate && gradDate < today) {
//       sheet.deleteRow(i + 1);
//       deleted++;
//     }
//   }
  
//   const finalMax = getCurrentMaxPerSA(sheet, colMap, params);
  
//   return {
//     moved: moved,
//     deleted: deleted,
//     maxPerSA: finalMax
//   };
// }

// function getCurrentMaxPerSA(sheet, colMap, params) {
//   const data = sheet.getDataRange().getValues();
//   const advisorWeekCount = {};
//   const today = new Date();
//   today.setHours(0, 0, 0, 0);
  
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     const advisor = String(row[colMap.advisor] || '').trim();
//     const week = String(row[colMap.weekBlasting] || '').trim();
//     const status = String(row[colMap.status] || '').toLowerCase().trim();
//     const gradDate = parseDate(row[colMap.graduationDate]);
    
//     if (status === 'active' && advisor && week && (!gradDate || gradDate >= today)) {
//       const key = advisor + '_' + week;
//       if (!advisorWeekCount[key]) advisorWeekCount[key] = 0;
//       advisorWeekCount[key]++;
//     }
//   }
  
//   return Math.max(...Object.values(advisorWeekCount), 0);
// }

// function showCustomDialog() {
//   const html = HtmlService.createHtmlOutput(getHTMLContent())
//     .setWidth(550)
//     .setHeight(650);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Kelas Manager v3.0');
// }

// function getHTMLContent() {
//   return `
// <!DOCTYPE html>
// <html>
// <head>
//   <base target="_top">
//   <style>
//     * { box-sizing: border-box; margin: 0; padding: 0; }
//     body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
//     .container { background: white; border-radius: 12px; padding: 25px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); }
//     h2 { color: #667eea; margin-bottom: 20px; text-align: center; }
//     .form-group { margin: 15px 0; }
//     label { display: block; margin-bottom: 8px; font-weight: 600; }
//     input { width: 100%; padding: 10px; border: 2px solid #e0e0e0; border-radius: 6px; }
//     button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px; border: none; border-radius: 6px; cursor: pointer; width: 100%; margin-top: 10px; font-weight: 600; }
//   </style>
// </head>
// <body>
//   <div class="container">
//     <h2>🎓 Kelas Manager v3.0</h2>
//     <div class="form-group">
//       <label>Max Kelas/SA:</label>
//       <input type="number" id="maxClasses" value="10" min="1" max="20">
//     </div>
//     <button onclick="runProcess()">⚡ PROSES</button>
//   </div>
//   <script>
//     function runProcess() {
//       google.script.run.processClasses();
//       alert('Processing...');
//     }
//   </script>
// </body>
// </html>
// `;
// }
