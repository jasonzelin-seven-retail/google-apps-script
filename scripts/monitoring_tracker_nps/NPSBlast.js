// const CONFIG = {
//   SOURCE_SHEET:  "Schedule",
//   START_ROW:     7701,
//   END_ROW:       999999,

//   // Urutan kolom sesuai database baru
//   COL_CENTER:       1,   // A - Center
//   COL_CLASS_ID:     3,   // C - Class ID
//   COL_CLASS_NAME:   4,   // D - Class
//   COL_GRAD_DATE:    6,   // F - Graduation Date
//   COL_WEEK_BLAST:  11,   // K - Week Blasting
//   COL_DATE_BLAST:  22,   // V - Date Blasting

//   SHEET_WEEKLY:    "Analisis_Mingguan",
//   SHEET_PER_CLASS: "Analisis_Per_Kelas",
//   SHEET_PIVOT:     "Pivot_Center_Week",
// };

// // ========= UTIL BACA DATA =========
// function loadData_() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sh = ss.getSheetByName(CONFIG.SOURCE_SHEET);
//   if (!sh) throw new Error(`Sheet "${CONFIG.SOURCE_SHEET}" tidak ditemukan`);

//   const lastRow = Math.min(CONFIG.END_ROW, sh.getLastRow());
//   const start   = CONFIG.START_ROW;
//   const numRows = lastRow - start + 1;
//   const numCols = Math.max(CONFIG.COL_DATE_BLAST, CONFIG.COL_WEEK_BLAST);
//   if (numRows <= 0) throw new Error("Tidak ada data pada range yang dipilih");

//   const values = sh.getRange(start, 1, numRows, numCols).getValues();

//   return values
//     .map(r => {
//       const center    = r[CONFIG.COL_CENTER     - 1];
//       const classId   = r[CONFIG.COL_CLASS_ID   - 1];
//       const className = r[CONFIG.COL_CLASS_NAME - 1];
//       const gradDate  = r[CONFIG.COL_GRAD_DATE  - 1];
//       const weekBlast = r[CONFIG.COL_WEEK_BLAST - 1];
//       const dateBlast = r[CONFIG.COL_DATE_BLAST - 1];

//       let blastFlag = "";
//       if (dateBlast && gradDate) {
//         blastFlag = new Date(dateBlast) > new Date(gradDate) ? "EXCEED" : "OK";
//       }

//       return { center, classId, className, gradDate, weekBlast, dateBlast, blastFlag };
//     })
//     .filter(r => r.center || r.classId);
// }

// // ========= 1️⃣ ANALISIS MINGGUAN =========
// function runWeeklyAnalysisOnly() {
//   const ss   = SpreadsheetApp.getActiveSpreadsheet();
//   const data = loadData_();
//   const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_WEEKLY);
//   sheet.clearContents().clearFormats();

//   const centers = uniqueSorted_(data.map(d => d.center));
//   const weekMap = {};

//   data.forEach(({ center, classId, weekBlast, blastFlag }) => {
//     if (!weekBlast) return;
//     const w = String(weekBlast);
//     const c = String(center);

//     if (!weekMap[w]) weekMap[w] = {
//       total: 0, classes: new Set(), exceed: 0, byCenter: {}
//     };

//     weekMap[w].total++;
//     weekMap[w].classes.add(`${c}||${classId}`);
//     if (blastFlag === "EXCEED") weekMap[w].exceed++;

//     if (!weekMap[w].byCenter[c]) weekMap[w].byCenter[c] = { total: 0, classes: new Set(), exceed: 0 };
//     weekMap[w].byCenter[c].total++;
//     weekMap[w].byCenter[c].classes.add(String(classId));
//     if (blastFlag === "EXCEED") weekMap[w].byCenter[c].exceed++;
//   });

//   const headers = ["Week", "Total Entri", "Kelas Unik", "⚠️ Post-Grad"];
//   centers.forEach(c => {
//     headers.push(`${c} — Entri`, `${c} — Kelas`, `${c} — Post-Grad`);
//   });

//   const rows = [headers];
//   const weeks = sortWeeks_(Object.keys(weekMap));
//   const exceedRows = [];

//   weeks.forEach((w, idx) => {
//     const obj = weekMap[w];
//     const row = [w, obj.total, obj.classes.size, obj.exceed];
//     centers.forEach(c => {
//       const cd = obj.byCenter[c];
//       row.push(cd ? cd.total : 0);
//       row.push(cd ? cd.classes.size : 0);
//       row.push(cd ? cd.exceed : 0);
//     });
//     rows.push(row);
//     if (obj.exceed > 0) exceedRows.push(idx + 2);
//   });

//   const r = sheet.getRange(1, 1, rows.length, rows[0].length);
//   r.setValues(rows);

//   const bg = Array.from({ length: rows.length }, () => Array(rows[0].length).fill(null));
//   const fw = Array.from({ length: rows.length }, () => Array(rows[0].length).fill("normal"));
//   const fc = Array.from({ length: rows.length }, () => Array(rows[0].length).fill("#000000"));

//   for (let c = 0; c < rows[0].length; c++) {
//     bg[0][c] = "#1a73e8"; fw[0][c] = "bold"; fc[0][c] = "#ffffff";
//   }
//   exceedRows.forEach(rowIdx => {
//     bg[rowIdx - 1][3] = "#f4cccc"; fw[rowIdx - 1][3] = "bold"; fc[rowIdx - 1][3] = "#c0392b";
//   });

//   r.setBackgrounds(bg).setFontWeights(fw).setFontColors(fc);
//   sheet.setFrozenRows(1);
// }

// // ========= 2️⃣ ANALISIS PER KELAS =========
// function runPerClassAnalysisOnly() {
//   const ss   = SpreadsheetApp.getActiveSpreadsheet();
//   const data = loadData_();
//   const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_PER_CLASS);
//   sheet.clearContents().clearFormats();

//   const classMap = {};
//   data.forEach(({ center, classId, className, gradDate, weekBlast, dateBlast, blastFlag }) => {
//     if (!classId) return;
//     const key = `${center}||${classId}`;
//     if (!classMap[key]) classMap[key] = {
//       center, classId, className, gradDate,
//       total: 0, before: 0, after: 0, weeks: new Set()
//     };

//     if (dateBlast) {
//       classMap[key].total++;
//       if (blastFlag === "EXCEED") classMap[key].after++;
//       else classMap[key].before++;
//     }
//     if (weekBlast) classMap[key].weeks.add(String(weekBlast));
//   });

//   const headers = [
//     "Center","Class ID","Class Name","Graduation Date",
//     "Total Blast","Sebelum Grad","Setelah Grad ⚠️",
//     "Jumlah Week","Status"
//   ];
//   const rows = [headers];
//   const probRows = [];

//   Object.values(classMap)
//     .sort((a,b) =>
//       String(a.center).localeCompare(String(b.center)) ||
//       String(a.classId).localeCompare(String(b.classId)))
//     .forEach((c, idx) => {
//       const gradStr = c.gradDate
//         ? Utilities.formatDate(new Date(c.gradDate), Session.getScriptTimeZone(), "dd-MMM-yyyy")
//         : "";
//       rows.push([
//         c.center, c.classId, c.className, gradStr,
//         c.total, c.before, c.after,
//         c.weeks.size,
//         c.after > 0 ? "⚠️ ADA BLAST POST-GRAD" : "✅ Aman"
//       ]);
//       if (c.after > 0) probRows.push(idx + 2);
//     });

//   const r = sheet.getRange(1,1,rows.length,rows[0].length);
//   r.setValues(rows);

//   const bg = Array.from({ length: rows.length }, () => Array(rows[0].length).fill(null));
//   const fw = Array.from({ length: rows.length }, () => Array(rows[0].length).fill("normal"));
//   const fc = Array.from({ length: rows.length }, () => Array(rows[0].length).fill("#000000"));

//   for (let c = 0; c < rows[0].length; c++) {
//     bg[0][c] = "#1a73e8"; fw[0][c] = "bold"; fc[0][c] = "#ffffff";
//   }
//   probRows.forEach(rIdx => {
//     for (let c = 0; c < rows[0].length; c++) bg[rIdx-1][c] = "#fce8e6";
//     fw[rIdx-1][8] = "bold"; fc[rIdx-1][8] = "#c0392b";
//   });

//   r.setBackgrounds(bg).setFontWeights(fw).setFontColors(fc);
//   sheet.setFrozenRows(1);
// }

// // ========= 3️⃣ PIVOT CENTER X WEEK =========
// function runPivotCenterWeekOnly() {
//   const ss   = SpreadsheetApp.getActiveSpreadsheet();
//   const data = loadData_();
//   const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_PIVOT);
//   sheet.clearContents().clearFormats();

//   const weeks   = sortWeeks_(uniqueSorted_(data.map(d => d.weekBlast).filter(Boolean)));
//   const centers = uniqueSorted_(data.map(d => d.center).filter(Boolean));

//   const pivot = {};
//   centers.forEach(c => {
//     pivot[c] = {};
//     weeks.forEach(w => pivot[c][w] = { classes: new Set(), exceed: 0 });
//   });

//   data.forEach(({ center, classId, weekBlast, blastFlag }) => {
//     if (!center || !weekBlast) return;
//     const c = String(center), w = String(weekBlast);
//     if (!pivot[c] || !pivot[c][w]) return;
//     pivot[c][w].classes.add(String(classId));
//     if (blastFlag === "EXCEED") pivot[c][w].exceed++;
//   });

//   let rowIndex = 1;
//   sheet.getRange(rowIndex,1).setValue("📊 KELAS UNIK PER CENTER PER WEEK")
//        .setFontWeight("bold").setFontSize(12).setFontColor("#1a73e8");
//   rowIndex++;

//   const header1 = ["Center / Week", ...weeks];
//   const rows1 = [header1];
//   centers.forEach(c => {
//     rows1.push([c, ...weeks.map(w => pivot[c][w].classes.size)]);
//   });

//   const r1 = sheet.getRange(rowIndex,1,rows1.length,rows1[0].length);
//   r1.setValues(rows1);
//   const bg1 = Array.from({length:rows1.length},()=>Array(rows1[0].length).fill(null));
//   const fw1 = Array.from({length:rows1.length},()=>Array(rows1[0].length).fill("normal"));
//   const fc1 = Array.from({length:rows1.length},()=>Array(rows1[0].length).fill("#000000"));
//   for (let c=0;c<rows1[0].length;c++){bg1[0][c]="#1a73e8";fw1[0][c]="bold";fc1[0][c]="#ffffff";}
//   r1.setBackgrounds(bg1).setFontWeights(fw1).setFontColors(fc1);
//   rowIndex += rows1.length + 2;

//   sheet.getRange(rowIndex,1).setValue("⚠️ BLAST SETELAH GRADUATION PER CENTER PER WEEK")
//        .setFontWeight("bold").setFontSize(12).setFontColor("#c0392b");
//   rowIndex++;

//   const header2 = ["Center / Week", ...weeks];
//   const rows2 = [header2];
//   centers.forEach(c => {
//     rows2.push([c, ...weeks.map(w => pivot[c][w].exceed)]);
//   });

//   const r2 = sheet.getRange(rowIndex,1,rows2.length,rows2[0].length);
//   r2.setValues(rows2);
//   const bg2 = Array.from({length:rows2.length},()=>Array(rows2[0].length).fill(null));
//   const fw2 = Array.from({length:rows2.length},()=>Array(rows2[0].length).fill("normal"));
//   const fc2 = Array.from({length:rows2.length},()=>Array(rows2[0].length).fill("#000000"));
//   for (let c=0;c<rows2[0].length;c++){bg2[0][c]="#c0392b";fw2[0][c]="bold";fc2[0][c]="#ffffff";}
//   for (let r=1;r<rows2.length;r++){
//     for (let c=1;c<rows2[r].length;c++){
//       if (rows2[r][c]>0){bg2[r][c]="#f4cccc";fw2[r][c]="bold";fc2[r][c]="#c0392b";}
//     }
//   }
//   r2.setBackgrounds(bg2).setFontWeights(fw2).setFontColors(fc2);

//   sheet.setFrozenRows(2);
//   sheet.setFrozenColumns(1);
// }

// // ========= HELPERS =========
// function getOrCreateSheet_(ss, name) {
//   return ss.getSheetByName(name) || ss.insertSheet(name);
// }

// function uniqueSorted_(arr) {
//   return [...new Set(arr.filter(v => v !== "" && v != null))].sort();
// }

// // Sort week berbasis Year + Week Number
// function sortWeeks_(weeks) {
//   return weeks.sort((a, b) => {
//     const pa = String(a).match(/(\d{4}).*?(\d{1,2})/);
//     const pb = String(b).match(/(\d{4}).*?(\d{1,2})/);
//     if (!pa || !pb) return String(a).localeCompare(String(b));
//     const ya = parseInt(pa[1], 10);
//     const wa = parseInt(pa[2], 10);
//     const yb = parseInt(pb[1], 10);
//     const wb = parseInt(pb[2], 10);
//     if (ya !== yb) return ya - yb;
//     return wa - wb;
//   });
// }

// // ========= MENU =========
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("📊 NPS Analyzer")
//     .addItem("1️⃣ Analisis Mingguan saja", "runWeeklyAnalysisOnly")
//     .addItem("2️⃣ Analisis Per Kelas saja", "runPerClassAnalysisOnly")
//     .addItem("3️⃣ Pivot Center x Week saja", "runPivotCenterWeekOnly")
//     .addToUi();
// }
