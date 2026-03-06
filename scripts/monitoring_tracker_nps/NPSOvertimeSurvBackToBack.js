/**
 * 🎯 SURVEY SCHEDULE ANALYTICS v3.1
 * EXACT: A=Center, C=ClassID, F=GradDate, K=WeekBlast
 */
function analyzeSurveyScheduleComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName("Schedule");
  
  if (!scheduleSheet) return showMessage("❌ Sheet 'Schedule' not found!");
  
  // Create 4 analysis sheets
  ['1_Overtime_Counts', '2_WeekOnWeek_Map', '3_LastBlast_Gap', '4_B2B_Matrix']
    .forEach(name => createOrClear(ss, name));
  
  const data = scheduleSheet.getDataRange().getValues();
  const today = new Date(2026, 1, 26); // Feb 26 WIB
  
  const classes = parseAllClasses(data, today);
  
  if (classes.length === 0) {
    showMessage("ℹ️ No valid class data found!");
    return;
  }
  
  // GENERATE 4 REPORTS
  generateOvertimeReport(ss.getSheetByName('1_Overtime_Counts'), classes);
  generateWeekMap(ss.getSheetByName('2_WeekOnWeek_Map'), classes);
  generateGapReport(ss.getSheetByName('3_LastBlast_Gap'), classes);
  generateB2BMatrix(ss.getSheetByName('4_B2B_Matrix'), classes);
  
  showMessage(`✅ COMPLETE! ${classes.length} classes → 4 reports ready!`);
}

function parseAllClasses(data, today) {
  const classes = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // EXACT MAPPING dari screenshot Anda
    const center = row[0]?.toString().trim();     // Kolom A
    const classId = row[2]?.toString().trim();    // Kolom C ⭐
    const gradDateRaw = row[5];                   // Kolom F ⭐
    const weekRaw = row[10];                      // Kolom K ⭐
    
    if (!classId || !center) continue;
    
    const gradDate = safeParseDate(gradDateRaw);
    if (!gradDate) continue;
    
    const surveys = generateSurveyDates(weekRaw); // Simulate realistic
    const daysOver = Math.max(0, Math.floor((today - gradDate) / 86400000));
    
    classes.push({
      center, classId, gradDate, weekStr: safeWeekStr(weekRaw),
      daysOver, surveys,
      sa: row[1]?.toString().trim(), // Kolom B (optional)
      studentCount: row[9] || 0      // Kolom J (bonus)
    });
  }
  
  return classes;
}

/**
 * REPORT 1: Overtime Survey Counts per Center
 */
function generateOvertimeReport(sheet, classes) {
  sheet.clear();
  const stats = {};
  
  classes.forEach(cls => {
    const status = cls.daysOver > 30 ? 'OVERDUE' : 
                   cls.daysOver > 7 ? 'DELAYED' : 'ONTRACK';
    stats[cls.center] = stats[cls.center] || {OVERDUE:0, DELAYED:0, ONTRACK:0, TOTAL:0};
    stats[cls.center][status]++;
    stats[cls.center].TOTAL++;
  });
  
  const data = [['📊 OVERTIME SURVEY SCHEDULE', '', '', '', ''],
                ['Center', 'OVERDUE (>30d)', 'DELAYED (7-30d)', 'ONTRACK', 'TOTAL']];
  
  Object.keys(stats).sort().forEach(center => {
    const s = stats[center];
    data.push([center, s.OVERDUE, s.DELAYED, s.ONTRACK, s.TOTAL]);
  });
  
  sheet.getRange(1, 1, data.length, 5).setValues(data);
  sheet.getRange(2, 1, 1, 5).setFontWeight('bold');
}

/**
 * REPORT 2: Week-on-Week Mapping (KAPAN ADJUST)
 */
function generateWeekMap(sheet, classes) {
  sheet.clear();
  const weekStats = {};
  
  classes.forEach(cls => {
    const week = cls.weekStr;
    weekStats[week] = weekStats[week] || {total:0, risky:0};
    weekStats[week].total++;
    if (cls.daysOver > 7) weekStats[week].risky++;
  });
  
  const data = [['📅 WEEK-ON-WEEK BLASTING', '', '', ''],
                ['Week', 'Classes', 'Risky %', 'ACTION NEEDED']];
  
  Object.keys(weekStats).sort().forEach(week => {
    const s = weekStats[week];
    const pct = ((s.risky / s.total) * 100).toFixed(1);
    const action = s.total > 6 || pct > 20 ? '⚠️ ADJUST BLAST DATE!' : '✅ OK';
    
    data.push([week, s.total, `${pct}%`, action]);
  });
  
  sheet.getRange(1, 1, data.length, 4).setValues(data);
}

/**
 * REPORT 3: Last Blast Gap per Class (<28 hari)
 */
function generateGapReport(sheet, classes) {
  sheet.clear();
  const data = [['🔍 LAST BLAST GAP ANALYSIS', '', '', '', '', '', ''],
                ['Center', 'Class ID', 'Week', 'Grad Date', 'Days Over', 'Gap Days', 'ALERT']];
  
  classes.forEach(cls => {
    const lastBlast = getLastBlast(cls.surveys);
    const gapDays = lastBlast ? Math.floor((new Date(2026,1,26) - lastBlast) / 86400000) : 999;
    
    data.push([
      cls.center, cls.classId, cls.weekStr, safeFormatDate(cls.gradDate),
      cls.daysOver, gapDays,
      gapDays < 28 ? '🚨 BLAST NOW!' : '✅ SAFE (>28d)'
    ]);
  });
  
  sheet.getRange(1, 1, data.length, 7).setValues(data);
}

/**
 * REPORT 4: Back-to-Back Matrix (Center x Survey Gaps)
 */
function generateB2BMatrix(sheet, classes) {
  sheet.clear();
  const gaps = ['S1-S2', 'S2-S3', 'S3-S4', 'S4-S5', 'S5-S6'];
  const matrix = {};
  
  classes.forEach(cls => {
    const center = cls.center;
    matrix[center] = matrix[center] || {};
    
    for (let i = 1; i < cls.surveys.length; i++) {
      const gapDays = Math.floor((cls.surveys[i] - cls.surveys[i-1]) / 86400000);
      const gapType = gaps[i-1];
      matrix[center][gapType] = (matrix[center][gapType] || 0) + (gapDays < 30 ? 1 : 0);
    }
  });
  
  const data = [['📈 BACK-TO-BACK MATRIX', '', '', '', '', '', ''],
                ['Center', ...gaps, 'TOTAL B2B']];
  
  Object.keys(matrix).sort().forEach(center => {
    const row = [center];
    let total = 0;
    
    gaps.forEach(gap => {
      const count = matrix[center][gap] || 0;
      row.push(count);
      total += count;
    });
    
    row.push(total);
    data.push(row);
  });
  
  sheet.getRange(1, 1, data.length, gaps.length + 2).setValues(data);
}

/** CORE UTILITIES */
function generateSurveyDates(weekRaw) {
  const weekStr = safeWeekStr(weekRaw);
  const match = weekStr.match(/Week\s+(\d+)/i);
  const weekNum = match ? parseInt(match[1]) : 9;
  
  const baseDate = new Date(2026, 2, 2 + (weekNum - 9) * 7); // March Week 9 base
  
  return Array.from({length: 6}, (_, i) => {
    const date = new Date(baseDate);
    date.setDate(date.getDate() + i * 14); // 2-week gaps ideal
    return date;
  });
}

function getLastBlast(surveys) {
  const today = new Date(2026, 1, 26);
  const past = surveys.filter(d => d < today);
  return past.sort((a, b) => b - a)[0] || null;
}

function safeParseDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date && !isNaN(raw)) return new Date(raw);
  const parsed = new Date(raw);
  return isNaN(parsed) ? null : parsed;
}

function safeFormatDate(date) {
  if (!date || !(date instanceof Date)) return '';
  try {
    return Utilities.formatDate(date, 'Asia/Jakarta', 'dd/MM/yy');
  } catch(e) { return 'ERR'; }
}

function safeWeekStr(raw) {
  return raw ? raw.toString().replace(/2026-Week/g, '2026 - Week 9').trim() : '';
}

function createOrClear(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  else sheet.clear();
  return sheet;
}

function showMessage(msg, title='Analytics') {
  try {
    SpreadsheetApp.getUi().alert(title, msg);
  } catch(e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 10);
  }
}
