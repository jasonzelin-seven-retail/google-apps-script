function updateComplyPCAndGroup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define worksheets
  const sourceSheet = ss.getSheetByName("Folder Structure");
  const targetSheet = ss.getSheetByName("Schedule (Update)");
  
  if (!sourceSheet || !targetSheet) {
    Logger.log("Error: Worksheet tidak ditemukan!");
    showMessage("Error: Worksheet tidak ditemukan!");
    return;
  }
  
  // Get source data (Folder Structure)
  const sourceData = sourceSheet.getDataRange().getValues();
  
  // Get target data (Schedule Update) - struktur kolom baru
  const targetData = targetSheet.getDataRange().getValues();
  
  // Build source lookup map for faster matching
  const sourceMap = {};
  
  for (let i = 1; i < sourceData.length; i++) {
    const center = sourceData[i][0]; // Column A - Center
    const weekFolder = sourceData[i][1]; // Column B - Week Folder (2026-Week 5)
    const saNameClassCode = sourceData[i][2]; // Column C - SA Name - Class Code
    const personalChatFiles = sourceData[i][4]; // Column E - Personal Chat Files
    const groupChatFiles = sourceData[i][6]; // Column G - Group Chat Files
    
    // Parse SA Name and Class Code
    if (!saNameClassCode || saNameClassCode.toString().trim() === "") continue;
    
    const parts = saNameClassCode.toString().split(" - ");
    if (parts.length < 2) continue;
    
    const saName = parts[0].trim();
    const classCode = parts[1].trim();
    
    // Normalize week format: "2026-Week 5" → "2026 - Week 5"
    const normalizedWeek = weekFolder.toString().replace("2026-Week", "2026 - Week");
    
    // Count files dengan handling berbagai format
    const complyPC = countFiles(personalChatFiles);
    const complyGroup = countFiles(groupChatFiles);
    
    // Debug log untuk kasus spesifik seperti Niluh - C25252
    if (classCode === "C25252") {
      Logger.log(`DEBUG - ${saName} - ${classCode}`);
      Logger.log(`Personal Chat Files: ${personalChatFiles}`);
      Logger.log(`Comply PC: ${complyPC}`);
      Logger.log(`Group Chat Files: ${groupChatFiles}`);
      Logger.log(`Comply Group: ${complyGroup}`);
    }
    
    // Create unique key for matching: Center|SA|ClassID|Week
    const key = `${center}|${saName}|${classCode}|${normalizedWeek}`;
    
    sourceMap[key] = {
      complyPC: complyPC,
      complyGroup: complyGroup
    };
  }
  
  // Update target sheet dengan struktur kolom baru
  const updates = [];
  
  for (let i = 1; i < targetData.length; i++) {
    const center = targetData[i][0];        // Kolom A (index 0)
    const studentAdvisor = targetData[i][1]; // Kolom B (index 1)
    const classID = targetData[i][2];        // Kolom C (index 2)
    const weekBlasting = targetData[i][10];  // Kolom K (index 10) - Week Blasting
    
    // Skip jika data kosong
    if (!center || !studentAdvisor || !classID || !weekBlasting) continue;
    
    // Create lookup key
    const key = `${center}|${studentAdvisor}|${classID}|${weekBlasting}`;
    
    // Find match in source
    if (sourceMap[key]) {
      updates.push({
        row: i + 1,
        complyPC: sourceMap[key].complyPC,
        complyGroup: sourceMap[key].complyGroup
      });
    }
  }
  
  // Batch update ke kolom H dan I
  if (updates.length > 0) {
    updates.forEach(update => {
      targetSheet.getRange(update.row, 8).setValue(update.complyPC);   // Kolom H - #Comply PC
      targetSheet.getRange(update.row, 9).setValue(update.complyGroup); // Kolom I - #Comply Group
    });
    
    const message = `✅ Berhasil update ${updates.length} baris di kolom H & I!`;
    Logger.log(message);
    showMessage(message);
  } else {
    const message = "⚠️ Tidak ada data yang match. Periksa:\n1. Format Week Blasting (harus '2026 - Week 5')\n2. Nama SA & Class ID persis sama\n3. Center code cocok";
    Logger.log(message);
    showMessage(message);
  }
}

/**
 * Count files dari cell value dengan robust handling
 * @param {string|number} value - Nilai sel berisi info file
 * @return {number} Jumlah file
 */
function countFiles(value) {
  // Handle empty/null
  if (!value || value === "") return 0;
  
  // Jika sudah number, return langsung
  if (typeof value === "number") return value;
  
  const strValue = value.toString().trim();
  
  // Handle "Tidak ada file" (case insensitive)
  if (strValue.toLowerCase().includes("tidak ada file")) return 0;
  
  // Jika string berupa angka murni
  if (!isNaN(strValue) && strValue !== "") return parseInt(strValue);
  
  // Split by newline dan filter baris kosong/"tidak ada"
  const files = strValue.split("\n").filter(line => 
    line.trim() !== "" && 
    !line.toLowerCase().includes("tidak ada")
  );
  return files.length;
}

/**
 * Tampilkan pesan dengan fallback untuk non-UI context
 */
function showMessage(message) {
  try {
    SpreadsheetApp.getUi().alert("Update Comply PC/Group", message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Update Status", 10);
    Logger.log(message);
  }
}
