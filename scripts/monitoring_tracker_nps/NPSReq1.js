function createSAVerticalTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName('SA_Vertical') || ss.insertSheet('SA_Vertical');
  target.clear();
  
  const saData = [
    // BSD (10 SA)
    ...Array(10).fill('BSD').map((center, i) => [center, ['Aya','Zahra','Revi','Vio','Mutia','Henni','Niluh','Ferdi','Addien'][i]]),
    // TJD (7 SA)
    ...Array(7).fill('TJD').map((center, i) => [center, ['Azna','Rai','Dinda','Alham','Puji','Tian','Okta'][i]]),
    // BGR (11 SA)
    ...Array(11).fill('BGR').map((center, i) => [center, ['Caca','Aisyah','Nabilla','Gita','Nina','Vita','April','Citra','Regita','caca','Dwie'][i]]),
    // DPK (9 SA)
    ...Array(9).fill('DPK').map((center, i) => [center, ['Dian','Aza','Desy','Eni','Mega','Syahla','Dinda','Fizka','Rahmah'][i]]),
    // KLM (9 SA)
    ...Array(9).fill('KLM').map((center, i) => [center, ['Qhisti','Ibam','Hani','Sinta','Feni','Talitha','Ara','Lilis','Adit'][i]]),
    // KGD (6 SA)
    ...Array(6).fill('KGD').map((center, i) => [center, ['Sofia','Kenny','Farisa','Anna','Tyas','Kennia'][i]]),
    // BKP (6 SA)
    ...Array(6).fill('BKP').map((center, i) => [center, ['Ruth','Fitra','Nurul Zahra','Meiza','Zee','Tika'][i]]),
    // BHI (5 SA)
    ...Array(5).fill('BHI').map((center, i) => [center, ['Akhira Yuniar','Thessy','Triwik','Ririn','Riza'][i]]),
    // BTR (6 SA)
    ...Array(6).fill('BTR').map((center, i) => [center, ['Syifa','Zahra','Diah','Dila','Adji','zahra'][i]]),
    // TCT (3 SA)
    ...Array(3).fill('TCT').map((center, i) => [center, ['Ira','Dhea','Harby'][i]]),
    // PJT (4 SA)
    ...Array(4).fill('PJT').map((center, i) => [center, ['Reza','Alya','Ainna','Rizka'][i]]),
    // SBY (3 SA)
    ...Array(3).fill('SBY').map((center, i) => [center, ['Kevin','Alvi','Atha'][i]]),
    // BSD Bidex (8 SA)
    ...Array(8).fill('BSD Bidex').map((center, i) => [center, ['Jessi','Ferdi','Niluh','Addien','Mutia','Revi','Henni','Phylia'][i]]),
    // DPK Semesta (3 SA)
    ...Array(3).fill('DPK Semesta').map((center, i) => [center, ['Tsana','Adlina','Riskiana'][i]]),
    // BSDTK (2 SA)
    ...Array(2).fill('BSDTK').map((center, i) => [center, ['Phylia','Jessi'][i]]),
    // DPKS (2 SA)
    ...Array(2).fill('DPKS').map((center, i) => [center, ['Tsana','Riskiana'][i]]),
    // SGB (4 SA)
    ...Array(4).fill('SGB').map((center, i) => [center, ['Alvi','Atha','Rara','Rany'][i]]),
    // CKR (3 SA)
    ...Array(3).fill('CKR').map((center, i) => [center, ['Sinta','Fitra','Nida'][i]]),
    // SKL (2 SA)
    ...Array(2).fill('SKL').map((center, i) => [center, ['Dita','Fila'][i]]),
    // YOG (2 SA)
    ...Array(2).fill('YOG').map((center, i) => [center, ['Cecil','Bagy'][i]])
  ];
  
  // Header + Data
  const fullData = [['Center', 'Student Advisor'], ...saData];
  target.getRange(1, 1, fullData.length, 2).setValues(fullData);
  
  // Format
  target.getRange('1:1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  target.autoResizeColumns(1, 2);
  
  console.log(`✅ Vertical SA Table: ${fullData.length-1} records`);
}
