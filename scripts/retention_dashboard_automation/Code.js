// Global variables
const financeManagementGsheetKey = '1EZ6X6wsaSCvcm8gteGOmb7uaYjxheWRDtDQmDxn8AuM'

// Helper functions
const range = (start, end, step = 1) => {
  const length = Math.floor((end - start) / step) + 1; // Calculate the required length
  return Array.from({ length }, (_, index) => start + (index * step));
}

function refreshBQConnector(centerSpreadsheet) {
  const dataSources = centerSpreadsheet.getDataSources()
  
  SpreadsheetApp.enableBigQueryExecution()

  dataSources.forEach(ds => {
    ds.refreshAllLinkedDataSourceObjects()
  });
}

// Main functions
function generateRetentionData() {
  const mappingSpreadsheet = SpreadsheetApp.openById(id=financeManagementGsheetKey)
  const mappingSheet = mappingSpreadsheet.getSheetByName(name='url mapping')
  const mappingData = mappingSheet.getRange('H2:K').getValues().filter(rows => rows[0] != '') //.map(row => indexArr.map(row[i]))
  
  mappingSisDict = {}
  mappingRetentionDict = {}
  mappingData.forEach(row => {
    const [center, keyDashboard, keySis, keyRetention] = row
    mappingSisDict[center] = keySis
    mappingRetentionDict[center] = keyRetention
  })

  centers = Object.keys(mappingSisDict)
  
  // Looping through each center, checking for students & classes for whom needs to be generated their retention data
  centers.forEach((center) => {
    console.log(`Generating center ${center}...`)
    const sisSpreadsheet = SpreadsheetApp.openById(id=mappingSisDict[center])
    const sisStudentSheet = sisSpreadsheet.getSheetByName(name='Student Database')
    const sisStudentData = sisStudentSheet.getRange('A2:H').getValues()

    const retentionSpreadsheet = SpreadsheetApp.openById(mappingRetentionDict[center])
    const retentionSheet = retentionSpreadsheet.getSheetByName(name='Retention Data')
    const retentionData = retentionSheet.getRange('A2:I').getValues()

    // Filter when student + class not in retention
    var filteredRetentionData = retentionData.filter(rows => (rows[8] != ''))
    var filteredStudentData = sisStudentData.filter(rows => (rows[7] != '' & rows[0] == 'Active'))

    var sisIds = []
    filteredStudentData.forEach(row =>{
      const [status, studentId, name, dob, age, parentName, phoneNumber, currentClassId] = row
      sisIds.push([studentId, name, currentClassId, 'Not yet been asked', `${studentId}${currentClassId}`])
    })

    var retentionIds = []
    filteredRetentionData.forEach(row =>{
      const [studentId, studentName, dob, age, program, level, schedule, weekendWeekday, classId] = row
      retentionIds.push(`${studentId}${classId}`)
    })

    const newRetentionIds = sisIds.filter(rows => !retentionIds.includes(rows[4])) // Filter SIS student + class IDs that do not exist in retention data
    console.log(newRetentionIds)
    const lastRow = retentionData.length + 1
    console.log(lastRow)
    try {
      retentionSheet.getRange(`A${lastRow + 1}:B${lastRow + newRetentionIds.length}`)
                    .setValues(newRetentionIds.map(row => [0,1].map(i => row[i])))
      retentionSheet.getRange(`I${lastRow + 1}:I${lastRow + newRetentionIds.length}`)
                    .setValues(newRetentionIds.map(row => [row[2]]))
      retentionSheet.getRange(`N${lastRow + 1}:N${lastRow + newRetentionIds.length}`)
                    .setValues(newRetentionIds.map(row => [row[3]]))
    }
    catch(e) {
      console.log(`All students + classes in center ${center} are already in retention.`)
    }

    // Check for each row to automatically set retention status
    const originRetentionStatuses = retentionSheet.getRange('L4:N').getValues()
    var newRetentionStatuses = []
    originRetentionStatuses.forEach(row =>{
      const [creditRemaining, poolFlag, retentionStatus] = row
      if (retentionStatus != 'Not yet been asked') {
        newRetentionStatuses.push([retentionStatus])
      }
      else if (creditRemaining > 0) {
        newRetentionStatuses.push(['Already paid for 2C/4C'])
      }
      else {
        newRetentionStatuses.push([retentionStatus])
      }
    })
    retentionSheet.getRange('N4:N').setValues(newRetentionStatuses)
    
    // // Refresh BQ connector for center spreadsheet
    refreshBQConnector(retentionSpreadsheet)
    console.log(`Done generating center ${center}!`)
  })
}

// Migration function
function migrateRetention(destinationCenter) {
  // ---------------------- More Ideal & Sustainable way to get key mappings ----------------------
  const mappingSpreadsheet = SpreadsheetApp.openById(id='1EZ6X6wsaSCvcm8gteGOmb7uaYjxheWRDtDQmDxn8AuM')
  const mappingSheet = mappingSpreadsheet.getSheetByName(name='url mapping')
  // const indexArr = [0,3,4]
  const mappingData = mappingSheet.getRange('F2:H').getValues() //.map(row => indexArr.map(row[i]))
  
  mappingSisDict = {}
  mappingRetentionDict = {}
  mappingData.forEach(row => {
    const [center, keySis, keyRetention] = row
    mappingSisDict[center] = keySis
    mappingRetentionDict[center] = keyRetention
  })

  // ---------------------- More Ideal & Sustainable way to get key mappings ----------------------

  const destinationCenterGsheetId = mappingRetentionDict[destinationCenter]
  const destinationCenterSpreadsheet = SpreadsheetApp.openById(id=destinationCenterGsheetId)
  try {
    destinationCenterSpreadsheet.deleteSheet(destinationCenterSpreadsheet.getSheetByName(name='Retention Data')) // Remove pre-existing
  }
  catch(err) {
    null
  }

  const bsdTemplateGsheetId = '1tHjgEG8YsUYwcB9cRJBsjeMcFRHDe0fM7borItpfyG4'
  const bsdTemplateSpreadsheet = SpreadsheetApp.openById(id=bsdTemplateGsheetId)

  const sheetNamesToBeCopied = [
    '[Connector] retention_credits_data',
    '[Connector] attendance_rate_data',
    'extract_query_params',
    'SIS - Combined Class DB',
    'SIS - Student Database',
    'SIS - Class Database',
    'CX - Churn Analysis',
    'Acad - Progress Test',
    'Retention Data'
  ]
  sheetNamesToBeCopied.forEach(sheetName => {
    const bsdTemplateQueryParams = bsdTemplateSpreadsheet.getSheetByName(name=sheetName)
    bsdTemplateQueryParams.copyTo(destinationCenterSpreadsheet)

    const copySheet = destinationCenterSpreadsheet.getSheetByName(name=`Copy of ${sheetName}`).activate()
    try {
      copySheet.setName(sheetName)
    }
    catch(err) {
      null
    }

    destinationCenterSpreadsheet.moveActiveSheet(1) // Moving the copied sheet to leftmost position
    if (sheetName == 'extract_query_params') { // Replacing center name for each gsheet to define query params
      try {
        copySheet.getRange('B1').setValue(destinationCenter)
      }
      catch(err) {
        null
      }
    }
    else if ((sheetName == 'SIS - Student Database') || (sheetName == 'SIS - Class Database')) { // Replacing source gsheet for SIS tabs
      try {
        copySheet.getRange('B1').setValue(`https://docs.google.com/spreadsheets/d/${mappingSisDict[destinationCenter]}`)
      }
      catch(err) {
        null
      }
    }
    else if (sheetName == 'Retention Data') { // Replacing source gsheet for SIS tabs
      try {
        copySheet.getRange('A5:BF').clearContent()
      }
      catch(err) {
        null
      }
    }
  })
}

function migrateCenters() {
  destinationCenters = [
    'BSDTK'
  ]
  destinationCenters.forEach(destinationCenter => {
    console.log(`Migrating center ${destinationCenter}...`)
    migrateRetention(destinationCenter)
    console.log(`Done migrating center ${destinationCenter}!`)
  })
}