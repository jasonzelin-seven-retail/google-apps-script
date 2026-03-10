// Helper functions
const range = (start, end, step = 1) => {
  const length = Math.floor((end - start) / step) + 1; // Calculate the required length
  return Array.from({ length }, (_, index) => start + (index * step));
}

function addDaysToDate(date, days) {
  const d = new Date(date);    // always clone to avoid mutating original
  d.setDate(d.getDate() + days);
  return d;
}

function truncateToWeekMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1) - day;

  const result = new Date(d);
  result.setDate(d.getDate() + diff);
  result.setHours(0, 0, 0, 0);

  return result;
}

function lookupLastData(array, lookupValue, lookupCol, sortCol) {
  let result = array  
                .filter(row => row[lookupCol] == lookupValue)
                .map(a => `${a[lookupCol]}.${a[sortCol]}`)
                .sort()
                .map(a => [a.split('.')[lookupCol], a.split('.')[sortCol]])
                .at(-1);
  return result;
}

// Main functions
function generateSurveySchedule() {
  var gsheet_key = '1R_-hi6Ncp1QwQyfc3aRbAzXuJO1SdoqzCnGWK8YSBYE'
  var spreadsheet = SpreadsheetApp.openById(gsheet_key)
  var sourceSheet =  spreadsheet.getSheetByName(name='Combined Class Database_v2')
  var outSheet = spreadsheet.getSheetByName(name="Schedule");

  var data = sourceSheet.getRange('A2:AF').getValues()
  var outData = outSheet.getRange('A2:W').getValues().map(row => [2, 22].map(i => row[i])) // Get only class id and blast date
  var indexArr = [0,10,2,7,11,13,1,26,27,28,29,31]
  // Filtering only relevant columns and rows
  data = data.map(row => indexArr.map(i => row[i])).filter(rows => rows[6] == 'Active' & rows[10] == false & rows[11] < rows[9] & rows[5] != '')
  var outputRows = [];
  var newDate = None;

  data.forEach(row => {
    const [center, sa, classId, className, startDate, graduationDate, status, surveyStartDate, surveyStartWeek, plannedSurveyCount, createSchedule, actualSurveyCount] = row // assigning values from each column to respective column name to improve code readibility
    if ((createSchedule === false || createSchedule === "false") & plannedSurveyCount > 0) {
      if (actualSurveyCount === 0) {
        newDate = surveyStartDate
        outputRows.push([center, sa, classId, className, startDate, graduationDate, status, newDate]);
      }
      else if (actualSurveyCount > 0) {
        var lastSurveyDate = lookupLastData(outData, classId, 0, 1)
        if (lastSurveyDate.getDate() >= 15) {
          deltaDays = 28
        }
        else {
          deltaDays = 31
        }

        newDate = truncateToWeekMonday(addDaysToDate(date=lastSurveyDate, days=deltaDays))
        outputRows.push([center, sa, classId, className, startDate, graduationDate, status, newDate]);
      };
    }
  });

  // Filter only schedule blast date in next week
  const today = new Date();
  const nextWeek = addDaysToDate(date=truncateToWeekMonday(today), days=7);
  outputRows = outputRows.filter(row => row[7] == nextWeek);

  // Insert an empty row to signify separation of each data input
  var lastRow = outSheet.getLastRow(outData, classId, 0, 1);
  outSheet.insertRowAfter(lastRow);
  outSheet.insertRowAfter(lastRow);
  // Recalculate last row after the insertion 
  lastRow = lastRow + 2

  // Append all generated schedules to the Schedule sheet
  if (outputRows.length > 0) {
    outSheet.getRange(`A${lastRow}:F${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => [0,1,2,3,4,5].map(i => row[i])));
    outSheet.getRange(`H${lastRow}:H${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => [row[6]]));
    outSheet.getRange(`W${lastRow}:W${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => [row[7]]));

  }
}