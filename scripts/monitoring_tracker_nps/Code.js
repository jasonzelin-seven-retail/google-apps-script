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

// Main functions
function generateSurveySchedule() {
  var gsheet_key = '1R_-hi6Ncp1QwQyfc3aRbAzXuJO1SdoqzCnGWK8YSBYE'
  var spreadsheet = SpreadsheetApp.openById(gsheet_key)
  var sourceSheet =  spreadsheet.getSheetByName(name='Combined Class Database_v2')
  var outSheet = spreadsheet.getSheetByName(name="Schedule");

  var data = sourceSheet.getRange('A2:AD').getValues()
  var indexArr = [0,10,2,7,11,13,1,26,27,28,29]
  // Filtering only relevant columns and rows
  data = data.map(row => indexArr.map(i => row[i])).filter(rows => rows[6] == 'Active' & rows[10] == false)
  const outputRows = [];
  // console.log(data)

  data.forEach(row => {
    const [center, sa, classId, className, startDate, graduationDate, status, surveyStartDate, surveyStartWeek, surveyCount, createSchedule] = row // assigning values from each column to respective column name to improve code readibility
    if ((createSchedule === false || createSchedule === "false") & surveyCount > 0) {
      var newDate = surveyStartDate
      var deltaDays = 0

      for (let i = 0; i < surveyCount; i++) {
        if (i == 0) {
          deltaDays = 0
        }
        else if (newDate.getDate() >= 15) {
          deltaDays = 28
        }
        else {
          deltaDays = 31
        }

        newDate = truncateToWeekMonday(addDaysToDate(date=newDate, days=deltaDays))
        outputRows.push([center, sa, classId, className, startDate, graduationDate, status, newDate]);
      }
    }
  });
  var lastRow = outSheet.getLastRow()

  // Insert an empty row to signify separation of each data input
  outSheet.insertRowAfter(lastRow)
  outSheet.insertRowAfter(lastRow)
  // Recalculate last row after the insertion 
  lastRow = lastRow + 2

  // Append all generated schedules to the Schedule sheet
  var indexArr0 = [0,1,2,3,4,5]
  if (outputRows.length > 0) {
    outSheet.getRange(`A${lastRow}:F${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => indexArr0.map(i => row[i])));
    outSheet.getRange(`H${lastRow}:H${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => [row[6]]));
    outSheet.getRange(`W${lastRow}:W${lastRow + outputRows.length - 1}`)
            .setValues(outputRows.map(row => [row[7]]));

  }
}