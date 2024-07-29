/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */
var CacheTodayCell = 'D13'; // Don't change this cell reference NEW
var CacheLastDayCell = 'E13'; // Don't change this cell reference NEW
var CacheTodayPosOnCal; // Cell position
var CacheLastDayPosOnCal; // Cell position
var CacheTodayValue; // Date value
var CacheLastDayValue; // Date value
var CacheErrorMessageCell = 'D21'; // Don't change this cell reference NEW
var CacheTodayValueCell = 'D18'; // Don't change this cell reference NEW
var SourceSheetNameCell = 'D3'; // Don't change this cell reference NEW
var NumTeamMembersCell = 'D6';
var NumTeamMembers;
 
var CacheSheetName = 'Settings*'
var SourceRangeStart;
var SourceRangeEnd;
var CacheSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CacheSheetName);
var SourceSheetName = CacheSheet.getRange(SourceSheetNameCell).getValue();
var SourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SourceSheetName);
var ScheduleSourceRange;
var ScheduleSourceValues;
var ScheduleSourceRowColArray = [];
// var SourceRangeRaw = SourceRangeStart + ':' + SourceRangeEnd;
var SourceRangeClean = ''
var OutputSheetName = 'Output';
var ToSlackSheetName = 'ToSlack'; // Added 'ToSlack' sheet name
var BackgroundOffice = '#8fd4f3';
var BackgroundWFH = '#a0dec4';

var IsCachedDateValid;
var IsCacheADate;
var IsCacheEndOfMonth;
var PatternsToRemove = ['T1 - ', 'T2 - '];
var Pattern = new RegExp(PatternsToRemove.join('|'), 'g');

/** test cache */
function TestCache(){
  //CacheLastDayPosOnCal = MoveToTheNextMonthEnd(CacheLastDayPosOnCal);
  NumTeamMembers = CacheSheet.getRange(NumTeamMembersCell).getValue();
  SpreadsheetApp.getUi().alert(NumTeamMembers);
  //CacheTodayPosOnCal = CacheSheet.getRange('B9').getValue();
  //var SourceTestAppend = [CacheTodayPosOnCal, CacheTodayValue];

  //CacheSheet.appendRow(SourceTestAppend);
}

/** Adds a custom menu item */
function addMenu() {
  var menu = SpreadsheetApp.getUi().createMenu('Schedule');
  menu.addItem('Re-Sync Schedule', 'ReSyncSchedule');
  menu.addItem('Test Cache', 'TestCache');
  menu.addToUi();
}

/** Initializes Menu */
function onOpen(e) {
  addMenu();
}

/** Function to check if a value is a date */
function isDate(value) {
  return Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value);
}

/** */
/** General function to move down by numRows */
/** */
function MoveCellDownByNRow(cellPosition, numRows) {
  // Extract the column part of the cell position
  var column = cellPosition.match(/[A-Z]+/)[0];
  
  // Extract the row part of the cell position and parse it as an integer
  var row = parseInt(cellPosition.match(/\d+/)[0]);
  
  // Increment the row number by the specified number of rows
  row += numRows;
  
  // Construct the new cell position by combining the column and row parts
  var newCellPosition = column + row;
  
  return newCellPosition;
}

/** */
/** General function to move right by 1 column */
/** */
function MoveCellRight(cellPosition) {
  // Extract the column part of the cell position
  var column = cellPosition.match(/[A-Z]+/)[0];
  
  // Extract the row part of the cell position
  var row = cellPosition.match(/\d+/)[0];
  
  // Convert the column letters to a numeric index
  var columnNumericIndex = 0;
  for (var i = 0; i < column.length; i++) {
    columnNumericIndex = columnNumericIndex * 26 + (column.charCodeAt(i) - 64); // ASCII code of 'A' is 65
  }
  
  // Increment the numeric index for the column
  columnNumericIndex++;
  
  // Convert the numeric index back to column letters
  var newColumn = '';
  while (columnNumericIndex > 0) {
    var remainder = columnNumericIndex % 26;
    if (remainder === 0) {
      newColumn = 'Z' + newColumn;
      columnNumericIndex = Math.floor(columnNumericIndex / 26) - 1;
    } else {
      newColumn = String.fromCharCode(remainder + 64) + newColumn; // ASCII code of 'A' is 65
      columnNumericIndex = Math.floor(columnNumericIndex / 26);
    }
  }
  
  // Construct the new cell position by combining the new column and row parts
  var newCellPosition = newColumn + row;
  
  return newCellPosition;
}

function MoveCellToColumnC(cellReference) {
  var row = cellReference.match(/\d+/)[0]; // Extract the row number using regular expression
  var newCellReference = 'C' + row; // Construct the new cell reference with column C
  
  return newCellReference;
}


/** */
/** General function to test if 2 dates are equal */
/** */
function TestDateEquals(date1, date2) {
  if (date1.getFullYear() === date2.getFullYear() &&
      date1.getMonth() === date2.getMonth() &&
      date1.getDate() === date2.getDate()) {
    // Dates have the same year, month, and day
    return true;
  } else {
    // Dates are different
    return false;
  }
} 

/** */
/** Finds a) Today's Date Cell Position b) Extract Date Value from a) */
/** */
function GetDateAndPosFromCalToday()
{
  CacheTodayPosOnCal = CacheSheet.getRange(CacheTodayCell).getValue();
  CacheTodayValue = SourceSheet.getRange(CacheTodayPosOnCal).getValue();

  CacheLastDayPosOnCal = CacheSheet.getRange(CacheLastDayCell).getValue();
  CacheLastDayValue = SourceSheet.getRange(CacheLastDayPosOnCal).getValue();
}

/** */
/** Stores the Schedule Source Values */
/** */
function GetSourceRangeAndValues()
{
  ScheduleSourceRange = SourceSheet.getRange(SourceRangeStart + ':' + SourceRangeEnd);
  ScheduleSourceValues = ScheduleSourceRange.getValues();
}

/** */
/** Moves to the next month  */
/** */
function MoveToTheNextMonth(CellPositionToday)
{
  // Reset to Column C
  CacheTodayPosOnCal = MoveCellToColumnC(CellPositionToday);

  // Move down 4 +  rows
  NumTeamMembers = CacheSheet.getRange(NumTeamMembersCell).getValue();

  // 5 will always be used as long as the Calendars remain consistent in format
  // NumTeamMembers should be the only variable across both SG and US teams
  CacheTodayPosOnCal = MoveCellDownByNRow(CacheTodayPosOnCal, NumTeamMembers + 5);

  return CacheTodayPosOnCal;
}

function MoveToTheNextMonthEnd(CellPositionLastDay)
{
  // Move down 4 +  rows
  NumTeamMembers = CacheSheet.getRange(NumTeamMembersCell).getValue();

  // 5 will always be used as long as the Calendars remain consistent in format
  // NumTeamMembers should be the only variable across both SG and US teams
  CacheLastDayPosOnCal = MoveCellDownByNRow(CellPositionLastDay, NumTeamMembers + 5);

  return CacheLastDayPosOnCal;
}

/** */
/** Checks if the cache is pointing at today's date */
/** */
function ValidateAndUpdateCachedDate(){

  var todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);

  if (isDate(CacheTodayValue))
  {
    /** If we have a valid date, we can check if it is the current date */
    /** We must do this or the error will stop all of the other validations that come after */
    if (TestDateEquals(CacheTodayValue, todayDate)) 
    {
      /** if CacheTodayValue is today, update the Date value on Settings Sheet*/
      CacheSheet.getRange(CacheErrorMessageCell).setValue('Success: Already on the current date.');
      CacheSheet.getRange(CacheTodayValueCell).setValue(CacheTodayValue);
    }
    else 
    {
      /** if not today, go to the next day - move right */
      CacheTodayPosOnCal = MoveCellRight(CacheTodayPosOnCal);
      CacheTodayValue = SourceSheet.getRange(CacheTodayPosOnCal).getValue();

      if (isDate(CacheTodayValue))
      {
        /** We have to test again if this new cell contains a date value */
        /** If it's not a date value, we need the user to provide the right cell position */
        if (TestDateEquals(CacheTodayValue, todayDate)) 
        {
          /** if CacheTodayValue is today, update the Date Position and Value on Settings Sheet */
          CacheSheet.getRange(CacheErrorMessageCell).setValue('Success: Moved to the next day. ');
          CacheSheet.getRange(CacheTodayCell).setValue(CacheTodayPosOnCal);
          CacheSheet.getRange(CacheTodayValueCell).setValue(CacheTodayValue);

          CacheLastDayValue = SourceSheet.getRange(CacheLastDayPosOnCal).getValue();
        }
        else
        {
          /** If it's a date value but not today, we should not attempt to move again. */
          /** Ask the user to provide a valid cell instead */
          CacheSheet.getRange(CacheErrorMessageCell).setValue('Error: Please provide valid cell position referencing the current date on the Calendar. See screenshot for example.');          
        }
      }
    
      else
      {
        /** if we get an empty cell after moving right, go to the next month - 1) move down 14 rows 2) move left to column C */
        CacheTodayPosOnCal = MoveToTheNextMonth(CacheTodayPosOnCal);
        CacheLastDayPosOnCal = MoveToTheNextMonthEnd(CacheLastDayPosOnCal);

        /** Use the new position to get the new Today Value */
        CacheTodayValue = SourceSheet.getRange(CacheTodayPosOnCal).getValue();
        CacheLastDayValue = SourceSheet.getRange(CacheLastDayPosOnCal).getValue();

        /** Update the Date Position and Value on Settings Sheet */
        CacheSheet.getRange(CacheErrorMessageCell).setValue('Success: Moved to the next month.');
        CacheSheet.getRange(CacheTodayCell).setValue(CacheTodayPosOnCal);
        CacheSheet.getRange(CacheTodayValueCell).setValue(CacheTodayValue);

        CacheSheet.getRange(CacheLastDayCell).setValue(CacheLastDayPosOnCal);
      }
    }
  } 
  else 
  {
    /** If we have no valid date, we need user to provide the right cell position */
    /** This error can surface when it's the last day of the entire sheet and there's no new date to move to */
    CacheSheet.getRange(CacheErrorMessageCell).setValue('Error: Please provide the correct sheet name or cell positions. If a new quarter or year has started, the sheet name (B3) needs to be updated');
  }
}


/**
 * Resync Function when Menu Button is clicked
 */
function ReSyncSchedule() {

  /** Initializes the latest cache range to look up */
  GetDateAndPosFromCalToday();

  /** Check and grab the cache for date positions first */
  ValidateAndUpdateCachedDate();

  /** Sets the SourceRangeStart from the Cache to use in  GetSourceRangeAndValues */
  SourceRangeStart = MoveCellDownByNRow(CacheTodayPosOnCal,3);
  SourceRangeEnd = MoveCellDownByNRow(CacheLastDayPosOnCal, 9 + 3);
  /** Initializes the latest source range to look up */
  GetSourceRangeAndValues();
  // SpreadsheetApp.getUi().alert(CacheTodayPosOnCal);

  /** Clear the 'Output' sheet before appending new data */
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OutputSheetName);
  outputSheet.clear();

  StoreRowColumnArray();
  var outputData = [];

  for (var i = 0; i < ScheduleSourceRowColArray.length; i++) {
    var Rownumber = ScheduleSourceRowColArray[i][0];
    var Columnnumber = ScheduleSourceRowColArray[i][1];
    var BackgroundColorItem = ScheduleSourceRowColArray[i][2];

    /** Only add a row if the person is working in Office or from home */
    if (BackgroundColorItem == BackgroundOffice || BackgroundColorItem == BackgroundWFH) {
      ExtractAndPaste(Rownumber, Columnnumber);
    } else {
      // Do nothing for non-matching conditions
    }
  }

  // Batch append rows
  if (outputData.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OutputSheetName).getRange(outputData.length, 1, outputData.length, outputData[0].length).setValues(outputData);
  }

  compileAndSortData();
}

function StoreRowColumnArray() {
  ScheduleSourceRowColArray = [];

  // Iterate through each cell in the range
  for (var i = 0; i < ScheduleSourceValues.length; i++) {
    for (var j = 0; j < ScheduleSourceValues[i].length; j++) {
      var Rownumber = ScheduleSourceRange.getCell(i + 1, j + 1).getRow(); // Adding 1 to convert from 0-based index to 1-based index
      var Columnnumber = ScheduleSourceRange.getCell(i + 1, j + 1).getColumn(); // Adding 1 to convert from 0-based index to 1-based index
      var Backgroundcoloritem = ScheduleSourceRange.getCell(i + 1, j + 1).getBackground();

      // Store row and column information in the array
      ScheduleSourceRowColArray.push([Rownumber, Columnnumber, Backgroundcoloritem]);
    }
  }

  // Log the resulting 2-dimensional array
  Logger.log(ScheduleSourceRowColArray);
}

function ExtractAndPaste(rowvalue, columnvalue) {
  var DateExtracted = '';

  // Check if rowvalue or columnvalue is valid
  if (rowvalue == null || columnvalue == null) {
    return;
  }

  /**Looks up column B to get the name, and remove the 'TX -' prefix */
  var rawNameCellValue = SourceSheet.getRange(rowvalue, 2).getValue();
  var RawName = rawNameCellValue.toString();
  var NameExtracted = RawName.replace(Pattern, '');

  /** Look up the date value */
  /** Loop upward until a cell with a date is found or until reaching the top of the sheet */
  while (rowvalue > 1) {
    /** Move one row up and get the value */
    CellValue = SourceSheet.getRange(rowvalue - 1, columnvalue).getValue();

    /** Check if the cell value is a date */
    if (isDate(CellValue)) {
      DateExtracted = Utilities.formatDate(CellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd'); // Format date as yyyy-MM-dd
      // Check if the date is today or in the future
      var today = new Date();
      today.setHours(0, 0, 0, 0);

      if (CellValue >= today) {
        var NameAndDate = [DateExtracted, NameExtracted];
        var targetsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OutputSheetName);

        targetsheet.appendRow(NameAndDate);
      }

      break;
    }
    /** Move to the next row up */
    rowvalue--;
  }

  // If the loop reaches the top and hasn't found a valid date, check if the last CellValue is a valid date
  if (rowvalue === 1 && isDate(CellValue) && CellValue >= today) {
    DateExtracted = Utilities.formatDate(CellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd'); // Format date as yyyy-MM-dd
    var NameAndDate = [DateExtracted, NameExtracted];
    var targetsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OutputSheetName);

    targetsheet.appendRow(NameAndDate);
  }
}


function compileAndSortData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OutputSheetName);
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Create an object to store names by date
  var compiledData = {};

  // Iterate through each row in the data
  for (var i = 0; i < data.length; i++) {
    var date = data[i][0];
    var name = data[i][1];

    // Check if date exists in the compiledData object
    if (!compiledData[date]) {
      compiledData[date] = [name];
    } else {
      // Add name to the existing array for the date
      compiledData[date].push(name);
    }
  }

  // Target ToSlack Sheet
  var compiledSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ToSlackSheetName);

  // Clear existing data in the compiled sheet
  compiledSheet.clear();

  // Write the compiled data to the new sheet, including the header row
  var compiledArray = [['Date', 'Person']];
  for (var date in compiledData) {
    var names = compiledData[date].join(', ');
    var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    compiledArray.push([formattedDate, names]);
  }

  // Append the compiled data to the new sheet
  compiledSheet.getRange(1, 1, compiledArray.length, 2).setValues(compiledArray);

  // Sort the data in the ToSlack sheet by date (column A)
  compiledSheet.getRange(2, 1, compiledArray.length - 1, 2).sort(1); // Exclude the header row from sorting
}

function fetchNamesForTodayAndSendToSlack() {
  // Open the Google Sheet by sheetname
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ToSlackSheetName);


  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();

  // Get today's date and format it
  var today = new Date();
  var formattedToday = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
  console.log(formattedToday);

  // Filter data for today's date and extract names
  var namesForToday = data
    .filter(function (row, index) {
      // Skip header row
      return index > 0;
    })
    .filter(function (row) {
      // Convert the date object to a string in the format 'YYYY-MM-DD'
      var formattedDate = Utilities.formatDate(row[0], "GMT+8", "yyyy-MM-dd");

      // Trim any extra spaces and compare
      return formattedDate.trim() === formattedToday.trim();
    })
    .map(function (row) {
      return row[1]; // Assuming the names are in the second column
    });

  // Log the names to the Google Apps Script logs
  Logger.log("Names for " + formattedToday + ":");
  Logger.log(JSON.stringify(namesForToday));

  // Send the names to Slack
  var message =
    "TEST from new Integration - People working today in SG (" +
    formattedToday +
    "):\n" +
    namesForToday.join("\n");
  sendToSlack(message);
}

function sendToSlack(message) {
  // Specify the Slack webhook URL
  var slackWebhookUrl =
    "yourwebhooklink";

  // Create payload with the message
  var payload = {
    text: message,
  };

  // Set up options for the HTTP request
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  // Make the HTTP request to send the message to Slack
  UrlFetchApp.fetch(slackWebhookUrl, options);
}
// Run the fetchNamesForTodayAndSendToSlack function to fetch names for today and send them to Slack.
// You can schedule this function to run on a daily basis using Apps Script triggers.
