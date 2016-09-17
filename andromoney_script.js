var HEADING_COLUMN = {
  '6': '1', //Date
  '15': '2', //Type (default to personal)
  '4': '3', //Category
  '5': '4', //Subcategory
  '12': '5', //Vendor
  '7': '6', //Payment (check if empty)
  '2': '7', //Currency
  '3': '8', //Amount
  '9': '9' //Note
};

var MONTHS = [
  'Test',
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec'
];


var TRACKER_ID = '1AweSw57sw-a33ijYq9n7aQOR3xuizXeJoYExPKfMTQE';

var MAX_ROW_NUMBER = 200;
var MXN_CURRENCY = 'MXN';

var spreadsheet;
var formattedSheet;
var sheet;
var exchangeRate = null;

function onOpen() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  formattedSheet = spreadsheet.getSheetByName('Formatted');
  if (formattedSheet == null) {
    spreadsheet.insertSheet('Formatted');
    formattedSheet = spreadsheet.getSheetByName('Formatted');
  }
  sheet = spreadsheet.getSheetByName('AndroMoney');
  spreadsheet.setActiveSheet(sheet);

  var menuEntries = [
    {name: "Format & Transfer Data", functionName: "formatAndTransferData"}
  ];
  spreadsheet.addMenu("AndroidMoney", menuEntries);
}

function formatAndTransferData() {
  transformAmount();
  copyFormattedData();
  copyDataToTracker();
}

/*
* Assigns a positive or negative value to the amount according to
* type of transaction: income or expense
*/
function transformAmount() {
    var columnExpenseType = 'G';
    var columnCurrency = 'B';
    var columnAmount = 'C';
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet = spreadsheet.getSheetByName('AndroMoney');

    var id;
    var cellValue;
    var lastRow = sheet.getLastRow();
    var cellCurrency;
    var currency;

    for (var i = 3; i <= lastRow; i++) {
        // Column expense type check if its either a expense or income
        id = columnExpenseType + i;
        cell = sheet.getRange(id);
        cellValue = cell.getValue();
        id = columnAmount + i;

        cellCurrency = sheet.getRange(columnCurrency + i);
        currency = cellCurrency.getValue();

        if (cellValue != '') {
            cell = sheet.getRange(id);
            cell.setValue(0 - Math.abs(cell.getValue()));
        } else {
            cell = sheet.getRange(id);
            cell.setValue(0 + Math.abs(cell.getValue()));
        }

        // Open prompt
        if (currency === MXN_CURRENCY) {
          if (exchangeRate === null) {
            // open prompts
            var ui = SpreadsheetApp.getUi();
            var response = ui.prompt('Exchange Rate', 'What is current exchange rate?', ui.ButtonSet.OK_CANCEL);

            if (response.getSelectedButton() == ui.Button.OK) {
              exchangeRate = response.getResponseText();
              transformAmountToExchangeRate(cell);
            } else if (response.getSelectedButton() == ui.Button.CANCEL) {
             Logger.log('The user canceled the dialog.');
            } else {
             Logger.log('The user closed the dialog.');
            }
          } else {
            transformAmountToExchangeRate(cell);
          }
        }
    }
}

function transformAmountToExchangeRate(cell) {
  cell.setValue(cell.getValue / exchangeRate);
}

/*
 * Copy the raw data to formatted spreadsheet
 */
function copyFormattedData() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  formattedSheet = spreadsheet.getSheetByName('Formatted');
  sheet = spreadsheet.getSheetByName('AndroMoney');

  for (var columnSource in HEADING_COLUMN) {
    var values = sheet.getRange(3, columnSource, MAX_ROW_NUMBER); //getRange(row, column, numRows)
    var targetColumn = HEADING_COLUMN[columnSource];
    values.copyValuesToRange(formattedSheet, targetColumn, targetColumn, 2, MAX_ROW_NUMBER);
  }
}

/*
* Copies all data to tracker spreadsheet
*/
function copyDataToTracker() {
  // Get source data
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  formattedSheet = spreadsheet.getSheetByName('Formatted');
  var sourceData = formattedSheet.getDataRange().getValues()

  var targetSheetName = getCurrentMonth();

  // Copy data to tracker
  var trackerSpreadsheet = SpreadsheetApp.openById(TRACKER_ID).getSheetByName(targetSheetName);
  var targetRangeTop = trackerSpreadsheet.getLastRow();
  trackerSpreadsheet.getRange(1,1, sourceData.length, sourceData[0].length).setValues(sourceData);
}

/*
* Get the current month id for the expenses
* @returns {string}
*/
function getCurrentMonth() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getSheetByName('Formatted');

  var cell = sheet.getRange('A2');
  var dateString = cell.getValue() + '';

  var year        = parseInt(dateString.substring(0,4), 10);
  var monthStr    = dateString.substring(4,6);
  var month       = parseInt(monthStr, 10);
  var day         = parseInt(dateString.substring(6,8), 10);

  return MONTHS[month];
}