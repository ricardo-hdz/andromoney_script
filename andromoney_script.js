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

//2016_ID
//var TRACKER_ID = '1AweSw57sw-a33ijYq9n7aQOR3xuizXeJoYExPKfMTQE';
//2017_ID
//var TRACKER_ID = '1NZWrWFLTJYJmvA_T0gq4--7DJc1L_GaF5XcS0N3Pgps';
// Europe_2017
//var TRACKER_ID = '1FyMr67nUW6yFL8NJFpguUCft9Q6f5RbC4CcsTIXfuME';
//2018 ID
var TRACKER_ID = '16ch-qMh1XGIVHZkRoAA4Rw4HUNqe70oigoGiVNyu8JQ';

var MAX_ROW_NUMBER = 200;
var MXN_CURRENCY = 'MXN';
var currencies = {
  'MXN': 'MXN',
  'EUR': 'EUR',
  'HUF': 'HUF',
  'CZK': 'CZK'
};

var exchangeRate = {
  'MXN': '18.6',
  'EUR': '0.94',
  'HUF': '290.52',
  'CZK': '25.45'
};

var spreadsheet;
var formattedSheet;
var sheet;
var currentMonth;

// Application Triggers

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
    {name: "Format & Transfer Data", functionName: "formatAndTransferData"},
    {name: "Format Data", functionName: "formatData"},
  ];
  spreadsheet.addMenu("AndroidMoney", menuEntries);
}

function onEdit(e) {
    var user = e.user;    
    var sessionUser = Session.getActiveUser().getEmail();
    // formatAndTransferData();
    SpreadsheetApp.getActiveSpreadsheet().toast('Edit made by ' + user, 'Edit', 5);
    formatData();
}

// Application

function formatAndTransferData() {
  transformAmount();
  copyFormattedData();
  copyDataToTracker();
}

function formatData() {
  transformAmount();
  copyFormattedData();
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
        if (currency !== 'USD' && currencies[currency] !== null) {
          if (exchangeRate[currency] === null) {
            // open prompts
            var ui = SpreadsheetApp.getUi();
            var response = ui.prompt('Exchange Rate', 'What is current exchange rate?', ui.ButtonSet.OK_CANCEL);

            if (response.getSelectedButton() == ui.Button.OK) {
              exchangeRate[currency] = response.getResponseText();
              transformAmountToExchangeRate(cell, exchangeRate[currency]);
            } else if (response.getSelectedButton() == ui.Button.CANCEL) {
             Logger.log('The user canceled the dialog.');
            } else {
             Logger.log('The user closed the dialog.');
            }
          } else {
            transformAmountToExchangeRate(cell, exchangeRate[currency]);
          }
        }
    }
}

function transformAmountToExchangeRate(cell, rate) {
  cell.setValue(cell.getValue() / rate);
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
  
  var lastRow = formattedSheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var cell = formattedSheet.getRange('A' + i);
    var dateString = cell.getValue() + '';
    
    var year     = dateString.substring(0,4);
    var month    = dateString.substring(4,6);
    var day      = dateString.substring(6,8);

    cell.setValue(month + "/" + day + "/" + year);
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
  var date = Utilities.formatDate(cell.getValue(), spreadsheet.getSpreadsheetTimeZone(), "MM/dd/YY"); 
  var monthStr    = date.substring(0,2);
  var month       = parseInt(monthStr, 10);

  return MONTHS[month];
}