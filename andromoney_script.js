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

  var ENDPOINT_RATES = 'http://apilayer.net/api/live?access_key=b9f923b9b69e956ea34daa10694fc9b1&source=USD&currencies={currency}&format=1';

  function getLatestExchangeRate(currency) {
	var url = ENDPOINT_RATES.replace('{currency}', currency);
	var response = UrlFetchApp.fetch(url);
	var data = JSON.parse(response.getContentText());
	var rate = 0.0;
	if (data && data.quotes && data.quotes.USDMXN) {
		var rate = data.quotes.USDMXN;
	}
	return rate;
  }

  var exchangeRates = {};

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
	  var columnProject = 'K';
	  var columnCurrency = 'B';
	  var columnAmount = 'C';
	  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	  sheet = spreadsheet.getSheetByName('AndroMoney');

	  var id;
	  var cellValue;
	  var lastRow = sheet.getLastRow();
	  var cellCurrency;
	  var currency;
	  var project;

	  for (var i = 3, amount; (amount = sheet.getRange(columnAmount + i).getValue()); i++) {
		  project = sheet.getRange(columnProject + i).getValue();
		  while (project === 'Business') {
			  sheet.deleteRow(i);
			  project = sheet.getRange(columnProject + i).getValue();
			  amount = sheet.getRange(columnAmount + i).getValue();
		  }
		  // Column expense type check if its either a expense or income
		  cellValue = sheet.getRange(columnExpenseType + i).getValue();
		  cell = sheet.getRange(columnAmount + i);
		  if (cellValue != '') {
			  cell.setValue(0 - Math.abs(amount));
		  } else {
			  //income always come as empty string
			  cell.setValue(0 + Math.abs(amount));
		  }

		  cellCurrency = sheet.getRange(columnCurrency + i);
		  currency = cellCurrency.getValue();

		  // Open prompt
		  if (currency !== 'USD') {
			if (!exchangeRates.hasOwnProperty(currency)) {
				var rate = getLatestExchangeRate(currency);
				exchangeRates[currency] = rate;
			}
			transformAmountToExchangeRate(cell, exchangeRates[currency]);
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
	var cell = sheet.getRange('A2').getValue();
	var date = Utilities.formatDate(cell, spreadsheet.getSpreadsheetTimeZone(), "MM/dd/YY");
	var monthStr    = date.substring(0,2);
	var month       = parseInt(monthStr, 10);

	return MONTHS[month];
  }