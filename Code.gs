function CustomVStack(range, skipEmpty) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Get the formula of the active range
  var formula = SpreadsheetApp.getActiveRange().getFormula();

  // Extract the arguments from the formula
  var args = formula.match(/=\w+\((.*)\)/i)[1].split(',');

  // Extract sheet name and range from the arguments
  var args = args[0].split('!');
  // If sheet name is provided, get the sheet with that name, else get the active sheet
  if (args.length === 2) {
    var sheetName = args[0].replace(/^'/, '').replace(/'$/, '');
    var range = args[1];
    var sheet = ss.getSheetByName(sheetName);
  }
  else {
    range = args[0];
    sheet = ss.getActiveSheet();
  }
  // Get the range object based on the extracted range and sheet
  var selectedRange = sheet.getRange(range);

  // Get the number of rows and columns in the selected range
  var numRows = selectedRange.getNumRows();
  var numCols = selectedRange.getNumColumns();

  // Get the values of the selected range
  var rangeValues = selectedRange.getValues();

  // Initialize an empty array to store the result
  var result = [];

  // Iterate through each column and row of the range
  for (var col = 0; col < numCols; col++) {
    for (var row = 0; row < numRows; row++) {
      // Check if the skipEmpty flag is set and if the current cell is not empty
      if (skipEmpty == true || skipEmpty == undefined) {
        // If the cell is not empty, push its value to the result array
        if (rangeValues[row][col] !== '') {
          result.push(rangeValues[row][col]);
        }
      }
      // If skipEmpty is not set, push the value to the result array regardless
      else {
        result.push(rangeValues[row][col]);
      }
    }
  }
  return result;
}