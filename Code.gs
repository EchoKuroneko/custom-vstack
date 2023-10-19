function CustomVStack(range, skipEmpty) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Get the range object based on the provided range
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
      if (skipEmpty==true || skipEmpty == undefined) {
        // If the cell is not empty, push its value to the result array
        if (rangeValues[row][col] !== '') {
          result.push(rangeValues[row][col]);
        }
      }
      // If skipEmpty is not set, push the value to the result array regardless
      else{
        result.push(rangeValues[row][col]);
      }
    }
  }
  return result;
}