
  // Clear any previous highlights and reveal all rows initially
  range.setBackground(null);
  sheet.showRows(1, sheet.getMaxRows());

  var relevanceArray = []; // Array to store rows with relevance scores

  // Loop through the values starting from row 4 to search for the desired value
  for (var i = 3; i < values.length; i++) { // Start from row index 3 (which corresponds to row 4 in the sheet)
    var relevanceScore = 0; // Initialize relevance score for the row

    for (var j = 0; j < values[i].length; j++) {
      // Convert both cell content and search value to lower case for case-insensitive partial match
      if (values[i][j].toString().toLowerCase().indexOf(searchValue.toLowerCase()) !== -1) {
        relevanceScore++; // Increment relevance score for each match
        Logger.log('Partial match found in Row: ' + (i + 1) + ', Column: ' + (j + 1));
        sheet.getRange(i + 1, j + 1).setBackground('yellow'); // Highlight the found cell
      }
    }

    // Store the row data and its relevance score
    if (relevanceScore > 0) {
      relevanceArray.push({ rowIndex: i + 1, relevance: relevanceScore, rowData: values[i] });
    }
  }

  // Sort the rows based on relevance score in descending order
  relevanceArray.sort(function (a, b) {
    return b.relevance - a.relevance;
  });

  // Move relevant rows to start from row 4
  for (var k = 0; k < relevanceArray.length; k++) {
    var rowIndex = relevanceArray[k].rowIndex;
    var rowData = relevanceArray[k].rowData;

    // Clear the original row data
    sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).clearContent();

    // Move the row data to start from row 4
    sheet.getRange(k + 4, 1, 1, rowData.length).setValues([rowData]);
  }

  Logger.log('Rows sorted by relevance and moved to start from row 4.');
}
