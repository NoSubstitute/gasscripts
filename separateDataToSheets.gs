function separateDataToSheets() {
// These first five const need to be edited to fit your data
  const dataSheet = "Students"; // Name of the sheet we pull data from
  const dataRange = "A1:G"; // Data we want to separate into sheets
  const dataColumnNr = 5; // Column we pull separate data from to getRange 5 = E:E
  const dataColumn = "E:E"; // Column we pull separate data from to ourFormula - yes, I know E = 5, and yes, we need to set it twice
  const ourHeaders = "A1:G1"; // Cells to pull the headers we wish to copy to each sheet
  const sortColumns = "6;TRUE;1;TRUE" // Sort vaules in sheets by which columns

// The following should most likely not need to be changed, unless you want a completely other formula structure
// Then perhaps edit line 42 where ourFormula is set

  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get current spreadsheet
  const sourceWS = ss.getSheetByName(dataSheet); // Set the sheet we pull data from

  const valueName = sourceWS
    .getRange(2, dataColumnNr, sourceWS.getLastRow() - 1, 1) // Get values from dataColumnNr from row 2 till the end
    .getValues() // Get all values into an array of arrays
    .map(myValue => myValue[0]); // Turn it into a single array

  const uniqueValueName = [...new Set(valueName)]; // Create a set, deconstruct it (...) = find unique values, turn back into array []

  // console.log(uniqueValueName); // Log unique values

  const uniqueValueNameSorted = uniqueValueName.sort(); // Sort the uniqueValueName
  
  // console.log(uniqueValueNameSorted); // Log unique values sorted

  const currentSheetNames = ss.getSheets().map(s => s.getName()); // Get list of current sheets, turn into array

  // console.log(currentSheetNames); // Log names of current sheets

  let ws; // creating the variable outside the loop
  let ourFormula; // creating the variable outside the loop

  uniqueValueNameSorted.forEach(valueName => { // Will run though values alphabetically

    if (!currentSheetNames.includes(valueName)) {
      ws = null; // Reset ws
      ws = ss.insertSheet(); // Create a new sheet
      ws.setName(valueName); // Set the name of the sheet
      ourFormula = `=SORT(FILTER(${dataSheet}!${dataRange};${dataSheet}!${dataColumn}="${valueName}");${sortColumns})`;
      // Using ` when setting ourFormula is a way to make sure internal " and ' aren't broken
      ws.getRange("A2").setFormula(ourFormula); // Add the formula to A2
      sourceWS.getRange(ourHeaders).copyTo(ws.getRange(ourHeaders)); // Copy the headers of source sheet

      // Deleting empty columns
      var maxColumns = ws.getMaxColumns(); 
      var lastColumn = ws.getLastColumn();
      ws.deleteColumns(lastColumn+1, maxColumns-lastColumn); 

      // Remove empty rows
      var maxRows = ws.getMaxRows(); 
      var lastRow = ws.getLastRow();
      ws.deleteRows(lastRow+1, maxRows-lastRow);

      // Adding filters & freeze first row
      ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).activate();
      ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).createFilter();
      ws.getRange('A1').activate();
      ws.setFrozenRows(1);
      
      // Resizing columns
      for (var i=1; i<ws.getMaxColumns()+1; i++){
      ws.autoResizeColumn(i); // Autoresize each column
      var currentwidth = ws.getColumnWidth(i);
      ws.setColumnWidth (i, currentwidth+25); // This extra width is necessary after setting filter, else header text may be partially hidden
      };

    } // If valueName doesn't exist in the list currentSheetNames, forEach loop through the list of valueNames
  }); // create sheet & copy first row and add formula - If valueName exists do nothing for that valueName
} // Close createSheets function

// Idea from Learn Google Spreadsheets video
// https://www.youtube.com/watch?v=QTySwuhpHG0
