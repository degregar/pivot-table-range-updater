// Function to update all pivot table ranges in the active spreadsheet
function updateAllPivotTableRanges() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // Iterate through each sheet in the spreadsheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var pivotTables = sheet.getPivotTables();
    
    // Iterate through each pivot table in the current sheet
    for (var j = 0; j < pivotTables.length; j++) {
      try {
        var pivotTable = pivotTables[j];
        var sourceSheet = pivotTable.getSourceDataRange().getSheet();
        
        // Determine the last row and column of the source sheet
        var lastRow = sourceSheet.getMaxRows();
        var lastColumn = sourceSheet.getLastColumn();
        
        // Define the new source range for the pivot table
        var newSourceRange = sourceSheet.getRange(1, 1, lastRow, lastColumn);
        
        // Get the anchor cell for the new pivot table based on the old one
        var anchorCell = sheet.getRange(pivotTable.getAnchorCell().getRow(), pivotTable.getAnchorCell().getColumn());
        
        // Create a new pivot table with the updated data range at the anchor cell location
        var newPivotTable = anchorCell.createPivotTable(newSourceRange);
        
        // Recreate the configuration of the old pivot table on the new one
        recreatePivotTableConfiguration(pivotTable, newPivotTable);
        
        Logger.log("Updated pivot table in sheet: " + sheet.getName());
      } catch (error) {
        Logger.log("Error updating pivot table in sheet " + sheet.getName() + ": " + error.toString());
      }
    }
  }
  
  Logger.log("Completed updating pivot tables.");
}

// Function to recreate the configuration of an old pivot table on a new one
function recreatePivotTableConfiguration(oldPivotTable, newPivotTable) {
  // Recreate filters from the old pivot table
  oldPivotTable.getFilters().forEach(function(filter) {
    newPivotTable.addFilter(filter.getSourceDataColumn(), filter.getFilterCriteria());
  });
  
  // Recreate row groups from the old pivot table
  oldPivotTable.getRowGroups().forEach(function(group) {
    newPivotTable.addRowGroup(group.getSourceDataColumn());
  });
  
  // Recreate column groups from the old pivot table
  oldPivotTable.getColumnGroups().forEach(function(group) {
    newPivotTable.addColumnGroup(group.getSourceDataColumn());
  });
  
  // Recreate pivot values from the old pivot table
  oldPivotTable.getPivotValues().forEach(function(value) {
    if (value.getFormula()) {
      // Add calculated pivot value if formula exists
      newPivotTable.addCalculatedPivotValue(value.getName(), value.getFormula());
    } else {
      // Add pivot value based on summarized function
      var summarizeFunction = value.getSummarizedBy(); // Corrected line
      newPivotTable.addPivotValue(value.getSourceDataColumn(), summarizeFunction);
    }
  });

  // Set the display orientation of pivot values to match the old pivot table
  newPivotTable.setValuesDisplayOrientation(oldPivotTable.getValuesDisplayOrientation());
}
