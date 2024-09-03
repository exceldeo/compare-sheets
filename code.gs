function onInstall(e){
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Compare Sheets', 'compareMultipleSheets')
    .addToUi();
}

function compareMultipleSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(sheet) {
    return sheet.getName();
  });

  // Show the HTML dialog for sheet and action selection
  var html = HtmlService.createHtmlOutputFromFile('SheetAndActionSelection')
      .setWidth(400)
      .setHeight(600);
  html.setTitle('Select Sheets and Action');
  html.setContent(html.getContent().replace('<!--SHEET_NAMES-->', JSON.stringify(sheetNames)));
  SpreadsheetApp.getUi().showModalDialog(html, 'Choose Sheets and Action');
}

function columnLetterToIndex(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column - 1;
}

function processComparison(mainSheetName, selectedSheets, action, columnOption, startColumn, endColumn, specificColumns) {
  if (!mainSheetName) {
    throw new Error('Main sheet name is not provided.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(mainSheetName);

  if (!mainSheet) {
    throw new Error('Main sheet not found: ' + mainSheetName);
  }

  var mainRange = mainSheet.getDataRange();
  var mainValues = mainRange.getValues();
  if (mainValues.length === 0) {
    throw new Error('Main sheet is empty: ' + mainSheetName);
  }

  var columnIndices = [];

  if (columnOption === 'all') {
    columnIndices = Array.from({ length: mainValues[0].length }, (v, k) => k);
  } else if (columnOption === 'range') {
    var start = columnLetterToIndex(startColumn);
    var end = columnLetterToIndex(endColumn);
    columnIndices = Array.from({ length: end - start + 1 }, (v, k) => start + k);
  } else if (columnOption === 'specific') {
    columnIndices = specificColumns.split(',').map(function(col) {
      return columnLetterToIndex(col.trim());
    });
  }

  var differences = [];
  // Set cancellation flag to false
  PropertiesService.getScriptProperties().setProperty('CANCEL_PROCESS', 'false');

  try {
    selectedSheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getDataRange();
      var values = range.getValues();

      var maxRows = mainValues.length;

      // Compare the two ranges
      for (var row = 0; row < maxRows && row < values.length; row++) {
        var rowHasDifference = false;
        columnIndices.forEach(function(col) {
          var mainValue = mainValues[row][col];
          var compareValue = values[row][col];

          // Convert dates to strings in a specific format for comparison
          if (mainValue instanceof Date && compareValue instanceof Date) {
            mainValue = Utilities.formatDate(mainValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            compareValue = Utilities.formatDate(compareValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }

          if (mainValue !== compareValue) {
            Logger.log({
              "main value": mainValue,
              "compare value": compareValue,
              "row": row,
              "col": col
            });

            if (action !== 'summary') {
              // Highlight the cell in the main sheet
              sheet.getRange(row + 1, col + 1).setBackground('yellow');
            }
            rowHasDifference = true;
          }
        });
        if (rowHasDifference) {
          differences.push({
            sheet: sheetName,
            row: row + 1,
            status: 'different',
            masterData: mainValues[row],
            data: values[row]
          });
        }
      }

      // Highlight any additional data below the range of the main sheet
      if (values.length > maxRows) {
        for (var row = maxRows; row < values.length; row++) {
          differences.push({
            sheet: sheetName,
            row: row + 1,
            status: 'missing',
            masterData: [],
            data: values[row]
          });
          for (var col = 0; col < values[row].length; col++) {
            if (action !== 'summary') {
              // Highlight the cell in the main sheet
              sheet.getRange(row + 1, col + 1).setBackground('yellow');
            }
          }
        }
      }
    });
    
    Logger.log({
      "column":columnIndices
    });

    if (action === 'summary') {
      // Create a summary sheet
      var summarySheet = ss.getSheetByName('Comparison Summary') || ss.insertSheet('Comparison Summary');
      summarySheet.clear();
      summarySheet.appendRow(['Sheet', 'Status', 'Row', 'Data']);

      differences.forEach(function(diff) {
        var appendData = [
          diff.sheet,
          diff.status,
          diff.row,
        ];
        diff.data.forEach(function(cell, index) {
          var mainValue = diff.masterData[index];
          var compareValue = cell;

          appendData.push(cell);

          // Convert dates to strings in a specific format for comparison
          if (mainValue instanceof Date && compareValue instanceof Date) {
            mainValue = Utilities.formatDate(mainValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            compareValue = Utilities.formatDate(compareValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }

          if (columnIndices.includes(index)) {
            if (diff.status === 'missing') {
              summarySheet.getRange(summarySheet.getLastRow() + 1,appendData.length).setBackground('yellow');
            } else if (mainValue !== compareValue) {
              summarySheet.getRange(summarySheet.getLastRow() + 1, index + 4).setBackground('yellow');
            }
          }
        });
        summarySheet.appendRow(appendData);
      });
    }

    // Alert the user about the results
    var message = (differences.length === 0) ? 'No differences found.' : 'Comparison complete. ';
    if (action === 'highlight') {
      message += 'Differences have been highlighted in the comparison sheets.';
    } else if (action === 'summary') {
      message += 'Check the "Comparison Summary" sheet for details.';
    }
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    if (error.message === 'Process cancelled by user') {
      SpreadsheetApp.getUi().alert('The process was cancelled.');
    } else {
      SpreadsheetApp.getUi().alert('An error occurred: ' + error.message);
    }
  }
}

function cancelProcess() {
  PropertiesService.getScriptProperties().setProperty('CANCEL_PROCESS', 'true');
}
