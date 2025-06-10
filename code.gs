function onInstall(e){
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('NOC - Compare Sheets', 'compareMultipleSheets')
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
  html.setTitle('Compare Sheets');
  html.setContent(html.getContent().replace('<!--SHEET_NAMES-->', JSON.stringify(sheetNames)));
  SpreadsheetApp.getUi().showModalDialog(html, 'Compare Sheets');
}

function fetchSheetNamesFromUrl(url) {
  try {
    var ss = SpreadsheetApp.openByUrl(url);
    var sheets = ss.getSheets();
    var sheetNames = sheets.map(function(sheet) {
      return sheet.getName();
    });
    return sheetNames; // Return the sheet names to the client
  } catch (error) {
    throw new Error('Failed to fetch sheet names: ' + error.message);
  }
}

function columnLetterToIndex(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column - 1;
}

function indexToColumnLetter(index) {
  let letter = '';
  let column = index + 1; // Convert to 1-based index
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter; // 65 is the char code for 'A'
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function processComparison(mainSheetName, selectedSheets, action, columnOption, startColumn, endColumn, specificColumns, diffSpreadsheetUrl, diffSelectedSheets, comparationOption) {
  if (!mainSheetName) {
    throw new Error('Main sheet name is not provided.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Active spreadsheet
  var mainSheet = ss.getSheetByName(mainSheetName);

  if (!mainSheet) {
    throw new Error('Main sheet not found: ' + mainSheetName);
  }

  var mainRange = mainSheet.getDataRange();
  var mainValues, mainFormulas;

  if (comparationOption === 'values') {
    mainValues = mainRange.getValues();
  } else if (comparationOption === 'formulas') {
    mainFormulas = mainRange.getFormulas();
  } else if (comparationOption === 'valuesAndFormulas') {
    mainValues = mainRange.getValues();
    mainFormulas = mainRange.getFormulas();
  } else {
    throw new Error('Invalid comparison option: ' + comparationOption);
  }

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
    // Compare sheets from the active spreadsheet
    selectedSheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error('Sheet not found: ' + sheetName);
      }
      compareSheetWithMain(sheet, mainSheet, mainValues, mainFormulas, columnIndices, action, differences, comparationOption);
    });

    // Compare sheets from the provided URL spreadsheet
    if (diffSpreadsheetUrl && diffSelectedSheets.length > 0) {
      var diffSpreadsheet = SpreadsheetApp.openByUrl(diffSpreadsheetUrl);
      diffSelectedSheets.forEach(function(sheetName) {
        var sheet = diffSpreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          throw new Error('Sheet not found in the provided spreadsheet: ' + sheetName);
        }
        compareSheetWithMain(sheet, mainSheet, mainValues, mainFormulas, columnIndices, action, differences, comparationOption);
      });
    }

    if (action === 'summary') {
      createSummarySheet(ss, differences, columnIndices);
    }

    // Alert the user about the results
    var message = (differences.length === 0) ? 'No differences found.' : 'Comparison complete. ';
    if (action === 'highlight') {
      message += 'Differences have been highlighted in the main sheet.';
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

// Helper function to compare a sheet with the main sheet and highlight differences
function compareSheetWithMain(sheet, mainSheet, mainValues, mainFormulas, columnIndices, action, differences, comparationOption) {
  var range = sheet.getDataRange();
  var values, formulas;

  if (comparationOption === 'values' || comparationOption === 'valuesAndFormulas') {
    values = range.getValues();
  }
  if (comparationOption === 'formulas' || comparationOption === 'valuesAndFormulas') {
    formulas = range.getFormulas();
  }

  var maxRows = mainValues.length;

  // Compare the two ranges
  for (var row = 0; row < maxRows && row < values.length; row++) {
    var rowHasDifference = false;
    columnIndices.forEach(function(col) {
      var mainValue = mainValues[row][col];
      var compareValue = values ? values[row][col] : null;
      var mainFormula = mainFormulas[row][col];
      var compareFormula = formulas ? formulas[row][col] : null;

      // Convert dates to strings in a specific format for comparison
      if (mainValue instanceof Date && compareValue instanceof Date) {
        mainValue = Utilities.formatDate(mainValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        compareValue = Utilities.formatDate(compareValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      var valueDifference = mainValue !== compareValue;
      var formulaDifference = mainFormula !== compareFormula;

      if (valueDifference || formulaDifference) {
        Logger.log({
          "main value": mainValue,
          "compare value": compareValue,
          "main formula": mainFormula,
          "compare formula": compareFormula,
          "row": row,
          "col": col
        });

        if (action !== 'summary') {
          var cell = mainSheet.getRange(row + 1, col + 1);
          if (valueDifference && formulaDifference) {
            cell.setBackground('green'); // Highlight for both value and formula differences
          } else if (valueDifference) {
            cell.setBackground('yellow'); // Highlight for value differences
          } else if (formulaDifference) {
            cell.setBackground('blue'); // Highlight for formula differences
          }
        }

        var status = 'different';

        if (valueDifference && formulaDifference) {
          status = 'different value and formula';
        } else if (valueDifference) {
          status = 'different value';
        } else if (formulaDifference) {
          status = 'different formula';
        }

        differences.push({
          sheet: sheet.getName(),
          row: row + 1,
          col: col + 1,
          status: status,
          masterData: mainValues[row][col] || null,
          data: values[row][col] || null,
          masterFormulas: mainFormulas[row][col] || null,
          formulas: formulas[row][col] || null
        });
      }
    });
  }

  // Highlight any additional data below the range of the main sheet
  if (values && values.length > maxRows) {
    for (var row = maxRows; row < values.length; row++) {
      for (var col = 0; col < values[row].length; col++) {
        if (action !== 'summary') {
          mainSheet.getRange(row + 1, col + 1).setBackground('yellow');
        }
        differences.push({
          sheet: sheet.getName(),
          row: row + 1,
          col: col + 1,
          status: 'missing data',
          masterData: mainValues[row][col] || null,
          data: values[row][col] || null,
          masterFormulas: mainFormulas[row][col] || null,
          formulas: formulas[row][col] || null,

        });
      }
    }
  }
}

// Helper function to create a summary sheet
function createSummarySheet(ss, differences, columnIndices) {
  var summarySheet = ss.getSheetByName('Comparison Summary') || ss.insertSheet('Comparison Summary');
  summarySheet.clear();
  summarySheet.appendRow(['Sheet', 'Status', 'Cell', 'Master Value', 'Compared Value', 'Master Formula', 'Compared Formula']);

  differences.forEach(function(diff) {
    var appendData = [
      diff.sheet,
      diff.status,
      indexToColumnLetter(diff.col - 1) + diff.row, // Convert column index to letter
      diff.masterData !== null && diff.status != 'different formula' ? diff.masterData : '',
      diff.data !== null && diff.status != 'different formula'? diff.data : '',
      diff.masterFormulas !== null && diff.status != 'different value' ? "'" + diff.masterFormulas : '',
      diff.formulas !== null && diff.status != 'different value' ? "'" + diff.formulas : ''
    ];
    
    summarySheet.appendRow(appendData);
  });
}

function cancelProcess() {
  PropertiesService.getScriptProperties().setProperty('CANCEL_PROCESS', 'true');
}
