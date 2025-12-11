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
    if (mainValues.length === 0) {
      throw new Error('Main sheet is empty.');
    }
  } else if (comparationOption === 'formulas') {
    mainFormulas = mainRange.getFormulas();
    if (mainFormulas.length === 0) {
      throw new Error('Main sheet is empty.');
    }
  } else if (comparationOption === 'valuesAndFormulas') {
    mainValues = mainRange.getValues();
    mainFormulas = mainRange.getFormulas();
    if (mainValues.length === 0 && mainFormulas.length === 0) {
      throw new Error('Main sheet is empty.');
    }
  } else {
    throw new Error('Invalid comparison option: ' + comparationOption);
  }

  var columnIndices = [];

  if (columnOption === 'all') {
    columnIndices = [];
    if (comparationOption === 'formulas') {
      columnIndices = Array.from({ length: mainFormulas[0].length }, (v, k) => k);
    } else if (comparationOption === 'values') {
      columnIndices = Array.from({ length: mainValues[0].length }, (v, k) => k);
    } else if (comparationOption === 'valuesAndFormulas') {
      var maxLength = Math.max(mainValues[0].length, mainFormulas[0].length);
      columnIndices = Array.from({ length: maxLength }, (v, k) => k);
    }
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
      Logger.log("compare sheet :"+sheetName + " mainSheet :" + mainSheetName)
      if (!sheet) {
        throw new Error('Sheet not found: ' + sheetName);
      }
      compareSheetWithMain(sheet, mainSheet, mainValues, mainFormulas, columnIndices, action, differences, comparationOption);
    });

    // Compare sheets from the provided URL spreadsheet
    if (diffSpreadsheetUrl && diffSelectedSheets.length > 0) {
      Logger.log("diffSpreadsheetUrl")
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
    Logger.log("ERROR:" + error.message)
    if (error.message === 'Process cancelled by user') {
      SpreadsheetApp.getUi().alert('The process was cancelled.');
    } else {
      SpreadsheetApp.getUi().alert('An error occurred: ' + error.message);
    }
  }
}

// Helper function to compare a sheet with the main sheet and highlight differences
function compareSheetWithMain(sheet, mainSheet, mainValues, mainFormulas, columnIndices, action, differences, comparationOption) {
  Logger.log({
    "comparationOption": comparationOption,
    "mainSheet": sheet.getName(),
  });

  var range = sheet.getDataRange();
  var values, formulas;

  if (comparationOption === 'values' || comparationOption === 'valuesAndFormulas') {
    values = range.getValues();
  }
  if (comparationOption === 'formulas' || comparationOption === 'valuesAndFormulas') {
    formulas = range.getFormulas();
  }

  var maxRows = 0;
    if (mainValues && mainValues.length > 0) {
    maxRows = Math.max(maxRows, mainValues.length);
  }

  if (mainFormulas && mainFormulas.length > 0) {
    maxRows = Math.max(maxRows, mainFormulas.length);
  }

  var maxRowsDiff = 0;
  if (values && values.length > 0) {
    maxRowsDiff = Math.max(maxRowsDiff, values.length);
  }
  if (formulas && formulas.length > 0) {
    maxRowsDiff = Math.max(maxRowsDiff, formulas.length);
  }

  // Compare the two ranges
  for (var row = 0; row < Math.max(maxRows, maxRowsDiff); row++) {
    if (row >= (mainValues ? mainValues.length : 0) || row >= (values ? values.length : 0)) {
      Logger.log('Skipping row ' + row + ' as it is out of bounds.');
      continue;
    }
    var mainRowExists = row < (mainValues ? mainValues.length : 0);
    var compareRowExists = row < (values ? values.length : 0);
    
    columnIndices.forEach(function(col) {
      if (col >= (mainValues ? mainValues[0].length : 0) || col >= (values ? values[0].length : 0)) {
        Logger.log('Skipping column ' + col + ' as it is out of bounds.');
        return;
      }
      
      var mainValue = mainRowExists && mainValues[row] && col < mainValues[row].length ? mainValues[row][col] : null;
      var compareValue = compareRowExists && values[row] && col < values[row].length ? values[row][col] : null;
      var mainFormula = mainRowExists && mainFormulas && mainFormulas[row] && col < mainFormulas[row].length ? mainFormulas[row][col] : null;
      var compareFormula = compareRowExists && formulas && formulas[row] && col < formulas[row].length ? formulas[row][col] : null;

      // Skip if both mainValue and compareValue are null
      if (mainValue === null && compareValue === null && mainFormula === null && compareFormula === null) {
        return;
      }

      // Convert dates to strings in a specific format for comparison
      if (mainValue instanceof Date && compareValue instanceof Date) {
        mainValue = Utilities.formatDate(mainValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        compareValue = Utilities.formatDate(compareValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      var valueDifference = (mainValue !== compareValue) && !(mainValue == null && compareValue === "") && !(mainValue === "" && compareValue == null);
      var formulaDifference = (mainFormula !== compareFormula) && !(mainFormula == null && compareFormula === "") && !(mainFormula === "" && compareFormula == null);

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
          masterData: mainValues ? mainValues[row][col] : null,
          data: values ? values[row][col] : null,
          masterFormulas: mainFormulas ? mainFormulas[row][col] : null,
          formulas: formulas ? formulas[row][col] : null
        });
      }
    });
  }

  if (comparationOption === 'values' || comparationOption === 'valuesAndFormulas') {
    // Highlight any additional data below the range of the main sheet
    if (values && values.length > maxRows) {
      for (var row = maxRows; row < values.length; row++) {
        // Ensure the row exists and is not undefined
        if (!values[row]) {
          Logger.log('Skipping row ' + row + ' because it is undefined.');
          continue;
        }

        columnIndices.forEach(function(col) {
          if (col >= (values ? values[0].length : 0)) {
            Logger.log('Skipping column ' + col + ' as it is out of bounds.');
            return;
          }

          if (action !== 'summary') {
            mainSheet.getRange(row + 1, col + 1).setBackground('yellow');
          }

          differences.push({
            sheet: sheet.getName(),
            row: row + 1,
            col: col + 1,
            status: 'missing data',
            masterData: mainValues && mainValues[row] ? mainValues[row][col] : null,
            data: values && values[row] ? values[row][col] : null,
            masterFormulas: mainFormulas && mainFormulas[row] ? mainFormulas[row][col] : null,
            formulas: formulas && formulas[row] ? formulas[row][col] : null
          });
        })
      }
    }
  }
  if (comparationOption === 'formulas' || comparationOption === 'valuesAndFormulas') {
    // Highlight any additional data below the range of the main sheet
    if (formulas && formulas.length > maxRows) {
      for (var row = maxRows; row < formulas.length; row++) {
        // Ensure the row exists and is not undefined
        if (!formulas[row]) {
          Logger.log('Skipping row ' + row + ' because it is undefined.');
          continue;
        }

        columnIndices.forEach(function(col) {
          if (col >= (formulas ? formulas[0].length : 0)) {
            Logger.log('Skipping column ' + col + ' as it is out of bounds.');
            return;
          }

          if (action !== 'summary') {
            mainSheet.getRange(row + 1, col + 1).setBackground('yellow');
          }

          differences.push({
            sheet: sheet.getName(),
            row: row + 1,
            col: col + 1,
            status: 'missing formula',
            masterData: mainValues && mainValues[row] ? mainValues[row][col] : null,
            data: values && values[row] ? values[row][col] : null,
            masterFormulas: mainFormulas && mainFormulas[row] ? mainFormulas[row][col] : null,
            formulas: formulas && formulas[row] ? formulas[row][col] : null
          });
        })
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
      diff.data !== null && diff.status != 'different formula' ? diff.data : '',
      diff.masterFormulas !== null && diff.status != 'different value' & diff.status != 'missing data' ? "'" + diff.masterFormulas : '',
      diff.formulas !== null && diff.status != 'different value' & diff.status != 'missing data' ? "'" + diff.formulas : ''
    ];
    
    summarySheet.appendRow(appendData);
  });
}

function cancelProcess() {
  PropertiesService.getScriptProperties().setProperty('CANCEL_PROCESS', 'true');
}
