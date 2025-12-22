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
    comparationOption: comparationOption,
    mainSheet: mainSheet.getName(),
    sheet: sheet.getName()
  });

  const range = sheet.getDataRange();
  const values = (comparationOption === 'values' || comparationOption === 'valuesAndFormulas') ? range.getValues() : null;
  const formulas = (comparationOption === 'formulas' || comparationOption === 'valuesAndFormulas') ? range.getFormulas() : null;

  const maxRows = Math.max(mainValues?.length || 0, mainFormulas?.length || 0);
  const maxRowsDiff = Math.max(values?.length || 0, formulas?.length || 0);

  Logger.log({ maxRows, maxRowsDiff });

  // Helper function to handle highlighting and logging differences
  function handleDifference(row, col, status, mainValue, compareValue, mainFormula, compareFormula) {
    if (action !== 'summary') {
      const cell = mainSheet.getRange(row + 1, col + 1);
      if (status === 'different value and formula') {
        cell.setBackground('green');
      } else if (status === 'different value') {
        cell.setBackground('yellow');
      } else if (status === 'different formula') {
        cell.setBackground('blue');
      }
    }

    differences.push({
      sheet: sheet.getName(),
      row: row + 1,
      col: col + 1,
      status,
      masterData: mainValue,
      data: compareValue,
      masterFormulas: mainFormula,
      formulas: compareFormula,
    });
  }

  // Compare rows and columns
  for (let row = 0; row < Math.max(maxRows, maxRowsDiff); row++) {
    const mainRowExists = row < (mainValues?.length || 0);
    const compareRowExists = row < (values?.length || 0);

    if (!mainRowExists && !compareRowExists) continue;

    columnIndices.forEach((col) => {
      let mainValue = mainRowExists ? mainValues[row]?.[col] : null;
      let compareValue = compareRowExists ? values[row]?.[col] : null;
      let mainFormula = mainRowExists ? mainFormulas?.[row]?.[col] : null;
      let compareFormula = compareRowExists ? formulas?.[row]?.[col] : null;

      // Skip if both values and formulas are null
      if (!mainValue && !compareValue && !mainFormula && !compareFormula) return;

      // Convert dates to strings for comparison
      if (mainValue instanceof Date && compareValue instanceof Date) {
        mainValue = Utilities.formatDate(mainValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        compareValue = Utilities.formatDate(compareValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      const valueDifference = mainValue !== compareValue && !(mainValue == null && compareValue === "") && !(mainValue === "" && compareValue == null);
      const formulaDifference = mainFormula !== compareFormula && !(mainFormula == null && compareFormula === "") && !(mainFormula === "" && compareFormula == null);

      if (valueDifference || formulaDifference) {
        const status = valueDifference && formulaDifference
          ? 'different value and formula'
          : valueDifference
          ? 'different value'
          : 'different formula';

        handleDifference(row, col, status, mainValue, compareValue, mainFormula, compareFormula);
      }
    });
  }

  // Handle extra rows in the comparison sheet
  function highlightExtraRows(extraData, startRow, status) {
    for (let row = startRow; row < extraData.length; row++) {
      if (!extraData[row]) continue;

      columnIndices.forEach((col) => {
        if (col >= extraData[0].length) return;

        if (action !== 'summary') {
          mainSheet.getRange(row + 1, col + 1).setBackground('yellow');
        }

        differences.push({
          sheet: sheet.getName(),
          row: row + 1,
          col: col + 1,
          status,
          masterData: mainValues?.[row]?.[col] || null,
          data: values?.[row]?.[col] || null,
          masterFormulas: mainFormulas?.[row]?.[col] || null,
          formulas: formulas?.[row]?.[col] || null,
        });
      });
    }
  }

  if (comparationOption === 'values' || comparationOption === 'valuesAndFormulas') {
    if (values?.length > maxRows) highlightExtraRows(values, maxRows, 'missing data');
  }

  if (comparationOption === 'formulas' || comparationOption === 'valuesAndFormulas') {
    if (formulas?.length > maxRows) highlightExtraRows(formulas, maxRows, 'missing formula');
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
