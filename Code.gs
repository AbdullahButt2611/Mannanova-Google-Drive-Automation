/**
 * Creates custom menu when spreadsheet opens
 * This function is automatically triggered when the spreadsheet is opened
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Creates and adds custom menu to the spreadsheet UI
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Data Automation')
    .addItem('Run Script', 'runDataExtraction')
    .addToUi();
}

/**
 * Main automation function - currently prints test message
 * This will be expanded to handle data extraction logic
 */
function runDataExtraction() {
  try {
    console.log('Data Automation Started Successfully');
    console.log('Execution Time:', new Date().toLocaleString());
    console.log('Current Sheet:', SpreadsheetApp.getActiveSheet().getName());
    
    const ui = SpreadsheetApp.getUi();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pastepad");
    if (!sourceSheet) throw new Error("Sheet 'Pastepad' not found.");
    console.log("Source sheet 'Pastepad' found.");

    const data = sourceSheet.getDataRange().getValues();
    if (!data || data.length === 0) throw new Error("No data found in 'Pastepad'.");
    console.log(`Fetched ${data.length} rows from 'Pastepad'.`);

    const fileMap = {
      "REG-101": { fileId: "1Hb-agA100ZwuvPf9lAlHk-s5h5HG5fVWLePO7EYBzbY", sheetName: "Inventaire" },
      "REG-102 RRIBC": { fileId: "19Ew_W50zatGad97FCpfgzAptkSfrJExF3cnDjDY8ZEc", sheetName: "IBC Use and Clean" },
      // Add more mappings as needed
    };

    let currentFileName = null;
    let currentDataBlock = [];
    let skipNextRow = false;
    let blockColCount = 0;
    let headerRow = [];

    const fileDataMap = {}; // { fileName: [ [row1], [row2] ] }

    console.log("Data Length", data.length)
    for (let i = 0; i < data.length; i++) {
    const firstCell = String(data[i][0]).trim().toLowerCase();
    const secondCell = String(data[i][1]).trim();

    const isNewBlock = (firstCell === 'true' || firstCell === 'false');

    if (isNewBlock) {
      // Save previous block
      if (currentFileName && currentDataBlock.length > 0) {
        if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
        fileDataMap[currentFileName].push(...currentDataBlock);
        console.log(`Saved data block for file: ${currentFileName} with ${currentDataBlock.length} rows.`);
      }

      currentFileName = secondCell;
      currentDataBlock = [];
      skipNextRow = true; // Next row is header
      blockColCount = 0;
      headerRow = [];
      console.log(`Started new data block for file: ${currentFileName}`);
      continue;
    }

    if (skipNextRow) {
      headerRow = data[i].slice(1); // exclude first column (A)
      blockColCount = headerRow.length;
      skipNextRow = false;
      console.log(`Header for block ${currentFileName}: ${JSON.stringify(headerRow)} (columns: ${blockColCount})`);
      continue;
    }

    if (currentFileName && firstCell === '') {
      const rowData = data[i].slice(1, 1 + blockColCount); // only take as many columns as header
      const allEmpty = rowData.every(cell => String(cell).trim() === '');
      if (!allEmpty) {
        currentDataBlock.push(rowData);
        console.log(`Added row to block ${currentFileName}: ${JSON.stringify(rowData)}`);
      } else {
        console.log(`Skipped completely empty row in block ${currentFileName}: ${JSON.stringify(rowData)}`);
      }
    }
  }

  // Save the last block after the loop
  if (currentFileName && currentDataBlock.length > 0) {
    if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
    fileDataMap[currentFileName].push(...currentDataBlock);
    console.log(`Saved data block for file: ${currentFileName} with ${currentDataBlock.length} rows.`);
  }

    // Save last block
    if (currentFileName && currentDataBlock.length > 0) {
      if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
      fileDataMap[currentFileName].push(...currentDataBlock);
      console.log(`Saved data block for file: ${currentFileName} with ${currentDataBlock.length} rows.`);
    }

    // Push to destination files
    for (const [fileName, rows] of Object.entries(fileDataMap)) {
      const mapping = fileMap[fileName];
      if (!mapping || !mapping.fileId || !mapping.sheetName) {
        console.warn(`Missing mapping info for '${fileName}'. Skipping.`);
        continue;
      }

      const destinationSpreadsheet = SpreadsheetApp.openById(mapping.fileId);
      const destinationSheet = destinationSpreadsheet.getSheetByName(mapping.sheetName);

      if (!destinationSheet) {
        console.warn(`Sheet '${mapping.sheetName}' not found in file '${fileName}'. Skipping.`);
        continue;
      }

      // Use the header length for this block
      const colCount = rows.length > 0 ? rows[0].length : 0;
      const lastRow = destinationSheet.getLastRow();
      console.log(`Appending ${rows.length} rows to '${mapping.sheetName}' in file '${fileName}' (${colCount} columns).`);
      destinationSheet.getRange(lastRow + 1, 1, rows.length, colCount).setValues(rows);
    }

    console.log("File Data Map", fileDataMap)

    ui.alert("Success", "Data has been appended to the destination files successfully.", ui.ButtonSet.OK);
    console.log("Data append operation completed for all destination files.");    

  } catch (error) {
    console.error("Error in runDataExtraction:", error);
    SpreadsheetApp.getUi().alert("Error", `An error occurred:\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}