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

    const data = sourceSheet.getDataRange().getValues();
    if (!data || data.length === 1) throw new Error("No data found in 'Pastepad'.");

    // 🔁 Map of file names to their Google Sheet IDs
    const fileMap = {
      "REG-100": "1Hb-agA100ZwuvPf9lAlHk-s5h5HG5fVWLePO7EYBzbY",
      // Add more mappings as needed
    };

    // 🧠 State Tracking
    let currentFileName = null;
    let currentDataBlock = [];
    let skipNextRow = false;

    const fileDataMap = {}; // { fileName: [ [row1], [row2] ] }

    for (let i = 0; i < data.length; i++) {
      const firstCell = String(data[i][0]).trim().toLowerCase();
      const secondCell = String(data[i][1]).trim();

      const isNewBlock = (firstCell === 'true' || firstCell === 'false');

      if (isNewBlock) {
        // Save previous block
        if (currentFileName && currentDataBlock.length > 0) {
          if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
          fileDataMap[currentFileName].push(...currentDataBlock);
        }

        currentFileName = secondCell;
        currentDataBlock = [];
        skipNextRow = true; // 🚫 Skip the next row (header)
        continue;
      }

      // 🚫 Skip next row after block start
      if (skipNextRow) {
        skipNextRow = false;
        continue;
      }

      // ✅ Only process content rows inside a block
      if (currentFileName && firstCell === '') {
        const rowData = data[i].slice(1); // exclude first column (A)

        const hasEmptyCell = rowData.some(cell => String(cell).trim() === '');
        if (!hasEmptyCell) {
          currentDataBlock.push(rowData);
        }
      }
    }

    // Save last block
    if (currentFileName && currentDataBlock.length > 0) {
      if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
      fileDataMap[currentFileName].push(...currentDataBlock);
    }

    // 📤 Push to destination files
    for (const [fileName, rows] of Object.entries(fileDataMap)) {
      const fileId = fileMap[fileName];
      if (!fileId) {
        console.warn(`No file ID mapped for '${fileName}'. Skipping.`);
        continue;
      }

      const destinationSpreadsheet = SpreadsheetApp.openById(fileId);
      const destinationSheet = destinationSpreadsheet.getSheetByName("Inventaire");

      if (!destinationSheet) {
        console.warn(`Sheet 'Inventaire' not found in file '${fileName}'. Skipping.`);
        continue;
      }

      const lastRow = destinationSheet.getLastRow();
      destinationSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    ui.alert("Success", "Data has been appended to the destination files successfully.", ui.ButtonSet.OK);
    
  } catch (error) {
    console.error("Error in runDataExtraction:", error);
    SpreadsheetApp.getUi().alert("Error", `An error occurred:\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}