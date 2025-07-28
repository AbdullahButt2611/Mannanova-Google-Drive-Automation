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
      "REG-602 IBC Log and Inventory": { fileId: "16ONekxAeaADr9wyJ96aHOmaT7oRO89_9SPgH5_b9rUE", sheetName: "All IBCs - List" },
    };

    let currentFileName = null;
    let currentDataBlock = [];
    let skipNextRow = false;
    let currentBlockStartRow = null;

    const fileDataMap = {};  // { fileName: [ [row1], [row2] ] }
    const columnCountMap = {};  // { fileName: columnCount }
    const blockRowMap = {};           // { fileName: rowIndex }

    for (let i = 0; i < data.length; i++) {
      const firstCell = String(data[i][0]).trim().toLowerCase();
      const secondCell = String(data[i][1]).trim();

      // Start of a block only if checkbox is TRUE
      if (firstCell === 'true') {
        // Save previous block if exists
        if (currentFileName && currentDataBlock.length > 0) {
          if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
          fileDataMap[currentFileName].push(...currentDataBlock);
          console.log(`Saved data block for file: ${currentFileName} with ${currentDataBlock.length} rows.`);
        }

        currentFileName = secondCell;
        currentBlockStartRow = i + 1;
        blockRowMap[currentFileName] = currentBlockStartRow;
        currentDataBlock = [];
        skipNextRow = true;
        continue;
      }

      // Skip blocks explicitly marked as false (just skip this row)
      if (firstCell === 'false') {
        console.log(`Skipping block '${secondCell}' as it is marked false.`);
        continue;
      }

      // Handle header row (right after TRUE)
      if (skipNextRow) {
        const headerRow = data[i].slice(1);
        const colCount = headerRow.filter(cell => String(cell).trim() !== "").length;
        columnCountMap[currentFileName] = colCount;
        skipNextRow = false;
        console.log(`Header for block ${currentFileName}: ${JSON.stringify(headerRow)} (${colCount} columns)`);
        continue;
      }

      // Add data rows if current block is valid
      if (currentFileName && firstCell === '') {
        const colCount = columnCountMap[currentFileName] || 0;
        const rowData = data[i].slice(1, 1 + colCount);
        const hasEmpty = rowData.some(cell => String(cell).trim() === '');

        if (!hasEmpty) {
          currentDataBlock.push(rowData);
          console.log(`Added row to block ${currentFileName}: ${JSON.stringify(rowData)}`);
        } else {
          console.log(`Skipped row with empty cells in block ${currentFileName}: ${JSON.stringify(rowData)}`);
        }
      }
    }

    // Save last block
    if (currentFileName && currentDataBlock.length > 0) {
      if (!fileDataMap[currentFileName]) fileDataMap[currentFileName] = [];
      fileDataMap[currentFileName].push(...currentDataBlock);
      console.log(`Saved data block for file: ${currentFileName} with ${currentDataBlock.length} rows.`);
    }

    // After all blocks are processed, check if there is any data to push
    if (Object.keys(fileDataMap).length === 0) {
      ui.alert("Notice", "This file is either already parsed or there is nothing new to copy.", ui.ButtonSet.OK);
      console.log("No new data found to copy.");
      return;
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
      const colCount = columnCountMap[fileName] || (rows.length > 0 ? rows[0].length : 0);
      const lastRow = destinationSheet.getLastRow();
      console.log(`Appending ${rows.length} rows to '${mapping.sheetName}' in file '${fileName}' (${colCount} columns).`);
      destinationSheet.getRange(lastRow + 1, 1, rows.length, colCount).setValues(rows);

      // Skip formula copy for REG-102 RRIBC to avoid row duplication
      if (fileName !== "REG-102 RRIBC") {
        // Copy formulas column-wise instead of row-wise
        const totalCols = destinationSheet.getLastColumn();
        if (totalCols > colCount) {
          const formulaCols = totalCols - colCount;

          for (let col = colCount + 1; col <= totalCols; col++) {
            // Find a non-empty formula cell in this column
            const columnFormulas = destinationSheet.getRange(1, col, lastRow).getFormulas();
            let baseFormula = null;

            for (let row = lastRow - 1; row >= 0; row--) {
              const f = columnFormulas[row][0];
              if (f && f.trim() !== '') {
                baseFormula = f;
                break;
              }
            }

            if (baseFormula) {
              // Prepare an array to apply the formula to all new rows
              const formulaArray = Array(rows.length).fill([baseFormula]);
              destinationSheet
                .getRange(lastRow + 1, col, rows.length, 1)
                .setFormulas(formulaArray);
              console.log(`Applied formula in column ${col}: ${baseFormula}`);
            } else {
              console.warn(`No formula found in column ${col} to copy.`);
            }
          }
        }
      } else {
        // Custom logic for REG-102: copy only basic formulas, skip ARRAYFORMULA or complex ones
        const totalCols = destinationSheet.getLastColumn();
        if (totalCols > colCount) {
          for (let col = colCount + 1; col <= totalCols; col++) {
            const columnFormulas = destinationSheet.getRange(1, col, lastRow).getFormulas();
            let baseFormula = null;

            for (let row = lastRow - 1; row >= 0; row--) {
              const f = columnFormulas[row][0];
              if (
                f &&
                f.trim() !== "" &&
                !f.toLowerCase().includes("arrayformula") &&
                !f.toLowerCase().includes("query") &&
                !f.toLowerCase().includes("filter")
              ) {
                baseFormula = f;
                break;
              }
            }

            if (baseFormula) {
              const formulaArray = Array(rows.length).fill([baseFormula]);
              destinationSheet
                .getRange(lastRow + 1, col, rows.length, 1)
                .setFormulas(formulaArray);
              console.log(`[REG-102] Applied basic formula in col ${col}: ${baseFormula}`);
            } else {
              console.log(`[REG-102] Skipped formula in col ${col} (complex or empty).`);
            }
          }
        }
      }

      // Update Pastepad block header to "true"
      const blockRow = blockRowMap[fileName];
      if (blockRow) {
        sourceSheet.getRange(blockRow, 1).setValue(false);
        console.log(`Marked block '${fileName}' as processed at row ${blockRow}`);
      }
    }

    console.log("File Data Map", fileDataMap)

    ui.alert("Success", "Data has been appended to the destination files successfully.", ui.ButtonSet.OK);
    console.log("Data append operation completed for all destination files.");    

  } catch (error) {
    console.error("Error in runDataExtraction:", error);
    SpreadsheetApp.getUi().alert("Error", `An error occurred:\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}