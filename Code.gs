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

    let reg101Initials = null; // Store initials for REG-101
    let reg102Dates = []; // Store dates for each row of REG-102

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
        const statusColIndex = 2;  // Column C (0-indexed: A=0, B=1, C=2)
        const statusCellValue = String(data[i][statusColIndex]).trim().toLowerCase();

        if (statusCellValue === 'processed') {
          console.log(`Skipping block '${secondCell}' — already marked as processed.`);
          continue;
        }

        // Check if this is REG-101 and capture initials from column E (index 4)
        if (secondCell === 'REG-101' && i > 0) {
          // Get initials from the previous row, column E (index 4)
          reg101Initials = String(data[i - 1][4]).trim();
          console.log(`Captured initials for REG-101: ${reg101Initials}`);
        }

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

        // Reset REG-102 dates when starting a new block
        if (secondCell === 'REG-102 RRIBC') {
          reg102Dates = [];
          console.log('Reset dates array for REG-102 RRIBC');
        }

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
        let colCount = headerRow.filter(cell => String(cell).trim() !== "").length;
        
        // For REG-102, exclude the Date column from main column count
        if (currentFileName === 'REG-102 RRIBC') {
          // Assuming Date is the last column in the header
          const dateColumnIndex = colCount - 1;
          if (headerRow[dateColumnIndex] && String(headerRow[dateColumnIndex]).trim().toLowerCase() === 'date') {
            colCount = colCount - 1; // Exclude Date column from main data
            console.log(`REG-102 RRIBC: Found Date column at index ${dateColumnIndex}, adjusted column count to ${colCount}`);
          }
        }
        
        columnCountMap[currentFileName] = colCount;
        skipNextRow = false;
        console.log(`Header for block ${currentFileName}: ${JSON.stringify(headerRow)} (${colCount} columns)`);
        continue;
      }

      // Add data rows if current block is valid
      if (currentFileName && firstCell === '') {
        const colCount = columnCountMap[currentFileName] || 0;
        const rowData = data[i].slice(1, 1 + colCount);
        
        // For REG-102, also capture the date from the next column
        if (currentFileName === 'REG-102 RRIBC') {
          const dateValue = data[i][1 + colCount]; // Date is after the main columns
          const hasEmpty = rowData.some(cell => String(cell).trim() === '');
          
          if (!hasEmpty) {
            currentDataBlock.push(rowData);
            reg102Dates.push(dateValue); // Store the date for this row
            console.log(`Added row to REG-102 with date: ${JSON.stringify(rowData)}, Date: ${dateValue}`);
          } else {
            console.log(`Skipped row with empty cells in block ${currentFileName}: ${JSON.stringify(rowData)}`);
          }
        } else {
          const hasEmpty = rowData.some(cell => String(cell).trim() === '');
          if (!hasEmpty) {
            currentDataBlock.push(rowData);
            console.log(`Added row to block ${currentFileName}: ${JSON.stringify(rowData)}`);
          } else {
            console.log(`Skipped row with empty cells in block ${currentFileName}: ${JSON.stringify(rowData)}`);
          }
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

      // Section - Column Count Logic
      // Determine actual column count from the data rows
      const actualColCount = rows.length > 0 ? rows[0].length : columnCountMap[fileName] || 0;

      const lastRow = destinationSheet.getLastRow();
      console.log(`Appending ${rows.length} rows to '${mapping.sheetName}' in file '${fileName}' (${actualColCount} columns).`);

      // For REG-101, write data without initials (only the original 6 columns)
      if (fileName === 'REG-101' && reg101Initials) {
        // Extract only the first 6 columns (without the gap and initials)
        const dataWithoutInitials = rows.map(row => row.slice(0, 6));
        destinationSheet.getRange(lastRow + 1, 1, rows.length, 6).setValues(dataWithoutInitials);
        
        // Explicitly write initials to column K
        const initialsColumn = 11; // Column K (A=1, B=2, ... K=11)
        for (let i = 0; i < rows.length; i++) {
          destinationSheet.getRange(lastRow + 1 + i, initialsColumn).setValue(reg101Initials);
        }
        console.log(`Set initials '${reg101Initials}' in column K for ${rows.length} rows`);
      } 
      // For REG-102, write data and add dates to column H
      else if (fileName === 'REG-102 RRIBC' && reg102Dates.length > 0) {
        destinationSheet.getRange(lastRow + 1, 1, rows.length, actualColCount).setValues(rows);
        
        const dateColumn = 8;      // Column H (Date)
        const initialsColumn = 9;  // Column I (Initials)

        for (let i = 0; i < reg102Dates.length && i < rows.length; i++) {
          // Write Date
          if (reg102Dates[i]) {
            destinationSheet.getRange(lastRow + 1 + i, dateColumn).setValue(reg102Dates[i]);
          }
          // Write Initials (copied same way from Pastepad like in REG-101)
          if (reg101Initials) {
            destinationSheet.getRange(lastRow + 1 + i, initialsColumn).setValue(reg101Initials);
          }
        }

        console.log(`Set dates in column H and initials '${reg101Initials}' in column I for ${reg102Dates.length} rows in REG-102 RRIBC`);
      } 
      // For other files, write normally
      else {
        destinationSheet.getRange(lastRow + 1, 1, rows.length, actualColCount).setValues(rows);
      }
      
      // Copy formatting logic 
      const sourceRange = destinationSheet.getRange(lastRow, 1, 1, destinationSheet.getLastColumn());
      const targetRange = destinationSheet.getRange(lastRow + 1, 1, rows.length, destinationSheet.getLastColumn());
      for (let i = 0; i < rows.length; i++) {
        sourceRange.copyTo(destinationSheet.getRange(lastRow + 1 + i, 1, 1, destinationSheet.getLastColumn()), { formatOnly: true });
      }

      const templateFormulas = [];
      const lastRowFormulas = destinationSheet.getRange(lastRow, 1, 1, destinationSheet.getLastColumn()).getFormulas()[0];

      for (let i = 0; i < lastRowFormulas.length; i++) {
        const formula = lastRowFormulas[i];
        if (formula) {
          let pattern = formula;

          // Replace direct references like A2862, B2862, etc.
          pattern = pattern.replace(new RegExp(`([A-Z]+)${lastRow}`, 'g'), '$1{row}');

          // Replace ranges like $B$2:$B2861
          pattern = pattern.replace(new RegExp(`(\\$[A-Z]+\\$2:\\$[A-Z]+)${lastRow - 1}`, 'g'), '$1{rowMinus1}');

          templateFormulas.push(pattern);
        } else {
          templateFormulas.push(null);
        }
      }

      // Apply updated formulas with correct row numbers
      applyDynamicFormulas(destinationSheet, lastRow + 1, rows.length, templateFormulas, lastRow);

      // Update sheet to maintain the track of the processed block
      const blockRow = blockRowMap[fileName];
      if (blockRow) {
        const statusCol = 3;     // Column C (e.g., for status like "Processed")
        const timestampCol = 4;  // Column D (e.g., for timestamp)

        sourceSheet.getRange(blockRow, statusCol).setValue("Processed");
        sourceSheet.getRange(blockRow, timestampCol).setValue(new Date());

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

function applyDynamicFormulas(sheet, startRow, numRows, formulaTemplates, baseRow) {
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const rowMinus1 = row - 1;

    for (let col = 0; col < formulaTemplates.length; col++) {
      const template = formulaTemplates[col];
      if (template) {
        let formula = template
          .replaceAll('{row}', row)
          .replaceAll('{rowMinus1}', rowMinus1);
        sheet.getRange(row, col + 1).setFormula(formula);
      }
    }
  }
}