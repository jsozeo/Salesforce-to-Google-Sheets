async function getSalesforceData() {
    // Get the active spreadsheet and create/get sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = spreadsheet.getSheetByName("License Details") || spreadsheet.insertSheet("License Details");
    
    // Clear existing content
    detailsSheet.clear();
    
    let etiData = null;
    let prodData = null;
    
    try {
      // Try to fetch ETI data
      try {
        etiData = getSalesforceReport('ETI');
        Logger.log('Successfully fetched ETI data');
      } catch (etiError) {
        Logger.log('Error fetching ETI data: ' + etiError.toString());
      }
      
      // Try to fetch PROD data
      try {
        prodData = getSalesforceReport('PROD');
        Logger.log('Successfully fetched PROD data');
      } catch (prodError) {
        Logger.log('Error fetching PROD data: ' + prodError.toString());
      }
      
      // If both environments failed, throw error
      if (!etiData && !prodData) {
        throw new Error('Failed to fetch data from both ETI and PROD environments');
      }
      
      // Process and merge available data
      const mergedData = mergeEnvironmentData(etiData, prodData);
      
      // Process and display the data
      processAndDisplayData(detailsSheet, mergedData);
      
    } catch (error) {
      Logger.log('Error processing data: ' + error.toString());
      Logger.log('Error stack: ' + error.stack);
      throw new Error('Failed to process data: ' + error.toString());
    }
}

function mergeEnvironmentData(etiData, prodData) {
    // Create maps to track records by their Title
    const recordsMap = new Map();
    
    // Process ETI data if available
    if (etiData) {
      processEnvironmentData(etiData, 'ETI', recordsMap);
    }
    
    // Process PROD data if available
    if (prodData) {
      processEnvironmentData(prodData, 'PROD', recordsMap);
    }
    
    return Array.from(recordsMap.values());
}

function processEnvironmentData(data, environment, recordsMap) {
    if (!data.factMap || !data.groupingsDown) return;
    
    const groupingsMap = {};
    data.groupingsDown.groupings.forEach(grouping => {
      groupingsMap[grouping.value] = {
        title: grouping.label,
        key: grouping.key
      };
    });
    
    for (const key in data.factMap) {
      const entry = data.factMap[key];
      if (!entry.rows || !entry.rows[0] || !entry.rows[0].dataCells) continue;
      
      const cells = entry.rows[0].dataCells;
      const reportId = cells[1].recordId;
      const groupingData = groupingsMap[reportId] || { title: '', key: '' };
      
      // Use title as the key for merging
      const title = groupingData.title;
      if (!title) continue; // Skip if no title found
      
      if (recordsMap.has(title)) {
        // Record exists, update environment info
        const record = recordsMap.get(title);
        if (!record.environments.has(environment)) {
          record.environments.add(environment);
          // Sort environments to ensure consistent order (ETI, PROD)
          record.emplacement = Array.from(record.environments)
            .sort((a, b) => a.localeCompare(b))
            .join(', ');
        }
      } else {
        // New record
        recordsMap.set(title, {
          title: title,
          reportId: reportId,
          folder: cells[0].label || '',
          createdBy: cells[1].label || '',
          createDate: formatDate(cells[3].value) || '',
          modifyDate: formatDate(cells[2].value) || '',
          environments: new Set([environment]),
          emplacement: environment
        });
      }
      
      // Log for debugging
      Logger.log(`Processing ${title} - Environment: ${environment} - Current environments: ${recordsMap.get(title).emplacement}`);
    }
}

function processAndDisplayData(sheet, mergedRecords) {
    // Remove existing filters if any
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    
    // Clear existing content
    sheet.clear();
    
    // Define column configuration
    const columnConfig = {
      Key: { index: 0, width: 50, align: 'center' },
      Title: { index: 1, width: 400 },
      'Report ID': { index: 2, width: 155 },
      Folder: { index: 3, width: 200 },
      'Created By': { index: 4, width: 150 },
      'Create Date': { index: 5, width: 100 },
      'Modify Date': { index: 6, width: 100 },
      Environment: { index: 7, width: 100 }
    };
    
    // Sort records by title
    const sortedRecords = mergedRecords.sort((a, b) => 
      a.title.toLowerCase().localeCompare(b.title.toLowerCase())
    );
    
    // Calculate environment counts
    const etiCount = sortedRecords.filter(record => record.environments.has('ETI')).length;
    const prodCount = sortedRecords.filter(record => record.environments.has('PROD')).length;
    
    // Add total rows for environments first
    const totalRows = [
      ['', 'Total ETI', '', '', '', '', '', etiCount.toString()],
      ['', 'Total PROD', '', '', '', '', '', prodCount.toString()]
    ];
    
    // Write total rows at the very top
    const totalsRange = sheet.getRange(1, 1, 2, 8);
    totalsRange.setValues(totalRows);
    
    // Create headers (now after totals)
    const headers = Object.keys(columnConfig);
    const headerRange = sheet.getRange(3, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Create rows with sequential keys and hyperlinks for Environment
    const rows = sortedRecords.map((record, index) => {
      // Create hyperlinks for each environment
      let environmentLinks = record.environments ? Array.from(record.environments).map((env, i, arr) => {
        const config = SALESFORCE_CONFIG[env];
        // Replace <reportId> in CLICK_LINK with the actual report ID
        const clickLink = config.CLICK_LINK.replace('<reportId>', record.reportId);
        const hyperlink = `HYPERLINK("${clickLink}"; "${env}")`;
        
        // If this is not the last environment, add the concatenation operator
        return i < arr.length - 1 ? 
          `${hyperlink}&", "&` : 
          hyperlink;
      }).join('') : "";
  
      // Add the formula prefix only once at the start
      if (environmentLinks) {
        environmentLinks = '=' + environmentLinks;
      }
  
      return [
        (index + 1).toString(),
        record.title,
        record.reportId,
        record.folder,
        record.createdBy,
        record.createDate,
        record.modifyDate,
        environmentLinks
      ];
    });
    
    // Write data and format
    if (rows.length > 0) {
      // Write main data (starting after header)
      const dataRange = sheet.getRange(4, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Set column widths
      headers.forEach((header, index) => {
        sheet.setColumnWidth(index + 1, columnConfig[header].width);
      });
      
      // Format total rows at the top
      totalsRange
        .setBackground('#E8F5E9')  // Light green background
        .setFontWeight('bold')
        .setBorder(true, true, true, true, true, true, '#2E7D32', SpreadsheetApp.BorderStyle.SOLID);
      
      // Format header
      headerRange
        .setBackground('#2E7D32')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBorder(true, true, true, true, true, true, 'white', SpreadsheetApp.BorderStyle.SOLID);
      
      // Format data rows with specific alignments
      dataRange
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle')
        .setWrap(true);
      
      // Center the Key column
      sheet.getRange(4, 1, rows.length, 1)
        .setHorizontalAlignment('center');
      
      // Add alternating row colors for data rows
      for (let i = 4; i <= rows.length + 3; i++) {
        const rowRange = sheet.getRange(i, 1, 1, headers.length);
        if ((i - 4) % 2 === 0) {
          rowRange.setBackground('#FFFFFF');
        } else {
          rowRange.setBackground('#F5F5F5');
        }
        
        rowRange.setBorder(
          true, true, true, true, null, null,
          '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID
        );
      }
      
      // Special formatting for specific columns
      const dateColumns = [6, 7];
      dateColumns.forEach(col => {
        sheet.getRange(4, col, rows.length, 1)
          .setHorizontalAlignment('center');
      });
      
      // Format Environment column
      const environmentColumn = sheet.getRange(4, 8, rows.length, 1);
      environmentColumn
        .setHorizontalAlignment('center')
        .setFontWeight('bold')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      
      // Center the count values in total rows
      sheet.getRange(1, 8, 2, 1)
        .setHorizontalAlignment('center');
      
      // Add filters (for header and data rows only)
      sheet.getRange(3, 1, rows.length + 1, headers.length).createFilter();
      
      Logger.log(`Successfully processed ${rows.length} rows of data`);
      Logger.log(`ETI Count: ${etiCount}, PROD Count: ${prodCount}`);
    }
}

// Helper function to format dates
function formatDate(dateString) {
    if (!dateString) return '';
    
    // Check if the date is already in dd/mm/yyyy format
    if (dateString.includes('/')) {
      return dateString;
    }
    
    // Convert from yyyy-mm-dd to dd/mm/yyyy
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// Update the onOpen function to include OAuth menu items
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Salesforce Data');
    
    // Add OAuth menu items
    addOAuthMenuItems(menu);
    
    // Add the main refresh function
    menu.addItem('Refresh License Data', 'getSalesforceData')
        .addToUi();
}
