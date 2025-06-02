// Google Apps Script Code
// วิธีใช้:
// 1. ไปที่ Google Sheets ที่สร้างไว้
// 2. คลิก Extensions > Apps Script
// 3. ลบโค้ดเดิมแล้วคัดลอกโค้ดนี้ไปวาง
// 4. Save และ Deploy > New Deployment
// 5. เลือก Type: Web app
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Deploy แล้วคัดลอก Web app URL

function doGet(e) {
  const action = e.parameter.action;
  const sheetName = e.parameter.sheet;
  
  if (action === 'read') {
    return handleRead(sheetName);
  }
  
  return ContentService.createTextOutput(
    JSON.stringify({ error: 'Invalid action' })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const sheetName = data.sheet;
    const rowData = data.data;
    
    if (action === 'create') {
      return handleCreate(sheetName, rowData);
    } else if (action === 'update') {
      return handleUpdate(sheetName, rowData);
    } else if (action === 'delete') {
      return handleDelete(sheetName, rowData);
    }
    
    return ContentService.createTextOutput(
      JSON.stringify({ error: 'Invalid action' })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleRead(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      createSheet(sheetName);
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const jsonData = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    return ContentService.createTextOutput(
      JSON.stringify({ success: true, data: jsonData })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleCreate(sheetName, rowData) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = createSheet(sheetName);
    }
    
    // Add timestamp
    rowData.timestamp = new Date().toISOString();
    
    // Get headers
    let headers = [];
    if (sheet.getLastRow() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    // Add new headers if needed
    const newHeaders = Object.keys(rowData).filter(key => !headers.includes(key));
    if (newHeaders.length > 0) {
      headers = headers.concat(newHeaders);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // Prepare row data
    const row = headers.map(header => rowData[header] || '');
    
    // Append row
    sheet.appendRow(row);
    
    return ContentService.createTextOutput(
      JSON.stringify({ success: true, message: 'Data saved successfully' })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUpdate(sheetName, updateData) {
  // Implement update logic if needed
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, message: 'Update not implemented' })
  ).setMimeType(ContentService.MimeType.JSON);
}

function handleDelete(sheetName, deleteData) {
  // Implement delete logic if needed
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, message: 'Delete not implemented' })
  ).setMimeType(ContentService.MimeType.JSON);
}

function createSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(sheetName);
  
  // Set headers based on sheet type
  let headers = [];
  
  switch(sheetName) {
    case 'Letters':
      headers = ['from', 'to', 'content', 'date', 'timestamp'];
      break;
    case 'Wishlist':
      headers = ['title', 'description', 'by', 'date', 'timestamp'];
      break;
    case 'Dreams':
      headers = ['dream', 'category', 'completed', 'date', 'timestamp'];
      break;
    case 'Stories':
      headers = ['title', 'content', 'author', 'date', 'timestamp'];
      break;
    case 'Scores':
      headers = ['game', 'player', 'score', 'date', 'timestamp'];
      break;
    default:
      headers = ['data', 'date', 'timestamp'];
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#FF69B4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  return sheet;
}

// Helper function to initialize all sheets
function initializeSheets() {
  const sheetNames = ['Letters', 'Wishlist', 'Dreams', 'Stories', 'Scores'];
  
  sheetNames.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      createSheet(sheetName);
    }
  });
}

// Run this function once to set up all sheets
function setup() {
  initializeSheets();
  
  // You can also set up triggers or permissions here if needed
  
  return 'Setup complete! All sheets created.';
}