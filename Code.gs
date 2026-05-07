// Google Apps Script - Code.gs
// Handles form submissions and saves to Google Sheets
// Fields: Name, Email, Phone, Occasion Type, Color Theme, Delivery Date, Tell Us About

// Main doGet handler - processes query parameters
function doGet(e) {
  return handleFormSubmission(e);
}

// Main doPost handler - processes POST requests
function doPost(e) {
  return handleFormSubmission(e);
}

// Core function to process form data
function handleFormSubmission(e) {
  try {
    // Get parameters from query string or POST body
    const params = e.parameter || {};
    const name = params.name || '';
    const email = params.email || '';
    const phone = params.phone || '';
    const occasionType = params.occasionType || '';
    const colorTheme = params.colorTheme || '';
    const deliveryDate = params.deliveryDate || '';
    const tellUsAbout = params.tellUsAbout || '';

    // Validate required fields
    if (!name || !email || !phone) {
      return ContentService.createTextOutput(
        JSON.stringify({
          success: false,
          error: 'Missing required fields: name, email, phone'
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Get the spreadsheet and sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Responses');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Responses');
      // Add headers
      sheet.appendRow([
        'Timestamp',
        'Name',
        'Email',
        'Phone',
        'Occasion Type',
        'Color Theme',
        'Delivery Date',
        'Tell Us About'
      ]);
      // Make headers bold
      const headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold');
    }

    // Add timestamp
    const timestamp = new Date().toLocaleString('en-IN', {
      timeZone: 'Asia/Kolkata',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });

    // Append the new row
    sheet.appendRow([
      timestamp,
      name,
      email,
      phone,
      occasionType,
      colorTheme,
      deliveryDate,
      tellUsAbout
    ]);

    // Return success response
    return ContentService.createTextOutput(
      JSON.stringify({
        success: true,
        message: 'Form submitted successfully'
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Return error response
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        error: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
