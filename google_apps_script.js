/**
 * ============================================================================
 * GOOGLE APPS SCRIPT FOR REAL-TIME GOOGLE SHEET CONTRIBUTION LOGGING
 * ============================================================================
 * 
 * SETUP INSTRUCTIONS (Takes 1 minute):
 * 1. Open Google Sheets (https://sheets.new) and create a new sheet.
 *    Name it e.g. "One Last Onam - Contributions".
 * 
 * 2. In Google Sheets menu, click:
 *    Extensions > Apps Script
 * 
 * 3. Delete any default code in Code.gs, paste this entire file, and click Save (Ctrl+S).
 * 
 * 4. Click "Deploy" (top right blue button) > "New deployment".
 *    - Click the gear icon (Select type) > Choose "Web app".
 *    - Description: "Pookalam Contributions Webhook"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone"  <-- CRITICAL: Choose "Anyone" so visitors can record contributions!
 *    - Click "Deploy".
 * 
 * 5. Copy the "Web app URL" (starts with https://script.google.com/macros/s/...).
 * 
 * 6. Paste that URL into `pookalam.html` in the GOOGLE_SHEET_URL setting at line ~2640.
 * 
 * That's it! Every contribution will instantly appear as a new row in your Google Sheet in real-time.
 * ============================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for other concurrent writes
    lock.waitLock(10000);
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};
    
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    // Format timestamp in Indian Standard Time (Kolkata)
    var timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd/MM/yyyy, hh:mm:ss a");
    
    var name = (data.name || data.userName || "").toString().trim();
    var txnId = (data.transactionId || data.txnId || "").toString().trim();
    var payasam = Number(data.payasamCount || 0);
    var mullapoo = Number(data.mullapooCount || 0);
    var pookalam = Number(data.pookalamAmount || 0);
    var total = Number(data.totalAmount || 0);
    
    // Automatically initialize header row on first run
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Name",
        "Transaction ID / Ref",
        "Payasam Count",
        "Mullapoo Count",
        "Pookalam Contribution (₹)",
        "Total Amount (₹)"
      ]);
      
      // Style header in Onam Gold theme
      var headerRange = sheet.getRange(1, 1, 1, 7);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#D6A43A");
      headerRange.setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
    
    // Append the contribution row
    sheet.appendRow([
      timestamp,
      name,
      "'" + txnId, // Prefix apostrophe so Google Sheets preserves full 12+ digit UTR string
      payasam,
      mullapoo,
      pookalam,
      total
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "success", 
      "message": "Contribution recorded successfully!",
      "row": sheet.getLastRow() 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "error", 
      "message": error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Support GET requests for easy health check in browser
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    "status": "active",
    "service": "Pookalam Contributions Google Sheets Webhook"
  })).setMimeType(ContentService.MimeType.JSON);
}
