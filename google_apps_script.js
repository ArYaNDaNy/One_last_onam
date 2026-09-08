/**
 * Google Apps Script for Real-Time Google Sheet / Excel Logging
 * 
 * Instructions:
 * 1. Open a new Google Sheet (e.g. "Pookalam Contributions").
 * 2. Set Row 1 headers to:
 *    Timestamp | name | tranasaction id | payasam count | mullapoo count | contribution for pookalam amnt | total amt
 * 3. Go to Extensions > Apps Script.
 * 4. Paste this code and click "Deploy" > "New deployment".
 * 5. Select Type: "Web App", Execute as: "Me", Who has access: "Anyone".
 * 6. Copy the Web App URL and paste it in `pookalam.html` under GOOGLE_SHEET_WEBHOOK_URL.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var timestamp = new Date().toLocaleString();
    
    sheet.appendRow([
      timestamp,
      data.name,
      data.transactionId,
      data.payasamCount,
      data.mullapooCount,
      data.pookalamAmount,
      data.totalAmount
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
