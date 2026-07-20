/**
 * ══════════════════════════════════════════════════════════════════
 *  GHAFES Ministry Reports — Google Apps Script
 *  This file runs inside Google Apps Script (NOT your project folder).
 *  Follow the steps in README.md → "Setting Up Google Sheets Sync"
 * ══════════════════════════════════════════════════════════════════
 *
 *  HOW TO USE:
 *  1. Open https://script.google.com  and sign in with your Google account
 *  2. Click "New Project"
 *  3. Delete everything in the editor, then paste THIS ENTIRE FILE
 *  4. Find the line below that says YOUR_SPREADSHEET_ID_HERE
 *     and replace it with the real ID from your Google Sheet URL
 *     Example URL:  https://docs.google.com/spreadsheets/d/1BxiM.../edit
 *     Your ID is:                                              ^^^^^^
 *  5. Click File → Save  (name it "GHAFES Sync")
 *  6. Click Deploy → New Deployment
 *  7. Click the gear icon next to "Type" and choose "Web app"
 *  8. Set:  Execute as = "Me"
 *           Who has access = "Anyone"
 *  9. Click Deploy → copy the URL it gives you
 * 10. Paste that URL into the GHAFES app → ⚙ Settings → Script URL
 */

// ──────────────────────────────────────────────────────────────────
// PASTE YOUR GOOGLE SHEET ID BETWEEN THE QUOTES ON THE NEXT LINE
// ──────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";

// ─── Main POST handler — called by the GHAFES web app ─────────────
function doPost(e) {
  try {
    // Parse the data sent from the GHAFES app
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Sheet name comes from the app (e.g. "July 2026" or "Ministry Reports")
    const sheetName = data.sheetName || "Ministry Reports";
    let sheet = ss.getSheetByName(sheetName);

    // Create the sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);

      // Write column headers on row 1 if provided
      if (data.headers && data.headers.length > 0) {
        sheet.appendRow(data.headers);

        // Style the header row — GHAFES navy blue
        const headerRange = sheet.getRange(1, 1, 1, data.headers.length);
        headerRange
          .setBackground("#1E3A6E")
          .setFontColor("#FFFFFF")
          .setFontWeight("bold")
          .setFontSize(10);

        // Freeze the header row so it stays visible when scrolling
        sheet.setFrozenRows(1);

        // Auto-resize columns for readability
        sheet.autoResizeColumns(1, data.headers.length);
      }
    }

    // Append each data row sent by the app
    const rows = data.rows || [];
    rows.forEach(function(row) {
      sheet.appendRow(row);
    });

    // Return success message back to the app
    return ContentService
      .createTextOutput(JSON.stringify({
        success:  true,
        sheet:    sheetName,
        rowsAdded: rows.length,
        message:  "Data saved successfully to " + sheetName
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Return error details if something went wrong
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error:   err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Simple GET handler — lets you test the script is running ──────
function doGet(e) {
  return ContentService
    .createTextOutput("✓ GHAFES Ministry Reports Sync is active and running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ─── Optional: Run this manually to test the connection ───────────
function testConnection() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const name  = ss.getName();
  Logger.log("✓ Connected to spreadsheet: " + name);
  Logger.log("✓ Script is working correctly.");
}
