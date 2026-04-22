/**
 * =====================================================
 * Certificate Automation using Google Apps Script
 * GitHub Safe Public Version
 * =====================================================
 * Features:
 * - Bulk certificate generation
 * - Personalized names
 * - PDF export
 * - Email delivery
 * - Save certificate URL in Sheet
 * =====================================================
 */


/* ==============================
   CONFIGURATION
============================== */

const CERTIFICATION_URL = "PASTE_GOOGLE_SLIDES_TEMPLATE_URL_HERE";
const OUTPUT_FOLDER_NAME = "Completed PDFs";
const MENU_NAME = "Automation Menu";
const MENU_ITEM = "Email All Certificates";


/* ==============================
   ON OPEN MENU
============================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM, "createCopy")
    .addToUi();
}


/* ==============================
   MAIN FUNCTION
============================== */

function createCopy() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No student data found.");
    return;
  }

  var certId = SlidesApp.openByUrl(CERTIFICATION_URL).getId();
  var templateFile = DriveApp.getFileById(certId);

  var folder = getOrCreateFolder(OUTPUT_FOLDER_NAME);

  var mailedCount = 0;
  var skippedCount = 0;

  for (var row = 2; row <= lastRow; row++) {

    var name = sheet.getRange(row, 1).getDisplayValue().toString().trim();
    var email = sheet.getRange(row, 2).getDisplayValue().toString().trim();

    if (!name || !email) {
      skippedCount++;
      continue;
    }

    try {

      var certName = sanitizeFileName(name) + "_certificate";

      // Copy Template
      var certCopy = templateFile.makeCopy(certName);
      var certCopyId = certCopy.getId();

      // Update Slide
      var presentation = SlidesApp.openById(certCopyId);

      presentation.replaceAllText("[[name]]", name);
      presentation.replaceAllText("[[date]]", getTodayDate());

      presentation.saveAndClose();

      // Convert to PDF
      var pdfBlob = certCopy.getAs(MimeType.PDF)
        .setName(certName + ".pdf");

      var pdfFile = folder.createFile(pdfBlob);

      // Save URL in Column C
      sheet.getRange(row, 3).setValue(pdfFile.getUrl());

      // Email
      GmailApp.sendEmail(
        email,
        "Your Certificate of Participation",
        getEmailBody(name),
        { attachments: [pdfBlob] }
      );

      // Trash temp Slides file
      DriveApp.getFileById(certCopyId).setTrashed(true);

      mailedCount++;

    } catch (err) {

      Logger.log("Row " + row + " Error: " + err);
      skippedCount++;

    }
  }

  SpreadsheetApp.getUi().alert(
    "Process Completed\n\n" +
    "Certificates Mailed: " + mailedCount + "\n" +
    "Skipped Rows: " + skippedCount
  );
}


/* ==============================
   HELPERS
============================== */

function getOrCreateFolder(folderName) {

  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}


function getTodayDate() {
  return Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    "MMMM dd, yyyy"
  );
}


function getEmailBody(name) {

  return (
    "Hi " + name + ",\n\n" +
    "Attached is your Certificate of Participation.\n\n" +
    "Thank you for attending the event.\n\n" +
    "Regards"
  );
}


function sanitizeFileName(name) {
  return name.replace(/[\\\/:*?\"<>|]/g, "").trim();
}