

function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('My Custom Menu')
    .addItem('Create PDF using doc template', 'createPdf')
    .addToUi();
}




function createPdf() {

  // Replace this with ID of your template document.
  var TEMPLATE_ID = '1P9qsF1zMLxbbo36n3bAasa6sJdfMlOhbEkmnrufGJHs'

  if (TEMPLATE_ID === '') {

    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }



  // Set up the docs and the spreadsheet access
   
    
    SHEET_NAME = "sheet 1",
    SPREADSHEET_ID = '1rRNPKprgyDVPHbBywNcFz42IMLlw9yT_w44hgir47Uk',
    FOLDER_ID = '1wYR4c69FJTcuewGZc5IJenHpTwEOHwOs'
   

  //var imageId = "https://drive.google.com/open?id=1UcgIMLosmg6wXJLsbDla_OqNa1XTTlWU";
  //var fileID = imageId.match(/[\w\_\-]{25,}/).toString();
  //var blob   = DriveApp.getFileById(fileID).getBlob();
  //var newSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();


  /**
   * Used to find specific text and replace it with an image
   * @param body  The vody containing the text
   * @param searchText  The text that is replaced
   * @param image The image used to replace the text
   * @param width The width of the image
   * @return The picture in the right cell
   * 
   */
  var replaceTextToImage = function (body, searchText, image, width) {
    var next = body.findText(searchText);
    if (!next) return;
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);
    if (width && typeof width == "number") {
      var w = img.getWidth();
      var h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    }
    return next;
  };


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];


  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  var firstRow = selection.getActiveRange().getRow();
  var lastRow = selection.getActiveRange().getLastRow();
  var totalRow = lastRow - firstRow;



  Logger.log('first row: ' + firstRow);
  Logger.log('last row: ' + lastRow);
  Logger.log('total row: ' + totalRow.toString());

  var range = sheet.getRange("A" + firstRow + ":A" + lastRow);

  Logger.log('total row: ' + range.getNumRows());


  for (var i = firstRow; i < lastRow + 1; i++) {

 

    var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
    copyId = copyFile.getId(),
    copyDoc = DocumentApp.openById(copyId)
    copy = copyDoc.getBody()
    
  
    var sheet = SpreadsheetApp.getActiveSheet();
   

    var rowValues = getRowValues(i); 
      

    copy.replaceText("%LastName%", rowValues[0]);
    copy.replaceText("%FirstName%", rowValues[1]);
    copy.replaceText("%Birth%", Utilities.formatDate(new Date(rowValues[2]), "GMT", "dd-MM-yyyy"));
    copy.replaceText("%Guardian%", rowValues[3]);
    copy.replaceText("%GuardianRelation%", rowValues[4]);
    copy.replaceText("%GuardianJob%", rowValues[5]);
    copy.replaceText("%FatherName%", rowValues[6]);
    copy.replaceText("%FatherDateDeath%", Utilities.formatDate(new Date(rowValues[7]), "GMT", "dd-MM-yyyy"));
    copy.replaceText("%FatherCauseDeath%", rowValues[8]);
    copy.replaceText("%NumberSiblings%", rowValues[9]);
    copy.replaceText("%Country%", rowValues[10]);
    copy.replaceText("%Address%", rowValues[11]);
    copy.replaceText("%ContactInformation%", rowValues[12]); 

    // Replace the keys with the spreadsheet values

    PDF_FILE_NAME = rowValues[1];


    // Create the PDF file, rename it if required and delete the doc copy

    
    copyDoc.saveAndClose();
    var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'));

    if (PDF_FILE_NAME !== '') {

      newFile.setName(PDF_FILE_NAME)
      var fld = DriveApp.getFolderById(FOLDER_ID);
      finalPDF = newFile.makeCopy(fld);
      newFile.setTrashed(true);
    }
    copyFile.setTrashed(true)
    
   
  }



};
 
function getRowValues(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return rowValues;
} 

