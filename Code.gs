

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

  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(),
  copyId = copyFile.getId(),
  copyDoc = DocumentApp.openById(copyId),
  copyBody = copyDoc.getActiveSection(),
  activeSheet = SpreadsheetApp.getActiveSheet(),
  numberOfColumns = activeSheet.getLastColumn(),
  activeRowIndex = activeSheet.getActiveRange().getRowIndex(),
  activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues(),
  headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues(),
  columnIndex = 0,
  ReceiptID="",
  currentDate,
  finalPDF

  //var imageId = "https://drive.google.com/open?id=1UcgIMLosmg6wXJLsbDla_OqNa1XTTlWU";
  //var fileID = imageId.match(/[\w\_\-]{25,}/).toString();
  //var blob   = DriveApp.getFileById(fileID).getBlob();
  var newSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  
  


  

  var replaceTextToImage = function(body, searchText, image, width) {
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

  var data = activeSheet.getRange(2, 1, activeSheet.getLastRow()-1, activeSheet.getLastColumn()).getValues();
  


  var rowValues = getRowValues(); 

  for (var i in data){
      var row = data[i];
      
      
      var body = copyDoc.getActiveSection();
      
      body.replaceText("%LastName%", rowValues[0]);
      body.replaceText("%FirstName%",rowValues[1] );
      body.replaceText("%Birth%",Utilities.formatDate(new Date(rowValues[2]), "GMT" , "dd-MM-yyyy"));
      body.replaceText("%Guardian%", rowValues[3]);
      body.replaceText("%GuardianRelation%", rowValues[4]);
      body.replaceText("%GuardianJob%", rowValues[5]);
      body.replaceText("%FatherName%", rowValues[6]);
      body.replaceText("%FatherDateDeath%", Utilities.formatDate(new Date(rowValues[7]), "GMT" , "dd-MM-yyyy")); 
      body.replaceText("%FatherCauseDeath%", rowValues[8]);
      body.replaceText("%NumberSiblings%", rowValues[9]);
      body.replaceText("%Country%", rowValues[10]);
      body.replaceText("%Address%", rowValues[11]);
      body.replaceText("%ContactInformation%", rowValues[12]);



  }

  var tables = body.getTables();
  

  
  // Replace the keys with the spreadsheet values

  for (;columnIndex < headerRow[0].length; columnIndex++) {
  currentHeaderCell = headerRow[0][columnIndex] ;

  if (currentHeaderCell == "First Name(s)")
  {
  PDF_FILE_NAME = activeRow[0][columnIndex];
  }





  copyBody.replaceText('%' + headerRow[0][columnIndex] + '%', activeRow[0][columnIndex])
  }


  // Create the PDF file, rename it if required and delete the doc copy

  copyDoc.saveAndClose()

  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))

  if (PDF_FILE_NAME !== '') {

  newFile.setName(PDF_FILE_NAME + '_' + ReceiptID)
  var fld = DriveApp.getFolderById('1wYR4c69FJTcuewGZc5IJenHpTwEOHwOs');
  finalPDF = newFile.makeCopy(fld);
  newFile.setTrashed(true);
  }

  copyFile.setTrashed(true)
  



  // save the PDF URL to the active row
  //var pdfUrl = "https://drive.google.com/open?id=" + finalPDF.getId();
  //activeSheet.getRange(2,11).setValue(pdfUrl);



  };

  function getRowValues() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var rowIndex = sheet.getCurrentCell().getRow();
    var rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return rowValues;
  }




