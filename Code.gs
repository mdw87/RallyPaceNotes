// Note: this code is available on GitHub here:
// https://github.com/mdw87/RallyPaceNotes

// Add menu to UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("RallyPaceNotes")
      .addItem("Generate Pace Notes", "generatePaceNotes")
      .addToUi();
}

function generatePaceNotes() {
  var LINES_PER_PAGE = 7;

  // Style for left align
  var leftStyle = {};
  leftStyle[DocumentApp.Attribute.FONT_SIZE] = 10;  

  // Style for right align
  var rightStyle = {};
  rightStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
  DocumentApp.HorizontalAlignment.RIGHT;
  rightStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  
  // Style for center align
  var centerStyle = {};
  centerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  centerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

  // Style for the title
  var titlePageStyle = {};
  titlePageStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  titlePageStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  titlePageStyle[DocumentApp.Attribute.FONT_SIZE] = 36;
  titlePageStyle[DocumentApp.Attribute.BOLD] = true;
  
  // Style for the stage notes
  var noteStyle = {};
  noteStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  noteStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = 
    DocumentApp.VerticalAlignment.CENTER;
  noteStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  noteStyle[DocumentApp.Attribute.FONT_SIZE] = 32;
  noteStyle[DocumentApp.Attribute.BOLD] = true;
  
  // Style for the distance
  var distStyle = {};
  distStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  distStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = 
    DocumentApp.VerticalAlignment.CENTER;
  distStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  distStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  distStyle[DocumentApp.Attribute.BOLD] = true;
  
  // Style for the next note
  var nextStyle = {};
  nextStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.RIGHT;
  nextStyle[DocumentApp.Attribute.BOLD] = true;
  nextStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
  
  var sheet = SpreadsheetApp.getActive();
  var title = sheet.getName();
  
  // Figure out the last row index, in order to calculate how many notes there are
  var numRows = sheet.getLastRow() - 7; // Notes start on the 8th row
  var numPages = Math.ceil(numRows / LINES_PER_PAGE);
  var pageNum = 1;
  var rallyName = sheet.getRange("B2").getValue();
  var stageNumber = sheet.getRange("B4").getValue();
  var stageName = sheet.getRange("B5").getValue();
  var stageDistance = sheet.getRange("B6").getValue();;

  var outputDoc = DocumentApp.create(rallyName + " | " + "SS" + stageNumber + " | " + stageName);
  var outputBody = outputDoc.getBody();
  var docUrl = outputDoc.getUrl();

  // Set margin size
  outputBody.setMarginTop(20);
  outputBody.setMarginBottom(20);
  outputBody.setMarginLeft(20);
  outputBody.setMarginRight(20);
  
  // Set output cell
  var output_cell = sheet.getRange("D2");
  output_cell.setValue("Generating Notes...");
  
  outputBody.setText('');
  
/*
  // Create Title Page
  var stageTitle = outputBody.appendParagraph(stageName);
  stageTitle.setAttributes(titlePageStyle);
  outputBody.appendPageBreak();
  
  // Create Title Table
  var titleCell = [["SS" + stageNumber + ': ' + stageName, "Distance: " + stageDistance, "Page " + pageNum + "/" + numPages]];
  var titleTable = outputBody.appendTable(titleCell);
  titleTable.getCell(0, 0).getChild(0).setAttributes(leftStyle);
  titleTable.getCell(0, 2).getChild(0).setAttributes(rightStyle);
  titleTable.getCell(0, 1).getChild(0).setAttributes(centerStyle);
*/
  
  titleCell = [[stageNumber + ': ' + stageName, "Page " + pageNum + "/" + numPages]];
  titleTable = outputBody.appendTable(titleCell);
  titleTable.getCell(0, 1).getChild(0).setAttributes(rightStyle);
  titleTable.getCell(0, 0).getChild(0).setAttributes(leftStyle);
  
  // Create Notes Table
  
  var cells = [
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', '']
  ];
  var outputTable = outputBody.appendTable(cells);
  outputTable.setColumnWidth(0, 50);
  outputTable.setColumnWidth(2, 60);

  //Go row by row and print the notes
  var currRow = 0;
  for (var i = 8; i <= sheet.getLastRow(); i++) {
    //set the row height
    var row = outputTable.getRow(currRow);
    row.setMinimumHeight(85);
    var cell = outputTable.getCell(currRow, 1);
    var distCell = outputTable.getCell(currRow, 0);
    var remDistCell = outputTable.getCell(currRow, 2);
    //vertically align everything in the center
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    distCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    remDistCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    //remove blank text
    cell.removeChild(cell.getChild(0));
    distCell.removeChild(distCell.getChild(0));
    remDistCell.removeChild(remDistCell.getChild(0));
    //add note
    var distance = sheet.getRange("A" + i).getValue();
    var strRemDist = "";
    var strDist = "";
    if ( distance != "" ){
      strDist = parseFloat(distance).toFixed(1);
      var remDist = stageDistance - distance;
      var strRemDist = parseFloat(remDist).toFixed(1);
    }
    var note = sheet.getRange("B" + i).getValue();
    var par = cell.appendParagraph(note);
    var distPar = distCell.appendParagraph(strDist);
    var remDistPar = remDistCell.appendParagraph(strRemDist);
    par.setAttributes(noteStyle);
    distPar.setAttributes(distStyle);
    remDistPar.setAttributes(distStyle);
    currRow = currRow + 1;
    //after the last row is filled out, create new page
    if (currRow > LINES_PER_PAGE - 1) {
      //fill out the 'first call of next page'
      var nextCell = outputTable.getCell(LINES_PER_PAGE, 1);
      nextCell.removeChild(nextCell.getChild(0));
      var nextPar = nextCell.appendParagraph(sheet.getRange("B" + (i + 1)).getValue());
      nextPar.setAttributes(nextStyle);
      outputBody.appendPageBreak();
      //create next note page
      pageNum = pageNum + 1;
      titleCell = [[stageNumber + ': ' + stageName, "Page " + pageNum + "/" + numPages]];
      titleTable = outputBody.appendTable(titleCell);
      titleTable.getCell(0, 1).getChild(0).setAttributes(rightStyle);
      titleTable.getCell(0, 0).getChild(0).setAttributes(leftStyle);
      //add prev note
      var prevNote = outputBody.appendParagraph(sheet.getRange("B" + i).getValue());
      prevNote.setAttributes(leftStyle);
      //create empty table
      var outputTable = outputBody.appendTable(cells);
      outputTable.setColumnWidth(0, 50);
      outputTable.setColumnWidth(2, 60);
      //reset currRow
      currRow = 0;
      //resume filling notes!
    }
  }
  
  // Update the output cell with link
  output_cell.setValue(docUrl);
}
