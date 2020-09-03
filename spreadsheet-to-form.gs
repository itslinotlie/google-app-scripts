/*
General knowledge:
- Arrays are 0-indexed
- ss.getRange(row, col, numRow, numCol) is 1-indexed
*/
//global variables
var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //getActiveSheet() == current sheet opened in the spreadsheet

//simplifies methods if these 3 are global
var dataRange = ss.getDataRange(); //2d array dimensions
var data = dataRange.getValues(); //2d array with values
var question; 

var headerSize = 4;
var optionStart = 8; //0-indexed
var optionLength = 10;
var desRow = 20, desCol = 20;
var correctColor = "#00ff00"; //default neon green highlight

function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu("Forms");
  menu.addItem("Initilize Spreadsheet", "createTemplate").addToUi();
  menu.addItem("Create Google Form", "createForm").addToUi();
}
function createTemplate() {
  //setting up spreadsheet dimensions
  let curRow = ss.getMaxRows(), curCol = ss.getMaxColumns();
  if(curRow!==desRow) //Exception: Invalid argument is thrown if you .inserRowsAfter(X, 0)
    curRow>desRow? ss.deleteRows(desRow+1, curRow-desRow):ss.insertRowsAfter(curRow-1, desRow-curRow);
  if(curCol!==desCol)
    curCol>desCol? ss.deleteColumns(desCol+1, curCol-desCol):ss.insertColumnsAfter(curCol-1, desCol-curCol);
  curRow = ss.getMaxRows(); curCol = ss.getMaxColumns();

  ss.clear(); //clears formatting
  ss.setRowHeights(1, curRow, 21); ss.setColumnWidths(1, curCol, 100); //resize cells to default size
  ss.getRange(1, 1, desRow, desCol).setDataValidation(null); //clears data formatting so you dont need to create a new sheet

  //info to fill in/use
  ss.getRange("A1").setValue("Form Title:");
  ss.getRange("A2").setValue("Form Desciption:");
  ss.getRange("A3").setValue("Highlight Color"); 
  ss.getRange("B3").setBackground("#00ff00");
  ss.getRange("C1").setValue("Folder ID:");
  //delete the following line if you plan on copying this file
  ss.getRange("D1").setValue("1D2yMTtKfq9ey5awuTbEiHXViCHDYgejH"); //delete when in production
  ss.getRange("C2").setValue("Public URL:");
  ss.getRange("C3").setValue("Private URL:");

  //various boolean fields
  ss.getRange("E1").setValue("One Response per User?");
  ss.getRange("E2").setValue("Can Edit Response?");
  ss.getRange("E3").setValue("Collects Email?");
  ss.getRange("G1").setValue("Progress Bar?");
  ss.getRange("G2").setValue("Link to Respond Again?"); //test to see difference between this and one response per user
  ss.getRange("G3").setValue("Publishing Summary?");  

  //question headers
  let charReq = String.fromCharCode(65+optionStart-1); 
  let optStart = String.fromCharCode(65+optionStart)+headerSize; //3 represents row # (1-indexed)
  let optEnd = String.fromCharCode(65+optionStart+optionLength-1)+headerSize; //Have to -1 for some reason? (check later)

  ss.getRange("A"+headerSize).setValue("Question Type");
  ss.getRange("B"+headerSize).setValue("Question");
  ss.getRange("C"+headerSize).setValue("Instructions");
  ss.getRange("D"+headerSize).setValue("Points");
  ss.getRange("E"+headerSize).setValue("Correct Text");
  ss.getRange("F"+headerSize).setValue("Incorrect Text");
  ss.getRange("G"+headerSize).setValue("URL/ID");
  ss.getRange(charReq+headerSize).setValue("Required?");
  ss.getRange(optStart+":"+optEnd).setValue("OPTION");

  //formatting
  ss.getRange("4:4").setHorizontalAlignment("center");
  ss.setFrozenRows(headerSize); 
  ss.setFrozenColumns(2);
  ss.setColumnWidth(5, 200);
  ss.setColumnWidth(7, 200);
  ss.getRange("D1:D3").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  //data validation //https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder#setAllowInvalid
  const options = ["MC", "CHECKBOX", "SHORTANSWER", "PARAGRAPH", "PAGEBREAK", "HEADER", "IMAGE", "IMAGE-DRIVE", "VIDEO"];
  const bool = ["TRUE", "FALSE"]; 
  setValidation("F1:F3", bool); //boolean questions
  setValidation("H1:H3", bool);
  setValidation("A"+(headerSize+1)+":A"+curRow, options); //question type
  setValidation(charReq+(headerSize+1)+":"+charReq+curRow, bool); //required?

  //very "hacky" solution for "locking" cells from being edited
  // ss.getRange("G1").setValue("H1 and H2 are locked");
  // const blank = [""];  
  // ss.getRange("H1:H2").setDataValidation(SpreadsheetApp.newDataValidation()
  //   .setAllowInvalid(false).requireValueInList(blank, false).build());
  // const tmp = [ss.getRange("A1").getDisplayValue()];
  // ss.getRange("A1").setDataValidation(SpreadsheetApp.newDataValidation()
  //   .setAllowInvalid(false).requireValueInList(tmp, false).build());
}
function setValidation(range, list) {
  ss.getRange(range).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(list, true).build());
}
function createForm() {
  let row = dataRange.getNumRows();
  var form = FormApp.create("Untitled form"); //has to be initialized
  let formID = form.getId();
  let file = DriveApp.getFileById(formID);

  //setting form info to spreadsheet
  let publicUrl = form.getPublishedUrl();
  let privateUrl = form.getEditUrl();
  ss.getRange("D2").setValue(publicUrl);
  ss.getRange("D3").setValue(privateUrl);

  //moving form to the folder (if possible)
  if(data[0][3]!=='') {
    let folder = DriveApp.getFolderById(data[0][3]);
    file.moveTo(folder);
  }

  //filling in form info
  file.setName(data[0][1]);
  form.setTitle(data[0][1]);
  form.setDescription(data[1][1]);
  form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled

  //boolean info
  correctColor = ss.getRange("B3").getBackground();
  if(data[0][5]!=='') form.setLimitOneResponsePerUser(data[0][5]);
  if(data[1][5]!=='') form.setAllowResponseEdits(data[1][5]);
  if(data[2][5]!=='') form.setCollectEmail(data[2][5]); //reveals email address at the top of the form, allows you to send a copy to yourself at the bottom
  if(data[0][7]!=='') form.setProgressBar(data[0][7]);
  if(data[1][7]!=='') form.setShowLinkToRespondAgain(data[1][7]);
  if(data[2][7]!=='') form.setPublishingSummary(data[2][7]); //reveals question distribution, but no answers

  for (let i=headerSize;i<row;i++) {
    let x = data[i][0]; 
    if(x==='') continue; 
    if(x==="MC") question = form.addMultipleChoiceItem();
    else if(x==="CHECKBOX") question = form.addCheckboxItem();
    else if(x==="SHORTANSWER") question = form.addTextItem(); //how do points work with short responses...
    else if(x==="PARAGRAPH") question = form.addParagraphTextItem();
    else if(x==="PAGEBREAK") question = form.addPageBreakItem();
    else if(x==="HEADER") question = form.addSectionHeaderItem(); //these are stackable, but don't look the greatest
    else if(x==="IMAGE") question = form.addImageItem().setImage(UrlFetchApp.fetch(data[i][6])); //imageItem's helptext dont show in Forms
    else if(x==="IMAGE-DRIVE") question = form.addImageItem().setImage(DriveApp.getFileById(data[i][6]));
    else if(x==="VIDEO") question = form.addVideoItem().setVideoUrl(data[i][6]);
    setUpQuestion(i);
  }
}
const choices = [FormApp.ItemType.CHECKBOX, FormApp.ItemType.MULTIPLE_CHOICE];
const visual = [FormApp.ItemType.IMAGE, FormApp.ItemType.VIDEO];
const invalid = [
  FormApp.ItemType.PAGE_BREAK, FormApp.ItemType.SECTION_HEADER, 
  FormApp.ItemType.IMAGE, FormApp.ItemType.VIDEO
];
function setUpQuestion(i) {
  if(data[i][1]!=='') question.setTitle(data[i][1]);
  if(data[i][2]!=='') question.setHelpText(data[i][2]);

  for (let j=0;j<choices.length;j++) { //only certain item types can add options
    if(question.getType()===choices[j]) addOptions(i);
  }
  for (let j=0;j<visual.length;j++) { //only vertian item types can be visually formatted
    if(question.getType()===visual[j]) formatVisual();
  }
  for (let j=0;j<invalid.length;j++) { //makes sure invalid item types end here  
    if(question.getType()===invalid[j]) return;
  }
  if(data[i][3]!=='') question.setPoints(data[i][3]);
  if(data[i][4]!=='') question.setFeedbackForCorrect(FormApp.createFeedback().setText(data[i][4]).build());
  if(data[i][5]!=='') question.setFeedbackForIncorrect(FormApp.createFeedback().setText(data[i][5]).build());
  if(data[i][optionStart-1]!=='') question.setRequired(data[i][optionStart-1]);
}
function addOptions(i) {
  const arr = [];
  for(let j=optionStart;j<optionStart+optionLength;j++) {
    if(ss.getRange(i+1, j+1, 1, 1).getValue()==='') continue;
    if(ss.getRange(i+1, j+1, 1, 1).getBackground()===correctColor) arr.push(question.createChoice(data[i][j], true));
    else arr.push(question.createChoice(data[i][j], false));
  }
  question.setChoices(arr);
}
function formatVisual() {
  question.setAlignment(FormApp.Alignment.CENTER).setWidth(600);
}