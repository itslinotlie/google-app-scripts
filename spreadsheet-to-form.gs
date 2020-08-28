/*
General knowledge:
- Arrays are 0-indexed
- ss.getRange(row, col, numRow, numCol) is 1-indexed
*/
//global variables
var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //getActiveSheet() == current sheet opened in the spreadsheet
var form = FormApp.create("Untitled form"); //had to be initialized

var dataRange = ss.getDataRange(); //2d array dimensions
var data = dataRange.getValues(); //2d array with values
var question; //simplifies methods if this is global

var optionStart = 7; //0-indexed
var optionLength = 10;
var desRow = 20, desCol = 20;
var correctColor = "#00ff00"; //default neon green highlight

function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu("Forms");
  menu.addItem("Initilize Spreadsheet", "createTemplate").addToUi();
  menu.addItem("Create Google Form", "createForm").addToUi();
}
function createTemplate() {
  ss.clearContents(); ss.clearFormats();

  //setting up spreadsheet dimensions
  let curRow = ss.getMaxRows(), curCol = ss.getMaxColumns();
  if(curRow!==desRow) //Exception: Invalid argument is thrown if you .inserRowsAfter(X, 0)
    curRow>desRow? ss.deleteRows(desRow+1, curRow-desRow):ss.insertRowsAfter(curRow-1, desRow-curRow);
  if(curCol!==desCol)
    curCol>desCol? ss.deleteColumns(desCol+1, curCol-desCol):ss.insertColumnsAfter(curCol-1, desCol-curCol);
  curRow = ss.getMaxRows(); curCol = ss.getMaxColumns();

  //predefined-info
  ss.getRange("A1").setValue("Form Title:");
  ss.getRange("A2").setValue("Form Desciption:");
  ss.getRange("C1").setValue("Folder ID:");
  //delete the following line if you plan on copying this file
  ss.getRange("D1").setValue("1D2yMTtKfq9ey5awuTbEiHXViCHDYgejH"); //delete when in production
  ss.getRange("E1").setValue("Public URL:");
  ss.getRange("E2").setValue("Private URL:");

  ss.getRange("A3").setValue("Question Type");
  ss.getRange("B3").setValue("Question");
  ss.getRange("C3").setValue("Instructions");
  ss.getRange("D3").setValue("Points");
  ss.getRange("E3").setValue("Correct Text");
  ss.getRange("F3").setValue("Incorrect Text");
  ss.getRange("G3").setValue("Required?");
  let charStart = String.fromCharCode(65+optionStart)+"3"; //3 represents row # (1-indexed)
  let charEnd = String.fromCharCode(65+optionStart+optionLength-1)+"3"; //Have to -1 for some reason? (check later)
  ss.getRange(charStart+":"+charEnd).setValue("OPTION");

  //Cell logic
  ss.setFrozenColumns(2);
  ss.setFrozenRows(3);
  const options = ["MC", "CHECKBOX", "SHORTANSWER", "PARAGRAPH", "PAGEBREAK", "HEADER"];
  const bool = ["TRUE", "FALSE"]; 
  let req = String.fromCharCode(65+optionStart-1); 
  ss.getRange("A4:A"+curRow).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(options, true).build()); //https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder#setAllowInvalid(Boolean)
  ss.getRange(req+"4:"+req+curRow).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(bool, true).build());

  //very "hacky" solution for "locking" cells from being edited
  ss.getRange("G1").setValue("H1 and H2 are locked");
  const blank = [""];  
  ss.getRange("H1:H2").setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(blank, false).build());
  const tmp = [ss.getRange("A1").getDisplayValue()];
  ss.getRange("A1").setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(tmp, false).build());
}
function createForm() {
  let row = dataRange.getNumRows(), col = dataRange.getNumColumns(); //col is never used

  //setting form info to spreadsheet
  let publicUrl = form.getPublishedUrl();
  let privateUrl = form.getEditUrl();
  ss.getRange("F1").setValue(publicUrl);
  ss.getRange("F2").setValue(privateUrl);

  //moving form to the folder
  let folder = DriveApp.getFolderById(data[0][3]);
  let formID = form.getId();
  let file = DriveApp.getFileById(formID);
  file.moveTo(folder);

  //filling in form info
  form.setTitle(data[0][1]);
  form.setDescription(data[1][1]);
  form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled

  //To ADD:
  /*
  - one sub per person
  - option for required or not
  - allow user to pick color for right answer (i.e. place highlight color in this cell)
  */

  for (let i=0;i<row;i++) {
    let x = data[i][0]; 

    if(x==='') continue; 
    if(x==="MC") {
      question = form.addMultipleChoiceItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(data[i][6]);
      setUpQuestion(i);
    }
    else if(x==="CHECKBOX") {
      question = form.addCheckboxItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(data[i][6]);
      setUpQuestion(i);
    }
    // You can set point values for short responses, but how does it work logisticially?
    else if(x==="SHORTANSWER") {
      question = form.addTextItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(data[i][6]);
      addPoints(i);
    }
    else if(x==="PARAGRAPH") {
      question = form.addParagraphTextItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(data[i][6]);
      addPoints(i);
    }
    else if(x==="PAGEBREAK") {
      form.addPageBreakItem().setTitle(data[i][1]).setHelpText(data[i][2]);
    }
    else if(x==="HEADER") { //these are stackable, but don't look the greatest
      form.addSectionHeaderItem().setTitle(data[i][1]).setHelpText(data[i][2]);
    }
  }
}
function setUpQuestion(i) {
  addOptions(i);
  addPoints(i);
  if(data[i][3]!=='') question.setPoints(data[i][3]);
  if(data[i][4]!=='') question.setFeedbackForCorrect(FormApp.createFeedback().setText(data[i][4]).build());
  if(data[i][5]!=='') question.setFeedbackForIncorrect(FormApp.createFeedback().setText(data[i][5]).build());
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
function addPoints(i) {
  if(data[i][3]!=='') question.setPoints(data[i][3]);
}