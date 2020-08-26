/*
General knowledge:
- Arrays are 0-indexed
- ss.getRange(row, col, numRow, numCol) is 1-indexed
*/
function onOpen() {
  let menu = SpreadsheetApp.getUi().createMenu("Forms");
  menu.addItem("Initilize Spreadsheet", "createTemplate").addToUi();
  menu.addItem("Create Google Form", "createForm").addToUi();
}
function createTemplate() {
  //Current sheet in the spreadsheet
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.clearContents();
  
  //Setting default spreadsheet size
  if(ss.getMaxColumns()>10) ss.deleteColumns(11, ss.getMaxColumns()-10);
  if(ss.getMaxRows()>10) ss.deleteRows(11, ss.getMaxRows()-10);
  
  //Info
  ss.getRange("A1").setValue("Form Name:");
  ss.getRange("A2").setValue("Form Desciption");
  ss.getRange("A3").setValue("Question Type");
  ss.getRange("B3").setValue("Question");
  ss.getRange("C1").setValue("Folder ID:");
  ss.getRange("C3").setValue("Instructionns");
  ss.getRange("D3:G3").setValue("OPTION");
  ss.getRange("E1").setValue("Public URL:");
  
  //Cell logic
  ss.getRange("A4:A10").setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(
      ["CHOICE", "TEXT"], true).build())
}
function createForm() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let val = ss.getDataRange().getValues();
  
  let folder = DriveApp.getFolderById(val[0][3]);
  let form = FormApp.create(val[0][1]); //creates and opens a form with title in [0][1]
  let formID = form.getId();
  
  //setting form info to spreadsheet
  let url = form.getPublishedUrl();
  ss.getRange("F1").setValue(url);

  //moves form to the folder
  let file = DriveApp.getFileById(formID);
  file.moveTo(folder);
  
  //Filling in form 
  let dataRange = ss.getDataRange(); //2d array with 
  let row = dataRange.getNumRows(), col = dataRange.getNumColumns();
  let data = dataRange.getValues();

  form.setDescription(val[1][1]);
  form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled
  
  for (let i=0;i<row;i++) {
    let x = data[i][0]; 
    
    if(x=='') continue; 
    if(x=="CHOICE") {
      const arr = [];
      let question = form.addMultipleChoiceItem();
      arr.push(question.createChoice("D4", true));
      arr.push(question.createChoice("F4", false));
      question.setTitle(data[i][1]);
      question.setHelpText(data[i][2]);//.setRequired(true);
      
      for(let j=3;j<7;j++) {
        if(ss.getRange(i+1, j+1, 1, 1).getValue()=='') continue;
        if(ss.getRange(i+1, j+1, 1, 1).getBackground()=="#80ff00") arr.push(question.createChoice(data[i][j], true));
        else arr.push(question.createChoice(data[i][j], false));
      }
      question.setChoices(arr);
    }
    else if(x=="TEXT") {
      let question = form.addTextItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(true);
    }
  }
}