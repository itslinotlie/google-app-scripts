//customizable variables (for when I create a global settings page)
var optionLength = 5;
var desRow = 21, desCol = optionLength+10;
var headerSize = 6;
var tagLength = 10;
var folderID = "";
var correctColor = "#29d57b";
var highlight = "#29d57b", topBackground = "#faefcf", bottomBackground = "#f0f8ff", outline = "#000000"; //default is: green, tan, blue, black respectively

//variables that could change, but would be a pain to do so
var charEnd = char(desCol);
var optionStart = 2, optionEnd = optionStart+optionLength; //0-indexed
var charOptionStart = char(optionEnd-optionLength+1), charOptionEnd = char(optionEnd);
var tagStart = optionStart+7;

//the array of header positioning (rows are flexible to be changed)
const header = [ //cell-letter, header-name, col (1-indexed), width respectively
  ["A", "Question Type",                1, 150], 
  ["B", "Question",                     2, 200],
  // <option(s)> in between ^^ and vv
  [char(optionEnd+1), "Points",         optionEnd+1],
  [char(optionEnd+2), "Tag",            optionEnd+2],
  [char(optionEnd+3), "Required?",      optionEnd+3],
  [char(optionEnd+4), "Other?",         optionEnd+4],
  [char(optionEnd+5), "Instructions",   optionEnd+5, 200],
  [char(optionEnd+6), "Correct Text",   optionEnd+6, 200],
  [char(optionEnd+7), "Incorrect Text", optionEnd+7, 200],
  [char(optionEnd+8), "URL / ID",       optionEnd+8]
];
const options = [ //supported Form question types (not the actual naming GAS uses, but my simplification of them)
  "MC", "CHECKBOX", "MCGRID", "CHECKGRID", "SHORTANSWER", 
  "PARAGRAPH", "DROPDOWN", "PAGEBREAK", "HEADER", "IMAGE", "IMAGE-DRIVE", "VIDEO"
];
const basicStyling = ["HCENTER", "VCENTER"]; //styling used for majority of the spreadsheet
const bool = ["TRUE", "FALSE"];
const tagNameArr = []; //tag names, default is Tag <Number>
const formSettings = [
  ["One Response per User?", false],
  ["Can Edit Response?",     false],
  ["Collects Email?",        false],
  ["Progress Bar?",          false],
  ["Link to Respond Again?", false],
  ["Publishing Summary?",    false]
]

//abbreviations, I don't even know how this is legal, but it works
let SA = SpreadsheetApp, UI = SA.getUi(), IT = FormApp.ItemType;

//column numbers based on header array. If header[x][1] values are changed, these need to be changed too (1-indexed)
let requiredNumber = find("Required?"), otherNumber = find("Other?"), instructionsNumber = find("Instructions"),
  incorrectTextNumber = find("Incorrect Text"), correctTextNumber = find("Correct Text"),
  url = find("URL / ID"), pointsNumber = find("Points"), tagNumber = find("Tag");

//other column numbers (1-indexed)
let titleNumber = 1, descriptionNumber = 1;

//letter number combo
let alertCell = "H5", randomOptionCell = "H6";
let alertBool = true, randomOptionBool = false;
let tagGap = 14;

//workaround to Authmode.NONE because publication requirements
var sa = function() {return SpreadsheetApp.getActiveSpreadsheet();}
var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}
var data = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues()}
var question;

var sheetName;

//google script triggers for when certain events happen
function onInstall(e) {
  onOpen(e);
}
function onOpen(e) {
  // let menu = UI.createMenu("Forms"); Used as standalone
  let menu = SpreadsheetApp.getUi().createAddonMenu(); //used with congunction with google marketplace
  menu.addItem("Initilize Spreadsheet", "createTemplate").addToUi();
  menu.addItem("Create Google Form", "createForm").addToUi();
  menu.addItem("Link to Documentation", "linkDoc").addToUi();
  menu.addItem("Update Settings", "update").addToUi();
  menu.addItem("Delete later", "createSetting").addToUi();
  alertBool = sa().getSheetByName("Settings").getRange(alertCell).getValue();
  if(alertBool) //sometimes this gives errors (but code still runs), sometimes it doesn't /shrug
    SA.getActiveSpreadsheet().toast("Remember to check the GitHub documentation or YouTube video for any help/clarifications. Have a good day :)", "Hello fellow human being");
  // keeping vvv iin case I do need the e.authMode and I forget that its a thing and end up on stackoverflow for hours
  // if(e.authMode !== ScriptApp.AuthMode.NONE && ss().getRange(alertCell).getValue()===true) //sometimes this gives errors (but code still runs), sometimes it doesn't /shrug
}
function onEdit(e) { //alerts user if they checked GRID question type, the row below should be void
  let row = e.range.getRow(), col = char(e.range.getColumn()), range = col+row;
  alertBool = sa().getSheetByName("Settings").getRange(alertCell).getValue();
  if(col!=="A" || row<=headerSize || !alertBool) return;
  if(ss().getRange(range).getValue()==="MCGRID" || ss().getRange(range).getValue()==="CHECKGRID") 
    UI.alert("Friendly Reminder", "Remember that the cell below "+range+" should be the columns for the "
    +ss().getRange(range).getValue()+" and nothing else. You can turn alerts off by setting the B4 cell to FALSE", UI.ButtonSet.OK);
}

//updates tagNames, updates Tag column validations
function update() {
  sheetName = ss().getName(); //used to get back to the previous sheet

  //if the settings sheet does not exist, create it
  if(sa().getSheetByName("Settings")==null) createSetting();
  sa().setActiveSheet(sa().getSheetByName("Settings")); //need to switch active sheet to settings to get the tag names

  //initial settings
  folderID     = ss().getRange("B5").getValue();
  optionLength = ss().getRange("B6").getValue(); optionEnd = optionStart+optionLength;
  tagLength    = ss().getRange("B7").getValue(); tagStart = optionStart+7;
  desRow       = ss().getRange("B8").getValue(); desCol = optionLength+10;
  charOptionStart = char(optionEnd-optionLength+1); charOptionEnd = char(optionEnd);
  charEnd = char(desCol);

  //colour settings
  correctColor     = ss().getRange("D5").getBackground();
  highlight        = ss().getRange("D6").getBackground();
  topBackground    = ss().getRange("D7").getBackground();
  bottomBackground = ss().getRange("D8").getBackground();
  outline          = ss().getRange("D9").getBackground();

  //boolean settings
  for(let i=0;i<formSettings.length;i++) formSettings[i][1] = ss().getRange("H"+(5+i)).getValue();

  //misc. settings
  alertBool        = ss().getRange(alertCell).getValue();
  randomOptionBool = ss().getRange(randomOptionCell).getValue();

  tagNameArr.length = 0; //dont know how to reset a js array, but heard this works
  for(let i=0;i<tagLength;i++) {
    var cl = char(5+(~~(i/5))); //some magical integer division from js
    if(ss().getRange(cl+(i%5+tagGap)).getValue()==='') ss().getRange(cl+(i%5+tagGap)).setValue("Tag "+(i+1));
    tagNameArr.push(ss().getRange(cl+(i%5+14)).getValue());
  }

  //validation for the tag column
  sa().setActiveSheet(sa().getSheetByName(sheetName));
  for (let i=0;i<sa().getSheets().length;i++) {
    if(sa().getSheets()[i].getName()==="Settings") continue;
    sa().setActiveSheet(sa().getSheets()[i]);
    for (let j=0;j<header.length;j++) {
      if(header[j][1]==="Tag") setValidation(header[j][0]+(headerSize+1)+":"+header[j][0]+ss().getMaxRows(), tagNameArr);
    }
    //renaming tag names in normal sheets
    for(let i=0;i<tagLength;i++) {
      let tagChar = char(7+(~~(i/5))*2);
      ss().getRange(tagChar+(i%5+1)).setValue(tagNameArr[i]);
    }
  }
  sa().setActiveSheet(sa().getSheetByName(sheetName));
}

//adds/removes columns to match the desired size
function resizeSheet(rowWant, colWant) {
  let curRow = ss().getMaxRows(), curCol = ss().getMaxColumns();
  if(curRow!==rowWant) //Exception: Invalid argument is thrown if you .inserRowsAfter(X, 0)
    curRow>rowWant? ss().deleteRows(rowWant+1, curRow-rowWant):ss().insertRowsAfter(curRow-1, rowWant-curRow);
  if(curCol!==colWant)
    curCol>colWant? ss().deleteColumns(colWant+1, curCol-colWant):ss().insertColumnsAfter(curCol-1, colWant-curCol);
}
function createSetting() {
  sheetName = ss().getName(); //used to get back the previous active sheet

  //temporarily just so I dont have to delte the settings sheet whenever I test stuff
  let newSheet = sa().getSheetByName("Settings");
  if(newSheet!==null) sa().deleteSheet(newSheet);
  newSheet = sa().insertSheet("Settings");
  sa().setActiveSheet(newSheet);

  //IMPORTANT
  //need to make note that the changes will only be applied to new spreadsheets

  // let newSheet = sa().insertSheet("Settings"); //creates a new sheet called Settings
  // sa().setActiveSheet(newSheet); //newSheet is now the active spreadsheet

  let settingRow = 20, settingCol = 11;
  resizeSheet(settingRow, settingCol);
  ss().setRowHeights(1, settingRow, 25); ss().setColumnWidths(1, settingCol, 175);
  setStrategy("A1:"+char(settingCol)+settingRow, basicStyling);
  setStrategy("A1:"+char(settingCol)+settingRow, ["WRAP"]);

  ss().getRange("A1").setValue("Global Settings for all your Sheets");
  
  ss().getRange("A3").setValue("Initial Settings");
  ss().getRange("A5").setValue("Folder ID");
  ss().getRange("A6").setValue("Option Length");     ss().getRange("B6").setValue(optionLength);
  ss().getRange("A7").setValue("Tag Amount");        ss().getRange("B7").setValue(tagLength);
  ss().getRange("A8").setValue("Initial Sheet Row"); ss().getRange("B8").setValue(desRow);
  ss().getRange("A9").setValue("Initial Sheet Col"); ss().getRange("B9").setValue(desCol);

  ss().getRange("C3").setValue("Colour Settings");   
  ss().getRange("C5").setValue("Correct Colour");    ss().getRange("D5").setBackground(correctColor);
  ss().getRange("C6").setValue("Highlight Colour");  ss().getRange("D6").setBackground(highlight);
  ss().getRange("C7").setValue("Top Background");    ss().getRange("D7").setBackground(topBackground);
  ss().getRange("C8").setValue("Bottom Background"); ss().getRange("D8").setBackground(bottomBackground);
  ss().getRange("C9").setValue("Outline");           ss().getRange("D9").setBackground(outline);
  ss().getRange("D5").setBorder(true, true, true, true, false, false, "#fbaed2", SA.BorderStyle.SOLID_THICK);
  for(let i=5;i<=9;i++) ss().getRange("D"+i).setBorder(true, true, true, true, false, false, "#fbaed2", SA.BorderStyle.SOLID_THICK); //range notation doesn't appear to work so for loop it is

  ss().getRange("E12").setValue("Tag Naming");
  for(let i=0;i<tagLength;i++) {
    var cl = char(5+(~~(i/5))); //some magical integer division from js
    tagNameArr.push("Tag "+(i+1));
    ss().getRange(cl+(i%5+tagGap)).setValue(tagNameArr[i]); //default tag naming
    //if I were to insert a filler number (0) after Tag 1, then the checking algorithm would detect something
    //and would then create google form with 0 questions, so not sure what to do (could create a "filler variable" and check for that)
  }

  ss().getRange("E3").setValue("Boolean Settings");
  for(let i=0;i<formSettings.length;i++) {
    ss().getRange("E"+(5+i)).setValue(formSettings[i][0]);
    ss().getRange("F"+(5+i)).setValue(formSettings[i][1]);
    setValidation("F"+(5+i), bool);
  } 
  setStrategy("E5:E10", ["HLEFT"]);

  ss().getRange("G3").setValue("Misc. Settings");
  ss().getRange("G5").setValue("I Want Alerts");     ss().getRange("H5").setValue(alertBool);
  ss().getRange("G6").setValue("Randomize OPTIONS"); ss().getRange("H6").setValue(randomOptionBool);
  setValidation("H5:H6", bool);

  sa().setActiveSheet(sa().getSheetByName(sheetName)); //UI now refocuses back to the original spreadsheet
}

//first menu item, creates template
function createTemplate() {
  let x = data().length; //checker to see if data is valid is "out of bounds" for empty spreadsheet, but js is weird and I need use variable rather than data().length
  alertBool = sa().getSheetByName("Settings").getRange(alertCell).getValue();
  if(alertBool && (x!=1) && (data()[0][1]!=="" || ss().getLastRow()>6) && data()[3][1]) { //have title? have info in the bottom?
    let response = UI.alert("Are you sure? (all information will be cleared)", UI.ButtonSet.YES_NO);
    if(response===UI.Button.NO) return;
  }
  update();

  resizeSheet(desRow, desCol);

  ss().clear(); //clears formatting
  ss().setRowHeights(1, desRow, 21); ss().setColumnWidths(1, desCol, 100); //resize cells to default cell size
  ss().getRange(1, 1, desRow, desCol).setDataValidation(null); //clears data formatting so you dont need to create a new sheet
  while(ss().getImages().length>0) ss().getImages()[0].remove(); //removes all images in the sheet

  //info to fill in/use
  ss().getRange("A1").setValue("Form Title:");
  ss().getRange("A2").setValue("Form Desciption:");
  ss().getRange("C1").setValue("Folder ID:");
  ss().getRange("D1").setValue(folderID);
  ss().getRange("C2").setValue("Public URL:");
  ss().getRange("C3").setValue("Private URL:");

  //header info
  ss().setRowHeight(headerSize, 50);
  for (let i=0;i<header.length;i++) {
    ss().getRange(header[i][0]+headerSize).setValue(header[i][1]);
  } ss().getRange(charOptionStart+headerSize+":"+charOptionEnd+headerSize).setValue("OPTION");

  ss().getRange("E2").setValue("Number of questions from those tags you want to have->");
  for(let i=0;i<tagLength;i++) {
    let tagChar = char(7+(~~(i/5))*2);
    ss().getRange(tagChar+(i%5+1)).setValue(tagNameArr[i]);
  }

  //validations
  for (let i=0;i<header.length;i++) {
    if(header[i][1]==="Question Type") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+desRow, options);
    if(header[i][1]==="Required?") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+desRow, bool);
    if(header[i][1]==="Other?") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+desRow, bool);
    if(header[i][1]==="Tag") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+desRow, tagNameArr);
  }

  //cell width formatting
  for (let i=0;i<header.length;i++) {
    if(header[i][1]==="Question Type" ||
        header[i][1]==="Question" ||
        header[i][1]==="Instructions" || 
        header[i][1]==="Correct Text" ||
        header[i][1]==="Incorrect Text") ss().setColumnWidth(header[i][2], header[i][3]);
  } ss().setColumnWidths(optionStart+1, optionLength, 175)

  //general formatting
  setStrategy("A1:"+charEnd+desRow, ["WRAP", "VTOP", "HLEFT"]);
  for (let i=0;i<header.length;i++) {
    let x = header[i][0]+(headerSize+1)+":"+header[i][0]+desRow;
    if(header[i][1]==="Question Type" ||
        header[i][1]==="Points" ||
        header[i][1]==="Required?" ||
        header[i][1]==="Other?") setStrategy(x, basicStyling);
    if(header[i][1]==="URL / ID") setStrategy(x, ["CLIP"]);
  } 
  
  //dunno how to categorize these
  setStrategy(headerSize+":"+headerSize, basicStyling);
  // setStrategy("D1:D3", ["CLIP"]);
  // setStrategy("B4:B5", basicStyling);

  //colors
  ss().getRange("A1:"+charEnd+desRow).setBackground(topBackground);
  ss().getRange(headerSize+":"+headerSize).setBackground(highlight);
  ss().getRange("A"+(headerSize+1)+":"+charEnd+desRow).setBackground(bottomBackground);

  //misc
  ss().setFrozenRows(headerSize); ss().setFrozenColumns(2);
  // setFormat(["A1:A2", "C1:C3", "E1:E3", "G1:G3", headerSize+":"+headerSize], "bold");
  // setFormat(["A1:A2", "C1:C3", headerSize+":"+headerSize], 12);
  //top, left, bottom, right, vertical, horizontal, color, style)
  ss().getRange(headerSize+":"+headerSize).setBorder(true, false, true, false, false, false, outline, SA.BorderStyle.SOLID_MEDIUM);

  // happy message :)
  alertBool = sa().getSheetByName("Settings").getRange(alertCell).getValue();
  if(alertBool)
    SA.getActiveSpreadsheet().toast("Remember to check the GitHub documentation or YouTube video for any help/clarifications. Have a good day :)", "Hello fellow human being");
}

//second menu item, creates the form
function createForm() {
  update();

  //subtle plug (:
  alertBool = sa().getSheetByName("Settings").getRange(alertCell).getValue();
  if(alertBool)
    SA.getActiveSpreadsheet().toast("Remember to check the GitHub documentation or YouTube video for any help/clarifications. Have a good day :)", "Hello fellow human being");
  
  let row = ss().getDataRange().getNumRows();
  var form = FormApp.create("Untitled form"); //has to be initialized
  let formID = form.getId();
  let file = DriveApp.getFileById(formID);

  //setting form info to spreadsheet
  let publicUrl = form.getPublishedUrl();
  let privateUrl = form.getEditUrl();
  ss().getRange("D2").setValue(publicUrl);
  ss().getRange("D3").setValue(privateUrl);

  //moving form to the folder (if possible)
  if(folderID!=='') file.moveTo(DriveApp.getFolderById(folderID));

  //filling in form info
  file.setName(data()[0][titleNumber]);
  form.setTitle(data()[0][titleNumber]);
  form.setDescription(data()[1][descriptionNumber]);
  form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled

  // boolean info
  correctColor = ss().getRange("B3").getBackground();
  for(let i=0;i<formSettings.length;i++) {
    if(formSettings[i][0]==="One Response per User?") form.setLimitOneResponsePerUser(formSettings[i][1]);
    if(formSettings[i][0]==="Can Edit Response?")     form.setAllowResponseEdits(formSettings[i][1]);
    if(formSettings[i][0]==="Collects Email?")        form.setCollectEmail(formSettings[i][1]); //reveals email address at the top of the form, allows you to send a copy to yourself at the bottom
    if(formSettings[i][0]==="Progress Bar?")          form.setProgressBar(formSettings[i][1]);
    if(formSettings[i][0]==="Link to Respond Again?") form.setShowLinkToRespondAgain(formSettings[i][1]);
    if(formSettings[i][0]==="Publishing Summary?")    form.setPublishingSummary(formSettings[i][1]); //reveals question distribution, but no answers
  }

  //tag questions
  let cntTag = [], tagRnd, tagArray = [];
  for(let i=0;i<tagLength;i++) {
    cntTag.push(data()[i%5][tagStart+(~~(i/5)*2)]);
    tagRnd = tagRnd || data()[i%5][tagStart+(~~(i/5)*2)]>0;
  }

  //adding questions to form
  for (let i=headerSize;i<row;i++) {
    let x = data()[i][0]; 
    if(x==='') continue;
    if(tagRnd) {
      let flag = false;
      for(let j=0;j<tagLength;j++) {
        let idx = findTagFromWord(data()[i][tagNumber-1]);
        if(cntTag[idx]>0) {
          flag = true;
          break;
        }
      }
      if(flag) tagArray.push(i);
      continue;
    }
    if(x==="MC") question = form.addMultipleChoiceItem();
    else if(x==="CHECKBOX") question = form.addCheckboxItem();
    else if(x==="MCGRID") question = form.addGridItem();  
    else if(x==="CHECKGRID") question = form.addCheckboxGridItem();
    else if(x==="SHORTANSWER") question = form.addTextItem();
    else if(x==="PARAGRAPH") question = form.addParagraphTextItem();
    else if(x==="DROPDOWN") question = form.addListItem();
    else if(x==="PAGEBREAK") question = form.addPageBreakItem();
    else if(x==="HEADER") question = form.addSectionHeaderItem(); //these are stackable, but don't look the greatest
    else if(x==="IMAGE") question = form.addImageItem().setImage(UrlFetchApp.fetch(data()[i][url-1])); //imageItem's helptext dont show in Forms
    else if(x==="IMAGE-DRIVE") question = form.addImageItem().setImage(DriveApp.getFileById(data()[i][url-1]));
    else if(x==="VIDEO") question = form.addVideoItem().setVideoUrl(data()[i][url-1]);
    setUpQuestion(i);
  }
  // shuffle(arrRnd); 
  shuffle(tagArray);
  if(tagRnd) {
    for(let i=0;i<tagArray.length;i++) {
      let x = data()[tagArray[i]][tagNumber-1], idx = findTagFromWord(x), flag = true;
      if(x==='') continue;
      for (let j=0;j<cntTag.length;j++) {
        if(cntTag[j]>0) flag = false;
      } if(flag) break;
      if(cntTag[idx]>0) {
        if(data()[tagArray[i]][0]==="MC") question = form.addMultipleChoiceItem();
        if(data()[tagArray[i]][0]==="CHECKBOX") question = form.addCheckboxItem();
        if(data()[tagArray[i]][0]==="SHORTANSWER") question = form.addTextItem();
        if(data()[tagArray[i]][0]==="PARAGRAPH") question = form.addParagraphTextItem();
        cntTag[idx]--;
        setUpQuestion(tagArray[i]);
      }
    }
  } 
}
const choices = [IT.CHECKBOX, IT.MULTIPLE_CHOICE, IT.LIST];
const visual = [IT.IMAGE, IT.VIDEO];
const mix = [IT.CHECKBOX, IT.MULTIPLE_CHOICE,
  IT.PARAGRAPH_TEXT, IT.TEXT, IT.LIST,
];
const twoD = [IT.CHECKBOX_GRID, IT.GRID];
function setUpQuestion(i) {
  if(data()[i][titleNumber-1]       !=='') question.setTitle(data()[i][titleNumber-1]);
  if(data()[i][instructionsNumber-1]!=='') question.setHelpText(data()[i][instructionsNumber-1]);
  let type = question.getType();
  for (let j=0;j<visual.length;j++) { //Visuals (Image + Video)
    if(type===visual[j]) formatVisual();
  }
  for (let j=0;j<choices.length;j++) { //Adding options + feedback
    if(type===choices[j]) {
      addOptions(i);
      if(data()[i][otherNumber-1]!=='' && type!==IT.LIST) question.showOtherOption(data()[i][otherNumber-1]);
      if(data()[i][correctTextNumber-1]!=='') question.setFeedbackForCorrect(FormApp.createFeedback().setText(data()[i][correctTextNumber-1]).build());
      if(data()[i][incorrectTextNumber-1]!=='') question.setFeedbackForIncorrect(FormApp.createFeedback().setText(data()[i][incorrectTextNumber-1]).build());
    }
  }
  for (let j=0;j<mix.length;j++) { //Adding points + setting required
    if(type===mix[j]) {
      if(data()[i][pointsNumber-1]  !=='') question.setPoints(data()[i][pointsNumber-1]);
      if(data()[i][requiredNumber-1]!=='') question.setRequired(data()[i][requiredNumber-1]);
    }
  }
  for (let j=0;j<twoD.length;j++) {
    if(type===twoD[j]) {
      addGrid(i);
      if(data()[i][requiredNumber-1]!=='') question.setRequired(data()[i][requiredNumber-1]);
    }
  }
}
function addOptions(i) {
  const arr = [];
  for(let j=optionStart;j<optionStart+optionLength;j++) {
    if(ss().getRange(i+1, j+1, 1, 1).getValue()==='') continue;
    if(ss().getRange(i+1, j+1, 1, 1).getBackground()===correctColor) arr.push(question.createChoice(data()[i][j], true));
    else arr.push(question.createChoice(data()[i][j], false));
  }
  if(arr.length===0) return;
  randomOptionBool = sa().getSheetByName("Settings").getRange(randomOptionCell);
  if(randomOptionBool) shuffle(arr);
  question.setChoices(arr);
}
function addGrid(x) {
  for(let i=x;i<=x+1;i++) {
    const arr = [];
    for(let j=optionStart;j<optionStart+optionLength;j++) {
      if(ss().getRange(i+1, j+1, 1, 1).getValue()==='') continue;
      arr.push(data()[i][j]);
    }
    if(data()[i][0]==="MCGRID" || data()[i][0]==="CHECKGRID") question.setColumns(arr);
    else question.setRows(arr);
  }
}
function formatVisual() {
  question.setAlignment(FormApp.Alignment.CENTER).setWidth(600);
}
//helper functions
function char(x) {
  return String.fromCharCode(64+x);
}
function findTagFromWord(x) {
  update();
  for(let i=0;i<tagNameArr.length;i++) {
    if(tagNameArr[i]===x) return i;
  } return 314;
}
function setSize(range, size, letter) {
  for (let i=0;i<range.length;i++) {
    if(letter==="R") ss().setRowHeight(range[i], size);
    else if(letter==="C") ss().setColumnWidth(range[i], size);
  }
}
function setStrategy(range, type) {
  for (let i=0;i<type.length;i++) {
    if(type[i]==="WRAP") ss().getRange(range).setWrapStrategy(SA.WrapStrategy.WRAP);
    else if(type[i]==="CLIP") ss().getRange(range).setWrapStrategy(SA.WrapStrategy.CLIP);
    else if(type[i]==="VTOP") ss().getRange(range).setVerticalAlignment("top");
    else if(type[i]==="VCENTER") ss().getRange(range).setVerticalAlignment("middle");
    else if(type[i]==="HLEFT") ss().getRange(range).setHorizontalAlignment("left");
    else if(type[i]==="HCENTER") ss().getRange(range).setHorizontalAlignment("center");
  }
}
function setValidation(range, list) {
  ss().getRange(range).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false).requireValueInList(list, true).build());
}
function setFormat(range, type) {
  for (let i=0;i<range.length;i++) {
    if(type==="bold") ss().getRange(range[i]).setFontWeight("bold");
    else if(!isNaN(type)) ss().getRange(range[i]).setFontSize(type);
  }
}
function shuffle(arr) { //Fisher-Yates shuffle
  for (let i=arr.length-1;i>0;i--) {
    let j = Math.floor(Math.random()*(i+1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}
function find(x) {
  for (let i=0;i<header.length;i++) {
    if(header[i][1]===x) return header[i][2];
  }
}
function linkDoc() { //copied from https://support.google.com/docs/thread/16869830?hl=en&msgid=17047454 (beyond the scope of this project)
  var url = "https://github.com/itslinotlie/google-app-scripts/blob/master/spreadsheet-to-form.md";
  var html = HtmlService.createHtmlOutput(
    '<html>'
  + '  <script>'
  + '    window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  + '    var a = document.createElement("a");a.href="'+url+'";a.target="_blank";'
  + '    if(document.createEvent) {'
  + '      var event = document.createEvent("MouseEvents");'
  + '      if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  + '      event.initEvent("click",true,true); a.dispatchEvent(event);'
  + '    } else{a.click()}'
  + '    close();'
  + '  </script>'
  // Offer URL as clickable link in case above code fails.
  + '  <body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  + '  <script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  + '</html>').setWidth(90).setHeight(1);
  UI.showModalDialog(html, "Redirecting ... (Your browser may prevent you from seeing the link)");
}