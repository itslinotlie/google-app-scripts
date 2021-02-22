//customizable variables (for when I create a global settings page)
var optionLength = 5;
var desRow = 21, desCol = optionLength+10;
var headerSize = 6;
var tagLength = 15;
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
  [char(optionEnd+1), "Points",   optionEnd+1],
  [char(optionEnd+2), "Tag", optionEnd+2],
  [char(optionEnd+3), "Required?",      optionEnd+3],
  [char(optionEnd+4), "Other?",         optionEnd+4],
  [char(optionEnd+5), "Instructions",   optionEnd+5, 200],
  [char(optionEnd+6), "Correct Text",   optionEnd+6, 200],
  [char(optionEnd+7), "Incorrect Text", optionEnd+7, 200],
  [char(optionEnd+8), "URL / ID", optionEnd+8]
];
const options = [ //supported Form question types (not the actual naming GAS uses, but my simplification of them)
  "MC", "CHECKBOX", "MCGRID", "CHECKGRID", "SHORTANSWER", 
  "PARAGRAPH", "DROPDOWN", "PAGEBREAK", "HEADER", "IMAGE", "IMAGE-DRIVE", "VIDEO"
];
const basicStyling = ["HCENTER", "VCENTER"]; //styling used for majority of the spreadsheet
const bool = ["TRUE", "FALSE"];
const tagNameArr = []; //names of the tags

//abbreviations
let SA = SpreadsheetApp, UI = SA.getUi(), IT = FormApp.ItemType;

//numbers based on header info
let req = find("Required?"), other = find("Other?"), instruc = find("Instructions"),
  inc = find("Incorrect Text"), cor = find("Correct Text"), url = find("URL / ID"),
  pnt = find("Points"), taggy = find("Tag");

//workaround to Authmode.NONE
var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}
var data = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues()}
var question;

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
  if(e.authMode !== ScriptApp.AuthMode.NONE)
   SA.getActiveSpreadsheet().toast("Remember to check the GitHub documentation or YouTube video for any help/clarifications. Have a good day :)", "Hello fellow human being");
}
function onEdit(e) { //alerts user if they checked GRID question type, the row below should be void
  let row = e.range.getRow(), col = char(e.range.getColumn()), range = col+row;
  if(col!=="A" || row<=headerSize) return;
  if((ss().getRange(range).getValue()==="MCGRID" || ss().getRange(range).getValue()==="CHECKGRID") && data()[3][1]) 
    UI.alert("Friendly Reminder", "Remember that the cell below "+range+" should be the columns for the "
    +ss().getRange(range).getValue()+" and nothing else. You can turn alerts off by setting the B4 cell to FALSE", UI.ButtonSet.OK);
}
function setupTag() {
  for(let i=0;i<tagLength;i++) {
    var cl = char(tagStart+(~~(i/5)*2)); //some magical integer division from js
    ss().getRange(cl+(i%5+1)).setValue("Tag "+(i+1)); //default tag naming
    //if I were to insert a filler number (0) after Tag 1, then the checking algorithm would detect something
    //and would then create google form with 0 questions, so not sure what to do (could create a "filler variable" and check for that)
  }
  update();
}
function update() {
  tagNameArr.length = 0; //dont know how to reset a js array, but heard this works
  for(let i=0;i<tagLength;i++) {
    var cl = char(tagStart+(~~(i/5)*2)); //magical integer division from js orz
    tagNameArr.push(ss().getRange(cl+(i%5+1)).getValue());
  }
  //validation for the tag column
  for (let i=0;i<header.length;i++) {
    if(header[i][1]==="Tag") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+ss().getMaxRows(), tagNameArr);
  }
}
function createTemplate() {
  let x = data().length; //checker to see if data is valid is "out of bounds" for empty spreadsheet, but js is weird and I need use variable rather than data().length
  if((x!=1) && (data()[0][1]!=="" || ss().getLastRow()>6) && data()[3][1]) { //have title? have info in the bottom?
    let response = UI.alert("Are you sure? (all information will be cleared)", UI.ButtonSet.YES_NO);
    if(response===UI.Button.NO) return;
  }
  // setting up spreadsheet dimensions
  let curRow = ss().getMaxRows(), curCol = ss().getMaxColumns();
  if(curRow!==desRow) //Exception: Invalid argument is thrown if you .inserRowsAfter(X, 0)
    curRow>desRow? ss().deleteRows(desRow+1, curRow-desRow):ss().insertRowsAfter(curRow-1, desRow-curRow);
  if(curCol!==desCol)
    curCol>desCol? ss().deleteColumns(desCol+1, curCol-desCol):ss().insertColumnsAfter(curCol-1, desCol-curCol);
  curRow = ss().getMaxRows(); curCol = ss().getMaxColumns();

  ss().clear(); //clears formatting
  ss().setRowHeights(1, curRow, 21); ss().setColumnWidths(1, curCol, 100); //resize cells to default cell size
  ss().getRange(1, 1, desRow, desCol).setDataValidation(null); //clears data formatting so you dont need to create a new sheet
  while(ss().getImages().length>0) ss().getImages()[0].remove(); //removes all images in the sheet

  //info to fill in/use
  ss().getRange("A1").setValue("Form Title:");
  ss().getRange("A2").setValue("Form Desciption:");
  ss().getRange("A3").setValue("Highlight Color");
  ss().getRange("A4").setValue("Alerts");
  setValidation("B4", bool); ss().getRange("B4").setValue("TRUE");
  ss().getRange("A5").setValue("Randomize OPTIONS");
  setValidation("B5", bool); ss().getRange("B5").setValue("FALSE");
  ss().getRange("C1").setValue("Folder ID:");
  // \o> Edit Me <o/
  // ss.getRange("D1").setValue("1D2yMTtKfq9ey5awuTbEiHXViCHDYgejH");
  ss().getRange("C2").setValue("Public URL:");
  ss().getRange("C3").setValue("Private URL:");
  ss().getRange("C4").setValue("Random subset of questions based on category");
  ss().getRange("C4:C5").merge();
  setStrategy("C4", ["WRAP"]);
  setupTag();

  //various boolean fields
  ss().getRange("E1").setValue("One Response per User?");
  ss().getRange("E2").setValue("Can Edit Response?");
  ss().getRange("E3").setValue("Collects Email?");
  setValidation("F1:F3", bool);
  ss().getRange("G1").setValue("Progress Bar?");
  ss().getRange("G2").setValue("Link to Respond Again?");
  ss().getRange("G3").setValue("Publishing Summary?");
  setValidation("H1:H3", bool);

  //random subset of questions
  ss().getRange("E4").setValue("# of MC:");
  ss().getRange("E5").setValue("# of SHORTANSWER:");
  ss().getRange("G4").setValue("# of CHECKBOX:");
  ss().getRange("G5").setValue("# of PARAGRAPH:");

  //kinda related to ^^^
  // let src = UrlFetchApp.fetch("https://imgur.com/QSzRPRL.png").getContent();
  // ss().insertImage(Utilities.newBlob(src, "image/png", "aName"), 4, 4, 70, -2);
  // ss().getRange("D4:D5").merge();
  setStrategy("F1:F5", basicStyling); setStrategy("H1:H5", basicStyling);
  
  //header info
  ss().setRowHeight(headerSize, 50);
  for (let i=0;i<header.length;i++) {
    ss().getRange(header[i][0]+headerSize).setValue(header[i][1]);
  } ss().getRange(charOptionStart+headerSize+":"+charOptionEnd+headerSize).setValue("OPTION");

  //validations
  for (let i=0;i<header.length;i++) {
    if(header[i][1]==="Question Type") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+curRow, options);
    if(header[i][1]==="Required?") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+curRow, bool);
    if(header[i][1]==="Other?") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+curRow, bool);
    if(header[i][1]==="Tag") setValidation(header[i][0]+(headerSize+1)+":"+header[i][0]+curRow, tagNameArr);
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
  setStrategy("A1:"+charEnd+curRow, ["WRAP", "VTOP", "HLEFT"]);
  for (let i=0;i<header.length;i++) {
    let x = header[i][0]+(headerSize+1)+":"+header[i][0]+curRow;
    if(header[i][1]==="Question Type" ||
        header[i][1]==="Points" ||
        header[i][1]==="Required?" ||
        header[i][1]==="Other?") setStrategy(x, basicStyling);
    if(header[i][1]==="URL / ID") setStrategy(x, ["CLIP"]);
  } 
  
  //dunno how to categorize these
  setStrategy(headerSize+":"+headerSize, basicStyling);
  setStrategy("D1:D3", ["CLIP"]);
  setStrategy("B4:B5", basicStyling);

  //colors
  ss().getRange("A1:"+charEnd+curRow).setBackground(topBackground);
  ss().getRange(headerSize+":"+headerSize).setBackground(highlight);
  ss().getRange("A"+(headerSize+1)+":"+charEnd+curRow).setBackground(bottomBackground);
  ss().getRange("B3").setBackground(correctColor);

  //misc
  ss().setFrozenRows(headerSize); ss().setFrozenColumns(2);
  setFormat(["A1:A2", "C1:C3", "E1:E3", "G1:G3", headerSize+":"+headerSize], "bold");
  setFormat(["A1:A2", "C1:C3", headerSize+":"+headerSize], 12);
  //top, left, bottom, right, vertical, horizontal, color, style)
  ss().getRange("B3").setBorder(true, true, true, true, false, false, outline, SA.BorderStyle.SOLID_MEDIUM);
  ss().getRange(headerSize+":"+headerSize).setBorder(true, false, true, false, false, false, outline, SA.BorderStyle.SOLID_MEDIUM);

  // happy message :)
  SA.getActiveSpreadsheet().toast("Remember to check the GitHub documentation or YouTube video for any help/clarifications. Have a good day :)", "Hello fellow human being");
}
function createForm() {
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
  if(data()[0][3]!=='') {
    let folder = DriveApp.getFolderById(data()[0][3]);
    file.moveTo(folder);
  }

  //filling in form info
  file.setName(data()[0][1]);
  form.setTitle(data()[0][1]);
  form.setDescription(data()[1][1]);
  form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled

  // boolean info
  correctColor = ss().getRange("B3").getBackground();
  if(data()[0][5]!=='') form.setLimitOneResponsePerUser(data()[0][5]);
  if(data()[1][5]!=='') form.setAllowResponseEdits(data()[1][5]);
  if(data()[2][5]!=='') form.setCollectEmail(data()[2][5]); //reveals email address at the top of the form, allows you to send a copy to yourself at the bottom
  if(data()[0][7]!=='') form.setProgressBar(data()[0][7]);
  if(data()[1][7]!=='') form.setShowLinkToRespondAgain(data()[1][7]);
  if(data()[2][7]!=='') form.setPublishingSummary(data()[2][7]); //reveals question distribution, but no answers

  //tag questions
  let cntTag = [], tagRnd, tagArray = [];
  for(let i=0;i<tagLength;i++) {
    cntTag.push(data()[i%5][tagStart+(~~(i/5)*2)]);
    tagRnd = tagRnd || data()[i%5][tagStart+(~~(i/5)*2)]>0;
  }

  //random category of questions
  let rnd, arrRnd = [], cntRnd = [0, 0, 0, 0]; //MC, CB, SA, PG respectively
  if(data()[3][5]!=='') cntRnd[0] = data()[3][5];
  if(data()[3][7]!=='') cntRnd[1] = data()[3][7];
  if(data()[4][5]!=='') cntRnd[2] = data()[4][5];
  if(data()[4][7]!=='') cntRnd[3] = data()[4][7];
  for (let i=0;i<4;i++) rnd = rnd || cntRnd[i]>0;

  //adding questions to form
  for (let i=headerSize;i<row;i++) {
    let x = data()[i][0]; 
    if(x==='') continue;
    if(tagRnd) {
      let flag = false;
      for(let j=0;j<tagLength;j++) {
        let idx = findTagFromWord(data()[i][taggy-1]);
        if(cntTag[idx]>0) {
          flag = true;
          break;
        }
      }
      if(flag) tagArray.push(i);
      continue;
    }
    else if(rnd) {
      if(x==="MC" && cntRnd[0]>0
        || x==="CHECKBOX" && cntRnd[1]>0
        || x==="SHORTANSWER" && cntRnd[2]>0
        || x==="PARAGRAPH" && cntRnd[3]>0
      ) arrRnd.push(i);
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
  shuffle(arrRnd); shuffle(tagArray);
  if(tagRnd) {
    for(let i=0;i<tagArray.length;i++) {
      let x = data()[tagArray[i]][taggy-1], idx = findTagFromWord(x), flag = true;
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
  } else if(rnd) {
    for (let i=0;i<arrRnd.length;i++) {
      let x = data()[arrRnd[i]][0], flag = true;
      if(x==='') continue;
      for (let j=0;j<cntRnd.length;j++) {
        if(cntRnd[j]>0) flag = false;
      } if(flag) break;
      if(x==="MC" && cntRnd[0]>0) {question = form.addMultipleChoiceItem(); cntRnd[0]--;}
      else if(x==='CHECKBOX' && cntRnd[1]>0) {question = form.addCheckboxItem(); cntRnd[1]--;}
      else if(x==='SHORTANSWER' && cntRnd[2]>0) {question = form.addTextItem(); cntRnd[2]--;}
      else if(x==='PARAGRAPH' && cntRnd[3]>0) {question = form.addParagraphTextItem(); cntRnd[3]--;}
      setUpQuestion(arrRnd[i]);
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
  if(data()[i][1]!=='') question.setTitle(data()[i][1]);
  if(data()[i][instruc-1]!=='') question.setHelpText(data()[i][instruc-1]);
  let type = question.getType();
  for (let j=0;j<visual.length;j++) { //Visuals (Image + Video)
    if(type===visual[j]) formatVisual();
  }
  for (let j=0;j<choices.length;j++) { //Adding options + feedback
    if(type===choices[j]) {
      addOptions(i);
      if(data()[i][other-1]!=='' && type!==IT.LIST) question.showOtherOption(data()[i][other-1]);
      if(data()[i][cor-1]!=='') question.setFeedbackForCorrect(FormApp.createFeedback().setText(data()[i][cor-1]).build());
      if(data()[i][inc-1]!=='') question.setFeedbackForIncorrect(FormApp.createFeedback().setText(data()[i][inc-1]).build());
    }
  }
  for (let j=0;j<mix.length;j++) { //Adding points + setting required
    if(type===mix[j]) {
      if(data()[i][pnt-1]!=='') question.setPoints(data()[i][pnt-1]);
      if(data()[i][req-1]!=='') question.setRequired(data()[i][req-1]);
    }
  }
  for (let j=0;j<twoD.length;j++) {
    if(type===twoD[j]) {
      addGrid(i);
      if(data()[i][req-1]!=='') question.setRequired(data()[i][req-1]);
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
  if(ss().getRange("B5")) shuffle(arr);
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