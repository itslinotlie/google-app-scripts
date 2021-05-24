/*
  Still a work in progress
*/

var ss = function () { return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() }
var sa = function () { return SpreadsheetApp.getActiveSpreadsheet(); }
var yrdsbID = function () { return "3kirfs2j6u2dob2n587e5om1gs@group.calendar.google.com"; } //calendar for YRDSB holidays -> took a long time of debugging, but you have to add the calendar to your personal calendar for this to work
const colorBank = [
  "#7986cb",
  "#33b679",
  "#8e24aa",
  "#e67c73",
  "#f6bf26",
  "#f4511e",
  "#039be5",
  "#616161",
  "#3f51b5",
  "#0b8043",
  "#d50000"
];
const basicStyling = ["HCENTER", "VCENTER"];
let startDeleteDate = new Date("01/09/2020"), endDeleteDate = new Date("01/06/2021");
let startCalDate, endCalDate;
let delColor = 11, headerSize = 3; //default color is red
let defCalID, curCalID, spreadID;
//abbreviations
let SA = SpreadsheetApp, UI = SA.getUi();
var highlight = "#29d57b", topBackground = "#faefcf", bottomBackground = "#f0f8ff", outline = "#000000"; //default is: green, tan, blue, black respectively

function onInstall(e) {
  onOpen(e);
}
function onOpen(e) {
  let menu = UI.createAddonMenu(); //used with congunction with google marketplace
  menu.addItem("Initialize Program", "initialize").addToUi();
  menu.addSubMenu(UI.createMenu("Spreadsheet helpers")
    .addItem("Fill in Lessons Dates", "fillDate") //add dates beside events in Spreadsheet (requires a first date)
    .addItem("Show YRDSB Holidays", "holiday") //lists YRDSB holidays from a date range
  ).addToUi();
  menu.addSubMenu(UI.createMenu("Calendar helpers")
    .addItem("Delete Generated Events", "deleteAll") //deletes events from calendar
  ).addToUi();
  menu.addSubMenu(UI.createMenu("Calendar <-> Spreadsheet")
    .addItem("Add Events to Calendar", "addToCalendar") //add spreadsheet events into calendar
    .addItem("List Calendar to Spreadsheet", "addToSpreadsheet") //list calendar events from a range onto spreadsheet
  ).addToUi();
}
function update() {
  let sheetName = ss().getName();
  if (sa().getSheetByName("Settings") == null) sa().insertSheet("Settings");
  sa().setActiveSheet(sa().getSheetByName("Settings"));
  resizeSheet(100, 5);
  ss().setColumnWidths(1, ss().getMaxColumns(), 175);
  ss().setRowHeights(1, ss().getMaxRows(), 25);

  if (ss().getRange("B3").getValue() !== "") startDeleteDate = ss().getRange("B3").getValue();
  if (ss().getRange("B4").getValue() !== "") endDeleteDate = ss().getRange("B4").getValue();
  if (ss().getRange("B5").getValue() !== "") delColor = ss().getRange("B5").getValue();

  ss().getRange("A3").setValue("Delete Start Date*:");
  ss().getRange("A4").setValue("Delete End Date*:");
  ss().getRange("A5").setValue("Delete Color Tag: (enter #)");
  ss().getRange("A6").setValue("*The range is [Start, end)");
  ss().getRange("B2").setValue("DD/MM/YYYY");
  ss().getRange("B3").setValue(startDeleteDate);
  ss().getRange("B4").setValue(endDeleteDate);
  ss().getRange("B5").setValue(delColor);

  for (let i = 0; i < 11; i++) {
    ss().getRange("C" + (5 + i)).setValue(i + 1);
    ss().getRange("C" + (5 + i) + ":" + "C" + (5 + i)).setBackground(colorBank[i]);
  }

  ss().getRange("D3").setValue("Default Calendar ID:");
  if (ss().getRange("E3").getValue() !== "") defCalID = ss().getRange("E3").getValue();

  ss().getRange("D5").setValue("Calendar ID (to Spreadsheet)");
  ss().getRange("D6").setValue("Calendar Start Date:");
  ss().getRange("D7").setValue("Calendar End Date:");
  spreadID = ss().getRange("E5").getValue();
  startCalDate = ss().getRange("E6").getValue();
  endCalDate = ss().getRange("E7").getValue();

  sa().setActiveSheet(sa().getSheetByName(sheetName));
}
function initialize() {
  let sheetName = ss().getName();
  ss().clear();
  //dimensions
  let desRow = 100, desCol = 3;
  resizeSheet(desRow, desCol);
  const columnSize = [300, 150, 150];
  for(let i=0;i<columnSize.length;i++) ss().setColumnWidth(i+1, columnSize[i]);
  ss().setRowHeights(1, desRow, 25);
  ss().setRowHeight(headerSize, 50);
  ss().getRange(1, headerSize, 25);
  //colors
  ss().getRange("A1:"+char(desCol)+(headerSize-1)).setBackground(topBackground);
  ss().getRange("A"+headerSize+":"+char(desCol)+headerSize).setBackground(highlight);
  ss().getRange("A"+(headerSize+1)+":"+char(desCol)+desRow).setBackground(bottomBackground);
  //border
  ss().getRange(headerSize+":"+headerSize).setBorder(true, false, true, false, false, false, outline, SA.BorderStyle.SOLID_MEDIUM);
  //misc
  setStrategy(headerSize+":"+headerSize, basicStyling);
  setStrategy("A"+(headerSize+1)+":"+char(desCol)+desRow, ["WRAP", "VTOP", "HLEFT"]);
  setStrategy("A"+(headerSize+1)+":A"+desRow, ["WRAP", "VTOP", "HLEFT"]);
  setStrategy("B"+(headerSize+1)+":"+char(desCol)+desRow, ["HCENTER", "VCENTER"]);
  setFormat(["A1:A1", "A3:C3"], ["bold", 12]);

  //content
  ss().getRange("A1").setValue("Cal ID:"); //335396990@gapps.yrdsb.ca
  ss().getRange("B2").setValue("Dates are in the form of DD/MM/YYYY");

  ss().getRange("A" + headerSize).setValue("Event Title");
  ss().getRange("B" + headerSize).setValue("Event Date");
  ss().getRange("C" + headerSize).setValue("Calendar Color");

  update(); holiday();
  sa().setActiveSheet(sa().getSheetByName(sheetName));
}
function deleteAll() {
  update(); let calendar;
  curCalID = ss().getRange("B1").getValue();
  if (curCalID.includes("@")) calendar = CalendarApp.getCalendarById(curCalID);
  else calendar = CalendarApp.getCalendarById(defCalID);

  if (calendar === null) {
    UI.alert("Something went wrong ):", "I did not detect a Calendar ID. Either move your Spreadsheet or add a Calendar ID in the Settings page 'Default Calendar ID'", UI.ButtonSet.OK);
    return;
  }

  let events = calendar.getEvents(startDeleteDate, endDeleteDate);
  for (let i = 0; i < events.length; i++) {
    if (events[i].getColor() == delColor) events[i].deleteEvent();
  }
}
function fillDate() {
  let sheetName = ss().getName();
  sa().setActiveSheet(sa().getSheetByName("Holidays"));
  var holiday = ss().getDataRange().getValues(), idx = 2;
  sa().setActiveSheet(sa().getSheetByName(sheetName));
  var data = ss().getDataRange().getValues();

  //I figured out the black magic.  IF(WEEKDAY<CELL>=6, 3, 1)
  //checks if the prev day was Friday, if so, add 3 days so it becomes monday, else add one day
  //that is some 100x engineer kind of code if I do say so myself. stackoverflow orz

  //needs to have the first row's date filled in
  ss().getRange("C" + (headerSize + 1)).setBackground(colorBank[ss().getRange("C" + (headerSize + 1)).getValue() - 1]);
  for (let i = headerSize + 2; i <= data.length; i++) {
    let cell = '(B' + (i - 1) + ")";
    let start = '=' + cell + '+IF(WEEKDAY' + cell + '=6,3,1)';
    ss().getRange("C1").setValue(start); //need a dummy cell
    let date = new Date(ss().getRange("C1").getValue());
    ss().getRange("C1").setValue(date); //to prevent infinte loop with the formula, I need to hardcode the date
    let srt = new Date(holiday[idx][1]), end = new Date(holiday[idx][2]);

    while (idx < holiday.length && +date >= +srt) {
      while (+srt <= +date && +date < +end) {
        ss().getRange("A2").setValue(ss().getRange("C1").getValue());
        ss().getRange("C1").setValue("=A2+IF(WEEKDAY(A2)=6,3,1)");
        date = new Date(ss().getRange("C1").getValue());
      }
      if (idx === holiday.length - 1) break;
      srt = new Date(holiday[++idx][1]); end = new Date(holiday[idx][2]);
    }
    ss().getRange("B" + i).setValue(date);
    ss().getRange("C" + i).setBackground(colorBank[ss().getRange("C" + i).getValue() - 1]);
  }
  ss().getRange("C1").setValue("");
}
function addToCalendar() {
  update(); let calendar;
  curCalID = ss().getRange("B1").getValue();
  if (curCalID !== "") calendar = CalendarApp.getCalendarById(curCalID);
  else calendar = CalendarApp.getCalendarById(defCalID);
  let data = ss().getDataRange().getValues();

  for (let i = 3; i < data.length; i++) {
    let title = data[i][0];
    let start = data[i][1];
    let color = data[i][2];
    ss().getRange(i + 1, 3).setBackground(colorBank[ss().getRange(i + 1, 3).getValue() - 1]); //1-index -> row, col
    if (start != null) {
      let event = calendar.createEvent(title, new Date(start), new Date(start));
      event.setColor(color);
      event.setAllDayDate(start);
    }
  }
}
function addToSpreadsheet() {
  update();
  if (sa().getSheetByName("CalendarEvents") == null) sa().insertSheet("CalendarEvents");
  sa().setActiveSheet(sa().getSheetByName("CalendarEvents"));
  ss().setColumnWidths(1, 3, 250);
  let calendar = CalendarApp.getCalendarById(spreadID);
  let events = calendar.getEvents(startCalDate, endCalDate);

  ss().getRange("A1").setValue("Title");
  ss().getRange("B1").setValue("Start");
  ss().getRange("C1").setValue("Color");
  for (let i = 0; i < events.length; i++) {
    ss().getRange("A" + (2 + i)).setValue(events[i].getTitle());
    ss().getRange("B" + (2 + i)).setValue(events[i].getStartTime());
    // ss().getRange("C"+(2+i)).setBackground(colorBank[events[i].getColor()-1]);
    ss().getRange("C" + (2 + i)).setValue(events[i].getColor());
  }
}
function holiday() {
  let sheetName = ss().getName();
  if (sa().getSheetByName("Holidays") == null) sa().insertSheet("Holidays");
  sa().setActiveSheet(sa().getSheetByName("Holidays"));
  //dimensions
  let desRow = 200, desCol = 4, headerSize = 2;
  resizeSheet(desRow, desCol);
  const columnSize = [300, 150, 150, 150];
  for(let i=0;i<columnSize.length;i++) ss().setColumnWidth(i+1, columnSize[i]);
  ss().setRowHeights(1, desRow, 25);
  ss().setRowHeight(1, 30); ss().setRowHeight(headerSize, 50);
  //colors
  ss().getRange("A1:"+char(desCol)+(headerSize-1)).setBackground(topBackground);
  ss().getRange("A"+headerSize+":"+char(desCol)+headerSize).setBackground(highlight);
  ss().getRange("A"+(headerSize+1)+":"+char(desCol)+desRow).setBackground(bottomBackground);
  //border
  ss().getRange(headerSize+":"+headerSize).setBorder(true, false, true, false, false, false, outline, SA.BorderStyle.SOLID_MEDIUM);
  //misc
  setStrategy("A1:"+char(desCol)+headerSize, basicStyling);
  setStrategy("A"+(headerSize+1)+":"+char(desCol)+desRow, ["WRAP", "VTOP", "HLEFT"]);
  setStrategy("B"+(headerSize+1)+":"+char(desCol)+desRow, ["HCENTER", "VCENTER"]);
  //init row, init col, # of rows, # of cols
  let range = ss().getRange(3, 1, ss().getMaxRows() - 3, ss().getMaxColumns());
  range.clear();
  setFormat(["A2:D2"], ["bold", 12]);

  //content
  let start = new Date("01/09/2020"), end = new Date("01/06/2021");
  ss().getRange("A1").setValue("Holiday type");
  ss().getRange("B1").setValue("Event Start");
  ss().getRange("C1").setValue("Event End");
  ss().getRange("A2").setValue("Start Date:");
  ss().getRange("C2").setValue("End Date:");
  if (ss().getRange("B2").getValue() === "") ss().getRange("B2").setValue(start.toLocaleDateString().substring(0, start.toLocaleString().indexOf(' ')));
  if (ss().getRange("D2").getValue() === "") ss().getRange("D2").setValue(end.toLocaleDateString().substring(0, end.toLocaleString().indexOf(' ')));

  start = ss().getRange("B2").getValue();
  end = ss().getRange("D2").getValue();

  let calendar = CalendarApp.getCalendarById(yrdsbID());
  let events = calendar.getEvents(start, end);

  for (let i = 0; i < events.length; i++) {
    ss().getRange("A" + (i + 3)).setValue(events[i].getTitle());
    ss().getRange("B" + (i + 3)).setValue(events[i].getAllDayStartDate());
    ss().getRange("C" + (i + 3)).setValue(events[i].getAllDayEndDate());
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
function setFormat(range, type) {
  for (let i = 0; i < range.length; i++) {
    for (let j = 0; j < type.length; j++) {
      if (type[j] === "bold") ss().getRange(range[i]).setFontWeight("bold");
      else if (!isNaN(type[j])) ss().getRange(range[i]).setFontSize(type[j]);
    }
  }
}
function setStrategy(range, type) {
  for (let i = 0; i < type.length; i++) {
    if (type[i] === "WRAP") ss().getRange(range).setWrapStrategy(SA.WrapStrategy.WRAP);
    else if (type[i] === "FLOW") ss().getRange(range).setWrapStrategy(SA.WrapStrategy.OVERFLOW);
    else if (type[i] === "CLIP") ss().getRange(range).setWrapStrategy(SA.WrapStrategy.CLIP);
    else if (type[i] === "VTOP") ss().getRange(range).setVerticalAlignment("top");
    else if (type[i] === "VCENTER") ss().getRange(range).setVerticalAlignment("middle");
    else if (type[i] === "HLEFT") ss().getRange(range).setHorizontalAlignment("left");
    else if (type[i] === "HCENTER") ss().getRange(range).setHorizontalAlignment("center");
  }
}
function char(x) {
  return String.fromCharCode(64+x);
}
/*
updated color bank:
1  #7986cb | pale blue
2  #33b679 | pale green
3  #8e24aa | purple
4  #e67c73 | salmon
5  #f6bf26 | yellow
6  #f4511e | orange
7  #039be5 | cyan
8  #616161 | gray
9  #3f51b5 | dark blue
10 #0b8043 | green
11 #d50000 | red
*/