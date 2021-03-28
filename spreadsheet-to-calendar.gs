
var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}
var sa = function() {return SpreadsheetApp.getActiveSpreadsheet();}
var calID = function() {return ss().getRange("B1").getValue();}
var color = function() {return 7;} //represents red
var  yrdsbID = function() {return "3kirfs2j6u2dob2n587e5om1gs@group.calendar.google.com";} //calendar for YRDSB holidays -> took a long time of debugging, but you have to add the calendar to your personal calendar for this to work
const colorBank = ["#a4bdfc","#7ae7bf","#bdadff","#ff887c","#fbd75b","#ffb878","#46d6db","#e1e1e1","#5484ed","#51b749", "#dc2127"]

function onInstall(e) {
    onOpen(e);
}
function onOpen(e) {
    // let menu = UI.createMenu("Forms"); Used as standalone
    let menu = SpreadsheetApp.getUi().createAddonMenu(); //used with congunction with google marketplace
    menu.addItem("Init", "init").addToUi();
    menu.addItem("Delete generated events", "deleteAll").addToUi(); //delete events from calendar
    menu.addItem("Fill in lessons", "fillInDate").addToUi(); //add dates beside events in Spreadsheet
    menu.addItem("Add events to calendar", "addToCalendar").addToUi(); //add spreadsheet events into calendar
    menu.addItem("Show YRDSB holidays", "holiday").addToUi();
}
function init() {
    let sheetName = ss().getName();
    ss().clear();
    ss().setColumnWidths(1, 3, 175); //set column widths to be bigger
    ss().getRange("A1").setValue("Cal ID:");
    ss().getRange("B1").setValue("335396990@gapps.yrdsb.ca");

    ss().getRange("A2").setValue("Title");
    ss().getRange("B2").setValue("Start");
    ss().getRange("C2").setValue("End (this doesn't matter)");
    ss().getRange("D2").setValue("Color");
    ss().getRange("E1").setValue("Date is in the form of DD/MM/YYYY");

    if(sa().getSheetByName("Settings")==null) sa().insertSheet("Settings");
    sa().setActiveSheet(sa().getSheetByName("Settings"));
    ss().clear();
    ss().setColumnWidths(1, 3, 175); //set column widths to be bigger

    ss().getRange("A1").setValue("Delete Events");
    ss().getRange("A3").setValue("Start date:");
    ss().getRange("A4").setValue("End date:");
    ss().getRange("A5").setValue("Delete Color Tag:");
    ss().getRange("B2").setValue("DD/MM/YYYY");
    ss().getRange("B3").setValue("01/09/2020");
    ss().getRange("B4").setValue("01/06/2021");
    ss().getRange("B5:B5").setBackground(colorBank[10]); //red
    for(let i=0;i<11;i++) {
        ss().getRange("C"+(5+i)+":"+"C"+(5+i)).setBackground(colorBank[i]);
        ss().getRange("C"+(5+i)).setValue((i+1));
    }
    sa().setActiveSheet(sa().getSheetByName(sheetName));;
}
function deleteAll() {
    let sheetName =ss().getName();
    sa().setActiveSheet(sa().getSheetByName("Settings"));
    let start = ss().getRange("B3").getValue();
    let end = ss().getRange("B4").getValue();
    sa().setActiveSheet(sa().getSheetByName(sheetName));

    let calendar = CalendarApp.getCalendarById(calID());
    let events = calendar.getEvents(start, end);

    //this feels very sketchy, but it works... (I would have thought that by deleting, indexes get shifted, but I guess you don't have to account for that)
    for(let i=0;i<events.length;i++) {
        if(events[i].getColor()==color()) events[i].deleteEvent();
    }
}
function fillInDate() {
    // var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = ss().getDataRange().getValues();

    //this is some google sheet formula black magic but I think whats happening is
    //it checks if the day is a weekend, if true, returns 3 else returns 1 -> from the IF(x, a, b) thing, but I dont 
    //know why in the world it works. I'm chalking this up on the fact that this is a sheets formula and not actual code...

    //needs to have the first row's date filled in
    for(let i=4;i<20;i++) {
        //skips over weekends (dont ask me how) and writes the next day on the right
        //because Calendar events need [date, date+1) or something....
        let cell  = '(B'+(i-1)+")";
        let start   = '='+cell+'+IF(WEEKDAY'+cell+'=6,3,1)';
        let end     = '=B'+i+"+1";

        //just need to check if current date = holiday/pa day, then skip

        ss().getRange("B"+i).setValue(start);
        ss().getRange("C"+i).setValue(end);
    }
}
function addToCalendar() {
    let data = ss().getDataRange().getValues();
    let calendar = CalendarApp.getCalendarById(calID());

    for(let i=2;i<data.length;i++) {
        let title = data[i][0];
        let start = data[i][1], end = data[i][2];
        let color = data[i][3];
        if(start!=null && end!=null) {
            let event = calendar.createEvent(title, new Date(start), new Date(end));
            event.setColor(color);
            event.setAllDayDate(start);
        }

        //things I need to do:
        //check to see if event already appears in calendar? -> how duplicates are handled
    }
}
function holiday() {
    let sheetName = ss().getName();
    if(sa().getSheetByName("Holidays")==null) sa().insertSheet("Holidays");
    sa().setActiveSheet(sa().getSheetByName("Holidays"));
    ss().setColumnWidths(1, 3, 175); //set column widths to be bigger

    let calendar = CalendarApp.getCalendarById(yrdsbID());
    let start = new Date("09/01/2020"), end = new Date("06/01/2021");
    let events = calendar.getEvents(start, end);

    ss().getRange("A1").setValue("Holiday type");
    ss().getRange("B1").setValue("Start date (event is all day)");
    for(let i=0;i<events.length;i++) {
        ss().getRange("A"+(i+3)).setValue(events[i].getTitle());
        ss().getRange("B"+(i+3)).setValue(events[i].getAllDayStartDate());
    }
    sa().setActiveSheet(sa().getSheetByName(sheetName));
}

/*
Useful event calendar functions:
-getColor()
-getTitle()
-getDescription()
-getStartTime()
-getEndTime()

color bank:
1  #a4bdfc | pale blue
2  #7ae7bf | pale green
3  #bdadff | dark blue
4  #ff887c | red
5  #fbd75b | yellow
6  #ffb878 | orange
7  #46d6db | cyan
8  #e1e1e1 | gray
9  #5484ed | blue
10 #51b749 | green
11 #dc2127 | red
*/