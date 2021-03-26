
var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}
var calID = function() {return ss().getRange("B1").getValue();}
var color = function() {return 7;} //represents red

function onInstall(e) {
    onOpen(e);
}
function onOpen(e) {
    // let menu = UI.createMenu("Forms"); Used as standalone
    let menu = SpreadsheetApp.getUi().createAddonMenu(); //used with congunction with google marketplace
    menu.addItem("Init", "init").addToUi();
    menu.addItem("Delete all events", "deleteAll").addToUi(); //delete events from calendar
    menu.addItem("Fill in lessons", "fillInDate").addToUi(); //add dates beside events in Spreadsheet
    menu.addItem("Add events to calendar", "addToCalendar").addToUi(); //add spreadsheet events into calendar
}
function init() {
    ss().clear();
    ss().getRange("A1").setValue("Cal ID:");
    ss().getRange("B1").setValue("335396990@gapps.yrdsb.ca");

    ss().getRange("A2").setValue("Title");
    ss().getRange("B2").setValue("Start");
    ss().getRange("C2").setValue("End");
    ss().getRange("D2").setValue("Color");
}
function deleteAll() {
    // in the form of MM/DD/YYYY, else its just present day
    let start = new Date('03/01/2021'), end = new Date('04/01/2021');

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
        }

        //things I need to do:
        //check to see if event already appears in calendar? -> how duplicates are handled
    }
}

/*
Useful event calendar functions:
-getColor()
-getTitle()
-getDescription()
-getStartTime()
-getEndTime()

color bank:
1 | pale blue
2 | pale green
3 | dark blue
4 | red
5 | yellow
6 | orange
7 | cyan
8 | gray
9 | blue
10 | green
11 | red
*/