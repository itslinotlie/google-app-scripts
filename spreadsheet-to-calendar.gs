
var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}
var sa = function() {return SpreadsheetApp.getActiveSpreadsheet();}
var calID = function() {return ss().getRange("B1").getValue();}
var  yrdsbID = function() {return "3kirfs2j6u2dob2n587e5om1gs@group.calendar.google.com";} //calendar for YRDSB holidays -> took a long time of debugging, but you have to add the calendar to your personal calendar for this to work
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
]

let startDate = new Date("01/09/2020"), endDate = new Date("01/06/2022");

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
    // ss().getRange("C2").setValue("End (this doesn't matter)");
    ss().getRange("D2").setValue("Color");
    ss().getRange("E1").setValue("Date is in the form of DD/MM/YYYY");

    if(sa().getSheetByName("Settings")==null) sa().insertSheet("Settings");
    sa().setActiveSheet(sa().getSheetByName("Settings"));
    ss().clear();
    ss().setColumnWidths(1, 3, 175); //set column widths to be bigger

    ss().getRange("A1").setValue("Delete Events");
    ss().getRange("A3").setValue("Start date:");
    ss().getRange("A4").setValue("End date:");
    ss().getRange("A5").setValue("Delete Color Tag: (enter #)");
    ss().getRange("B2").setValue("DD/MM/YYYY");
    ss().getRange("B3").setValue(startDate);
    ss().getRange("B4").setValue(endDate);
    ss().getRange("B5").setValue("11");
    ss().getRange("B5:B5").setBackground(colorBank[10]); //red

    for(let i=0;i<11;i++) {
        ss().getRange("C"+(5+i)).setValue(i+1);
        ss().getRange("C"+(5+i)+":"+"C"+(5+i)).setBackground(colorBank[i]);
    }

    holiday();

    sa().setActiveSheet(sa().getSheetByName(sheetName));
}
function deleteAll() {
    let sheetName = ss().getName();
    sa().setActiveSheet(sa().getSheetByName("Settings"));
    let start = ss().getRange("B3").getValue();
    let end = ss().getRange("B4").getValue();
    let colorNum = ss().getRange("B5").getValue();
    sa().setActiveSheet(sa().getSheetByName(sheetName));

    let calendar = CalendarApp.getCalendarById(calID());
    let events = calendar.getEvents(start, end);

    //this feels very sketchy, but it works... (I would have thought that by deleting, indexes get shifted, but I guess you don't have to account for that)
    for(let i=0;i<events.length;i++) {
        if(events[i].getColor()==colorNum) events[i].deleteEvent();
    }
}
function fillInDate() {
    let sheetName = ss().getName();
    sa().setActiveSheet(sa().getSheetByName("Holidays"));
    var holiday = ss().getDataRange().getValues(), idx = 2;
    sa().setActiveSheet(sa().getSheetByName(sheetName));
    var data = ss().getDataRange().getValues();

    //I figured out the black magic.  IF(WEEKDAY<CELL>=6, 3, 1)
    //checks if the prev day was Friday, if so, add 3 days so it becomes monday, else add one day
    //that is some 100x engineer kind of code if I do say so myself. stackoverflow orz

    let something = 7;
    //needs to have the first row's date filled in
    for(let i=4;i<=data.length;i++) { //thought it had to be i<length but /shrug
        let cell  = '(B'+(i-1)+")";
        let start = '='+cell+'+IF(WEEKDAY'+cell+'=6,3,1)';
        ss().getRange("D1").setValue(start); //need a dummy cell
        let date = new Date(ss().getRange("D1").getValue());
        ss().getRange("D1").setValue(date); //to prevent infinte loop with the formula
        let tmp = new Date(holiday[idx][1]);
        while(idx<holiday.length && +date>=+tmp) {
            ss().getRange("B"+something).setValue(idx+" | "+holiday.length+" = "+holiday[idx][1]);
            ss().getRange("C"+something).setValue(date);
            ss().getRange("D"+(something)).setValue(tmp);
            ss().getRange("E"+(something++)).setValue(+date>=+tmp);
            ss().getRange("D1").setValue(date);
            if(+date===+tmp) {
                ss().getRange("C1").setValue(ss().getRange("D1").getValue());
                ss().getRange("D1").setValue("=C1+IF(WEEKDAY(C1)=6,3,1)");
            }
            date = new Date(ss().getRange("D1").getValue());
            if(idx===holiday.length-1) break;
            tmp = new Date(holiday[++idx][1]);
        }
        ss().getRange("B"+i).setValue(date);
    }
    ss().getRange("C1").setValue("");
    ss().getRange("D1").setValue("");
}
function addToCalendar() {
    let data = ss().getDataRange().getValues();
    let calendar = CalendarApp.getCalendarById(calID());

    for(let i=2;i<data.length;i++) {
        let title = data[i][0];
        let start = data[i][1];//, end = data[i][2];
        let color = data[i][3];
        if(start!=null && end!=null) {
            let event = calendar.createEvent(title, new Date(start), new Date(start));
            event.setColor(color);
            event.setAllDayDate(start);
        }
    }
}
function holiday() {
    let sheetName = ss().getName();
    if(sa().getSheetByName("Holidays")==null) sa().insertSheet("Holidays");
    sa().setActiveSheet(sa().getSheetByName("Holidays"));
    ss().setColumnWidths(1, 3, 175); //set column widths to be bigger

    let calendar = CalendarApp.getCalendarById(yrdsbID());
    let start = new Date("01/01/2020"), end = new Date("01/01/2022");
    let events = calendar.getEvents(start, end);

    ss().getRange("A1").setValue("Holiday type");
    ss().getRange("B1").setValue("Start date (event is all day)");
    //javascript date class is weird, but stackoverflow is better
    ss().getRange("A2").setValue("Start Date: "+startDate.toLocaleDateString().substring(0, startDate.toLocaleString().indexOf(' ')));
    ss().getRange("B2").setValue("End Date: "+endDate.toLocaleDateString().substring(0, endDate.toLocaleString().indexOf(' ')));
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
/*
hi	01/03/2021	01/03/2021
hi	02/03/2021	03/03/2021
hi	03/03/2021	04/03/2021
hi	04/03/2021	05/03/2021
hi	05/03/2021	06/03/2021
hi	08/03/2021	09/03/2021
hi	09/03/2021	10/03/2021
hi	10/03/2021	11/03/2021
hi	11/03/2021	12/03/2021
hi	12/03/2021	13/03/2021
hi	15/03/2021	16/03/2021
hi	16/03/2021	17/03/2021
hi	17/03/2021	18/03/2021
hi	18/03/2021	19/03/2021
hi	19/03/2021	20/03/2021
hi	22/03/2021	23/03/2021
hi	23/03/2021	24/03/2021
hi	24/03/2021	25/03/2021
hi	25/03/2021	26/03/2021
hi	26/03/2021	27/03/2021
hi	29/03/2021	30/03/2021
hi	30/03/2021	31/03/2021
hi	31/03/2021	01/04/2021
*/