
// var ss = function() {return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()}

function onInstall(e) {
    onOpen(e);
}
function onOpen(e) {
    // let menu = UI.createMenu("Forms"); Used as standalone
    let menu = SpreadsheetApp.getUi().createAddonMenu(); //used with congunction with google marketplace
    menu.addItem("Delete all events", "deleteAll").addToUi(); //delete events from calendar
    menu.addItem("Fill in lessons", "fillInDate").addToUi(); //add dates beside events in Spreadsheet
    menu.addItem("Add events to calendar", "addToCalendar").addToUi(); //add spreadsheet events into calendar
}
function deleteAll() { //need to figure out how to delete only inserted ones (so I don't end up deleting the wrong things)
    // in the form of MM/DD/YYYY
    let start = new Date(), end = new Date();
    let calendarId = ''; // need to figure out how to get this

    let calendar = CalendarApp.getCalendarById(calendarId);
    let events = calendar.getEvents(start, end);
    while(events.length>0) {
        events[i].deleteEvent();
    }
}
function fillInDate() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = ss.getRange().getValues();

    for(let i=0;i<data.length;i++) {
        //check if date is holiday
        //else add proper date beside event
    }
}
function addToCalendar() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = ss.getRange().getValues(); //need to indicate range when I create the template
    let calendaryId = '';

    let calendar = CalendarApp.getCalendarById(calendaryId);
    for(let i=0;i<data.length;i++) {
        //check if start and end dates are filled
        //check to see if event apperas in calendar?
        var newEvent = defaultCalendar.createEvent(title, newDate(start), newDate(end), {location:something});
    }
}