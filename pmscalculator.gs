function getAllEventsSummary(fullSummary=true) {

  var now = new Date();
  var day = now.getDate();
  var month = now.getMonth();
  var year = now.getFullYear()
  var firstDayOfCurrentMonth = new Date(year, month);
  var millisecToHours = 60 * 60 * 1000;

  var allDayMeeting = "All day meeting";

  var outOfOfficeText = "This is an out-of-office event";
  // TODO: Need to remove TGIF (HydePark,..) as well

  var basicComment = "";

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Current sheet
  var activeCell = ss.getCurrentCell(); // The current highlighted cell

  row = activeCell.getRow();
  column = activeCell.getColumn();

  var totalTime = 0; //Total calendar (meeting time)

  // All events from own calendar
  var events = CalendarApp.getDefaultCalendar().getEvents(firstDayOfCurrentMonth, now);

  var excludeTime = false;
  var duration = 0;

  /**
    Create array all events to insert into the sheet
    **/
  var allEventsForSheet = events.map(event => {
    excludeTime = false;
    duration = (event.getEndTime() - event.getStartTime()) / millisecToHours;



    if (event.getDescription().indexOf(outOfOfficeText) !== -1) {
      comment = outOfOfficeText;
      excludeTime = true;
    } else {
      comment = basicComment;
    }

    if (duration >= 24) { //All day meeting to exclude
      excludeTime = true;
      comment = allDayMeeting;
    }

    totalTime += excludeTime ? 0 : duration;


    return [event.getTitle(), duration, comment]

  });



  var infos = [
    ['Calculation day', `${year}-${month + 1}-${day}`],
    ['Month', now.toLocaleString('default', { month: 'long' })],
    ['Total time (Hours)', totalTime]
  ]


  /**
  Insert the info into the spreadsheet
  **/
  var insertInfo = ss.getRange(row, column, infos.length, infos[0].length).setValues(infos);

  if(fullSummary){

  /**
  Insert the events list into the spreadsheet
  **/
  var results = ss.getRange(row + infos.length + 1, column, allEventsForSheet.length, allEventsForSheet[0].length).setValues(allEventsForSheet)

 }


}

function getOnlyTime(){
  getAllEventsSummary(false)
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Meeting Time Calculator')
    .addItem('Get only Total Time', 'getOnlyTime')
    .addItem('Get Total Time with Events Details', 'getAllEventsSummary')
    .addToUi()
}
