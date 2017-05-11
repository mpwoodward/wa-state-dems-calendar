function pushToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 1, lastRow, 20);  // row, column, num rows, num columns
  var values = range.getValues();
  var updateRange = sheet.getRange('T1');
  
  var calendar = CalendarApp.getCalendarById('');
  
  updateRange.setBackground('red');
  
  var numValues = 0;
  
  for (var i = 0; i < values.length; i++) {
    /*
    Google forms takes care of requiring the necessary data, but we better double-check that we have
    the mandatory info just in case someone edited the spreadsheet and deleted vital stuff, specifically:
        * Hosted By [8] -- used as the event description
        * Event Name [9]
        * Event Location Name [10]
        * Event Location Address [11]
        * Event Location City [12]
        * Event Location State [13]
        * Event Location Zip [14]
        * Event Date [15]
        * Event Start Time [16]
        * Event End Time [17] -- seems silly but can't create a Google Calendar event without it, and
                                 many calendar standards require an end time
    */
    if (values[i][8].length > 0 && values[i][9].length > 0 && values[i][10].length > 0 && 
        values[i][11].length > 0 && values[i][12].length > 0 && values[i][13].length > 0 && 
        (!isNaN(parseFloat(values[i][14])) || values[i][14].length > 0) && 
        isDate(values[i][15]) && isDate(values[i][16])) {
      // see if it's already in the calendar by checking for a value in the calendar id column
      if (values[i][19].length == 0) {
        // set up the event data
        var eventTitle = values[i][9];
        var dtEventStart = joinDateAndTime(values[i][15], values[i][16]);
        var dtEventEnd = joinDateAndTime(values[i][15], values[i][17]);        
        var location = values[i][10] + ', ' + values[i][11] + ', ' + values[i][12] + ', ' + values[i][13] + ' ' + values[i][14];
        var description = 'Hosted by ' + values[i][8];
        var options = {'location': location, 'description': description};
        
        // create the event
        var event = calendar.createEvent(eventTitle, dtEventStart, dtEventEnd, options);
        
        // add calendar event id to spreadsheet
        var eventID = event.getId();
        sheet.getRange(i + 2, 20).setValue(eventID);
      }
    }
  }
  
  updateRange.setBackground('white');
}

function isDate(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    if (isNaN(v.getTime())) {
      return false;
    } else {
      return true;
    }
  } else {
    return true;
  }
}

function joinDateAndTime(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

// add the push to calendar menu item when the spreadsheet is opened
function onOpen() {
  var menu = [{name: 'Push to Calendar', functionName: 'pushToCalendar'}];
  SpreadsheetApp.getActive().addMenu('Calendar', menu);
}
