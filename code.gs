/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Update calendar', functionName: 'updateCalendar_'}];
  SpreadsheetApp.getActive().addMenu('Calendar', menu);
}

function addEventsToCalendar_(values, range) {
  var cal = CalendarApp.getCalendarsByName("McCormick Shared Calendar Fall '19 [TEST]")[0];
  for (var i = 1; i < values.length; i++) {   
    var event = values[i];
    if (event[11] != "x" && event[9] == "Yes") {
      var title = event[3];
      var start = joinDateAndTime_(event[5], event[6]);
      var end = joinDateAndTime_(event[5], event[7]);
      var options = {location: event[8], description: event[4]};
      var newEvent = cal.createEvent(title, start, end, options);
      sendEmail_(event[2], event[3], true, event[10]);
      event[11] = 'x';
    }
    if (event[11]!="x" && event[9] == "No") {
      sendEmail_(event[2], event[3], false, event[10]);
      event[11] = 'x';
    }
  }
  range.setValues(values);
}

function sendEmail_(emailAddress, title, added, comments) {
  var subject = ""
  if (added) {
    subject = "[Mccorm-cal] ADDED: " + title;
  }
  else {
    subject = "[Mccorm-cal] NOT ADDED: " + title;
  }
  var message = "Hi,\n\nThank you for your event submission. Your event " + title + " was ";
  if (added) {
    message = message + "added.";
  }
  else {
    message = message + "not added.";
  }
  message = message + " Please see below for any further comments.\n\n" + comments + "\n\n";
  message = message + "If you have any questions, please email mccormick-secretary@mit.edu.\n\nSincerely,\nMcCormick Secretary";
  Logger.log(emailAddress, subject, message);
  MailApp.sendEmail(emailAddress, subject, message);
}

function updateCalendar_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();
  addEventsToCalendar_(values, range);
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}
