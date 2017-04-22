// Team Asha weekday mailer - to send reminders etc.
var VERSION = 2017.1;
var ADDRESS = "fill-in-your-group-address-here"; // group@googlegroups.com
var HILLS_RUN_DAY = 7; // Change to a valid day to trigger hills text
var FORM_ID = form-id-here; // Looks like 1NJLUKtKjdQhiLq2MX-YpbUNTY3nWm_UMlCz2Ml7YZF1";
var LOG_ONLY = false; // Logs, but not send the email, for testing.
var GenericSchedule = "Weekday Run Schedule"; // Sometimes, the sheet names are not the race names.
var raceNames = [GenericSchedule];
var SfHalf = "SF Half 2017"; // Generic race!


function findBeginRowCol(values, rows, cols) {
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols - 1; c++) {
            // Perhaps the following comparision should be a regexp.
            if (values[r][c] == "Monday" && values[r][c + 1] == "Sunday") {
                return {
                  row : r,
                  col : c,
                  endCol : c + 1
                };
            }
        }
    }
    return null;
}

function findMondayCol(values, rows, cols, row, col) {
    for (c = col; c < cols; c++) {
        if (values[row][c] == "Monday") {
            return c;
        }
    }
    return null;
}

function getDailyMileage(runningDate, values, rows, cols) {
    // Find the row that has the "Monday" and "Sunday"
    var beginRowCol = findBeginRowCol(values, rows, 4);
    if (beginRowCol == null) {
        throw Utilities.formatString("Unable to find a row with \"Monday\" and" +
                                     "\"Sunday\"");
    }

    // These are relative to the range.
    var row = beginRowCol.row;
    var beginCol = beginRowCol.col;
    var endCol = beginRowCol.endCol;

    var mondayCol = findMondayCol(values, rows, cols, row, endCol);
    if (mondayCol == null) {
        var str = Utilities.formatString("Can't find Monday column");
        throw str;
    }

    // Now look for the row that includes the given date.
    for (var r = row + 1; r < rows; r++) {
        var weekBegin = values[r][beginCol];
        var weekEnd = values[r][endCol];
        if (!((weekBegin instanceof Date) && (weekEnd instanceof Date))) {
            throw Utilities.formatString("%s and %s are not dates", weekBegin, weekEnd);
        }
        if (runningDate > weekBegin && runningDate < weekEnd) {
            var mileage = values[r][mondayCol + runningDate.getDay() - 1];
            return mileage.replace(/^\s+/, "").replace(/\s+$/, ""); 
        }
    }
    throw Utilities.formatString("Can't find row for %s", runningDate);
}

function getNextForWeekdayRun(fromDate) {
    // Move to noon to avoid day light troubles.
    var fromNoon = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate(),
                            12, 0, 0, 0);

    // If we are already on the required day, return
    var fromDay = fromDate.getDay();
    var hour = fromDate.getHours();

    // 0 is Sunday, we need either Monday(1) or Thursday(4)
    // Allow for the same day early morning sent emails :)
    if ((fromDay == 1 || fromDay == 4) && hour < 6)
        return fromDate;

    // Increment the date, till we get the correct day
    var nextDate = fromDate;
    var oneDay = 24 * 60 * 60 * 1000; // In milliseconds
    for (var i = 0; i < 5; i++) {
        var nextTime = nextDate.getTime() + oneDay;
        var nextDate = new Date(nextTime);
        if (nextDate.getDay() == 1 || nextDate.getDay() == 4)
            return nextDate;
    }
    Logger.log("Can't find next day for " + fromDate);
    throw "Can't find next day for " + fromDate;
    return fromDate;
}


function sendEmail(d, time, mileage, doHills, extraText) {
  var month = d.getMonth() + 1;
  var date = d.getDate();
  var year = d.getYear();
  
  // Get the email address of the active user - that's you.
  var myEmail = Session.getActiveUser().getEmail();
  var formatTime = month + "/" + date + "/" + year + " " + time;
  
  var subjectLoc;
  var bodyLoc;
  var hillText;
  var postRun;
  var extraDirections = "";
  if (d.getDay() == HILLS_RUN_DAY) { // Thursdays are hill repeats
    subjectLoc = "Team Asha, Weekday run @ Vasona Park";
    bodyLoc = "Vasona Park on Los Gatos Creek Trail here : http://goo.gl/maps/rVfRI"
    hillText = " (with some hills)"
    postRun = " + Post run stretches";
    extraDirections = "Park your car on the road, follow (by foot) along Garden Hill Dr" +
                      " into the park, let us meet near the gate. There is no auto entry" +
                      " from Garden Hill Dr. into the park. We will do a few hill repeats" +
                      " by looping around.";
  } else {
    subjectLoc = "Team Asha, Weekday run @ Los Gatos Creek Trail - ";
    bodyLoc = "Los Gatos Creek Tail - near Campbell downtown, here : http://goo.gl/maps/dsB1r";
    hillText = "";
    postRun = " + Post run strength exercises / stretches";
  }
  
  var plan = "Warm up + Pre run stretch + Required Mileage" + hillText + postRun + ".";
  var whatToBring = "Water bottle, your usual exercise wear (dress in layers) with your running shoes, visor / sun glasses.";
  var goodToGet = "Watch with 2 interval timers for run/walk; a banana or a bar to eat after the session.";
  var footer = "Please call Coach XXXX, if you need help getting there or getting delayed.";
 
  var milesToRun = "<p><b>Mileage</b>: ";
  var i = 0;
  for (var m in mileage) {
    if (mileage.hasOwnProperty(m)) {
      if (i != 0) {
        milesToRun += ",  ";
      }
      milesToRun += "<i>" + m + "</i> -- " + mileage[m];
      i++;
    }
  }
  
  var subject = subjectLoc + " " + formatTime;
  var htmlBody = "<b>Where</b>: " + bodyLoc;
  htmlBody += "<p><b>When</b>: " + formatTime;
  htmlBody += "<p><b>Plan</b>: " + plan;
  htmlBody += milesToRun;
  if (milesToRun.indexOf("(") != -1) {
    htmlBody += "<p><small>Numbers in () are for Season 1 runners training" +
                " for a second Half Marathon.</small>";
  }

  htmlBody += "<p><b>What to wear/bring</b>: " + whatToBring;
  htmlBody += "<p><b>Good to get</b>: " + goodToGet;
  htmlBody += "<p>" + footer;
  htmlBody += "<p>" + extraText + "</p>";
  htmlBody += "<p>" + extraDirections + "</p>";
  htmlBody += "<p><small>" + "(script version:" + VERSION + ")" + "</small>";
  var body = "This email needs HTML capable email client " + htmlBody;
  
  if (LOG_ONLY) {
    Logger.log("Sending email : subject " + subject + "\n htmlBody: " +
               htmlBody + "\n");
  } else { 
    MailApp.sendEmail(ADDRESS,
                   subject, body, {htmlBody : htmlBody, cc : myEmail});
  }
}

function getStartTime(runningDate) {
   var month = runningDate.getMonth() + 1;
   var date = runningDate.getDate();
   var year = runningDate.getYear();
   return " 06:30 AM"; 
}

function sendWeekDayRunEmail(runningDate, mileage) {    
    var doHills = false;
    if (runningDate.getDay() == HILLS_RUN_DAY) {
       doHills = true;
    }
    var startTime = getStartTime(runningDate);
    Logger.log("Sending email : start " +
                runningDate + " " + startTime + " doing Hills:" + doHills);
    sendEmail(runningDate, startTime, mileage, doHills, "");
}


function getMileage(runningDate, activeSpreadSheet) {
    var mileage = {};
    for(var i = 0, len = raceNames.length; i < len; i++) {
      raceName = raceNames[i];
      var dailySheet = activeSpreadSheet.getSheetByName(raceName);
      if (dailySheet) {
        var rows = dailySheet.getMaxRows();
        var cols = dailySheet.getMaxColumns();
        var values = dailySheet.getSheetValues(1, 1, rows, cols);
        var m = getDailyMileage(runningDate, values, rows, cols);
        Logger.log("Mileage for " + raceName + " is " + m);
        if (raceName == GenericSchedule) {
          raceName = SfHalf
        }
        mileage[raceName] = m;
      } else {
        Logger.log("No sheet found for " + raceName);
      }
    }
    return mileage;
}

function main() {
    var nextRunningDate = getNextForWeekdayRun(new Date());
    activeSpreadSheet = SpreadsheetApp.openById(FORM_ID);
    if (activeSpreadSheet == null) {
       throw "Can't open spreadsheet id " + id;
    }
    mileage = getMileage(nextRunningDate, activeSpreadSheet);
    sendWeekDayRunEmail(nextRunningDate, mileage);
}


function test() {
    var d = new Date();
    var v = d.getTime();
    activeSpreadSheet = SpreadsheetApp.openById(FORM_ID);
    if (activeSpreadSheet == null) {
       throw "Can't open spreadsheet id " + id;
    }
    for (var i = 0; i < 30; i++) {
        var next = new Date(v + 2 * 24 * 3600 * 1000 * i);
        var nextRunningDate = getNextForWeekdayRun(next);
        mileage = getMileage(nextRunningDate, activeSpreadSheet);
        sendWeekDayRunEmail(nextRunningDate, mileage);
    }
}
