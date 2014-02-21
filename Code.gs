function sendToCal() {
  //this is the ID of the calendar to add the event to, found on the calendar settings page of the calendar in question
  var calendarId = "6lom1o0nb0k401pb7fqbaopsgg@group.calendar.google.com";
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var lr = rows.getLastRow();
  var lc = rows.getLastColumn();

  function getValAtColNamed(title){
    return getValAtCol(getColNum(title))
  }

  function getValAtCol(colNum){
    return sheet.getRange(lr,colNum,1,1).getValue()
  }

  function getColNum(title){
    var re = new RegExp(title, "i");
    for (var i = 1; i <= lc; i++) {
      if(sheet.getRange(1,i,1,1).getValue().match(re)) return i;
    };
    return false;
  }

  var startDate = getValAtColNamed('Pickup Time');
  var endDate = addMinutes(startDate, 60);

  function addMinutes(date, minutes) {
    return new Date(date.getTime() + minutes*60000);
  }

  // pull relevant info from the new sheet entry
  var passengerName = getValAtColNamed('Passenger Name');
  var pickupLocation = getValAtColNamed('Pickup Location');
  var pickupInstructions = getValAtColNamed('Pickup Instructions');
  var dropoffAddress = getValAtColNamed('Dropoff Address');
  var otherNotes = getValAtColNamed('Other Notes');
  var timeStamp = getValAtColNamed('Timestamp');
  var submittedBy = getValAtColNamed('Your Name');

  var submissionInfo = "Submitted on: "+ timeStamp + ", by " + submittedBy;
  var description = "Instructions:\n  " + pickupInstructions +"\n\n" +
                    "Pickup location:\n  "+pickupLocation+"\n\n"+
                    "Dropoff Address(es):\n  " + dropoffAddress +"\n\n" +
                    "Other Notes:\n  " + otherNotes +"\n\n" +
                    submissionInfo;

  function getLatestAndSubmitToCalendar() {
    var title = passengerName
    createEvent(calendarId,title,startDate,endDate,description);
  }​


  function createEvent(calendarId,title,startDate,endDate,description) {
    var cal = CalendarApp.getCalendarById(calendarId);
    var start = startDate;
    var end = endDate;

    var event = cal.createEvent(title, start, end, {
        description : description,
        location : pickupLocation
    });
  };

  if ( getValAtColNamed('New Passenger?') === 'No' || getValAtColNamed('Schedule Ride now?') === 'Yes' ){
    getLatestAndSubmitToCalendar();
    MailApp.sendEmail('tickets@lifthero.uservoice.com', 'New Ride Scheduled!',
                    "Passenger: " + passengerName + "\n\n Description:\n" +
                    description);
  }
}
