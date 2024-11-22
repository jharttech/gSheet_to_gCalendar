/* 
20241119 - This script is designed to be ran against a gSheet
The script then creates an event in the assigned Google Calendar for each row
*/

// Create main function for the script
function importDataToCalendar() {
  
  // Assign the active sheet to the sheet variable
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Assign the desired calander ID to the calendarId variable
  var es_calendarId = 'FIXME'; // Google Calendar ID goes here
  var ms_calendarId = 'FIXME';
  var hs_calendarId = 'FIXME';
  var omtc_calendarId = 'FIXME';

  // Assign the data range in the active sheet to the eventsSheet variable
  var eventsSheet = sheet.getRange('A2:O'); // Adjust to your data range (include more rows)
  // Read the values in the row and assign to an array called events
  var events = eventsSheet.getValues();
  const date = new Date();
  var day = date.getDate();
  var prevDay = date.getDate() - 1;
  var month = date.getMonth() + 1;
  var year = date.getFullYear();
  let todayDate = new Date(`${month}/${day}/${year}`);
  let yesterdayDate = new Date(`${month}/${prevDay}/${year}`);
  Logger.log('INFO: Todays Date: ' + todayDate)
  
  // Loop through the data range
  for (var i = 0; i < events.length; i++) {
    Logger.log("HERE")
    
    // Function to set inital caledar info
    function setEventInfo() {
      // Assign the leave type to the calendar event
      var leaveType = events[i][7];
      // Get a list of buildings the event has a relation to 
      var listedBuildings = events[i][3].split(",");
      // Get what sub has been assigned.
      var assigned_sub = events[i][14]

      // Create title based on conditions
      var title;
      // If the event column 8 has the word "no" in it then append "No Sub Needed!" to the event title
      if (!events[i][2].toLowerCase().includes("extra") && events[i][8].toLowerCase().includes("no")) {
        title = events[i][2].concat(" - No Sub Needed!");
      } else {
        // If a sub is needed then append the sub name and leave type to the event title
        title = events[i][2].concat(" - ", events[i][14], ", Type: ", leaveType);
      }
      console.log("INFO: title:", title);
      
      // Grab the 12th column and assign its value to the decription variable
      var description = events[i][11];
      console.log("INFO : description:", description);
      
      // Set event start date
      var startDate = new Date(events[i][4]);
      
      // Set start time based on description of full day leave, A.M. Half day leave, P.M. Half day leave, or no sub needed
      if (description.toLowerCase().includes("am")) {
        startDate.setHours(7, 40, 0, 0);
      } else if (description.toLowerCase().includes("pm")) {
        startDate.setHours(11, 30, 0, 0);
      } else if (description.toLowerCase().includes("full") || description.toLowerCase().includes("no")) {
        startDate.setHours(7, 40, 0, 0);
      } else {
        startDate.setHours(7, 40, 0, 0);
      }
      console.log("INFO: Start Date:", startDate);
      
      // Determine end date time based on description of full day leave, A.M. Half day leave, P.M. Half day leave, or no sub needed
      var endDate;
      if (events[i][5] !== "") {
        endDate = new Date(events[i][5]);
        endDate.setHours(15, 30, 0, 0);
      } else {
        endDate = new Date(startDate);
        if (description.toLowerCase().includes("am")) {
          endDate.setHours(11, 30, 0, 0);
        } else if (description.toLowerCase().includes("pm") || description.toLowerCase().includes("full") || description.toLowerCase().includes("no")) {
          endDate.setHours(15, 30, 0, 0);
        } else {
          endDate.setHours(15, 30, 0, 0);
        }
      }
      console.log("INFO: End Date:", endDate);
      return{assigned_sub, description, title, startDate, endDate};
    }

    // Delete function set to run if column 13 contains the work delete
    function deleteEvent(calendar, title, startDate, endDate) {
      Logger.log(calendar + ", " + title + ", " + startDate + ", " + endDate)
      // Get the event ID so we can remove the correct event
      var eventId = calendar.getEvents(startDate, endDate, {search: title});
      Logger.log("INFO: eventId: " + eventId)
      for (var b in eventId){
        var id = eventId[b].getId();
        var tempTitle = eventId[b].getTitle()
        Logger.log("INFO: id is: " + id)
        if (id !== undefined) {
          console.log(id)
          eventToDel = calendar.getEventById(id);
          eventToDel.deleteEvent()
          Logger.log('Event: ' + tempTitle + ' has been deleted')
        }
      }
    }

    // Create function to create new events
    function create(calendar, title, startDate, endDate, description) {
      Logger.log("create function")
      // Check for valid dates, if the event already exists, and if there is an error in creating the event
      if ((!events[i][13].toLowerCase().includes("delete")) && (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime()) && assigned_sub != "")) {
        // Check for existing events
        var existingEvents = calendar.getEvents(startDate, endDate, { search: title });
        if (existingEvents.length === 0) {
          // Create the event if it doesn't exist
          Logger.log("Does not exist")
          try {
            calendar.createEvent(title, startDate, endDate, { description: description });
            Logger.log('Event created: ' + title);
          } catch (e) {
            Logger.log('1 Error creating event: ' + e.message);
          }
        // Check if there is a delete value
        //} else if (assigned_sub == "") {
            //Logger.log('No sub listed but one is needed!!!')
        } else {
          Logger.log('Event already exists: ' + title + ' at ' + startDate);
        }
      }
    }

    // Get the buildings that the teacher teaches in
    var listedBuildings = events[i][3].split(",");
    for (var x = 0; x < listedBuildings.length; x++){
      var building = listedBuildings[x]
      Logger.log('listed: ' + listedBuildings);

      if (building.toLowerCase().includes("high") || listedBuildings[x].toLowerCase().includes("technology")) {
        var calendar = CalendarApp.getCalendarById(hs_calendarId);
      } else if (building.toLowerCase().includes("middle")) {
        var calendar = CalendarApp.getCalendarById(ms_calendarId);
      } else if (building.toLowerCase().includes("omtc")) {
        var calendar = CalendarApp.getCalendarById(omtc_calendarId);
      } else if (building.toLowerCase().includes("elementary") || building.toLowerCase().includes("pre")) {
        var calendar = CalendarApp.getCalendarById(es_calendarId);
      }

      // Run the setEventInfo() function to set needed variables
      var {leaveType, assigned_sub, description, title, startDate, endDate} = setEventInfo();

      if (events[i][13].toLowerCase().includes("delete")) {
        Logger.log("Title: " + title)
        //Logger.log("INFO: eventId: " + eventId)
        deleteEvent(calendar, title, startDate, endDate);
      } else if (assigned_sub != "" && (events[i][4] >= yesterdayDate || events[i][5] >= yesterdayDate)){
        try {
          create(calendar, title, startDate, endDate, description)
        } catch (e) {
          Logger.log('2 Error creating event: ' + e.message);
        }
      } else if (!events[i][8].toLowerCase().includes("no") && assigned_sub == "") {
        Logger.log('No sub listed but one is needed!!!')
      } else {
        Logger.log('Event date passed, no need to create the event!')
      }
    
    }
  }
}

// The code below is only for removing duplicate events that were caused by manual entry into the spreadsheet - This many no longer be needed
function deleteMulitEvent() {
  // JHART ONLY BELOW!!!!
  // Assign the active sheet to the sheet variable
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Assign the desired calander ID to the calendarId variable
  var es_calendarId = 'FIXME'; // Google Calendar ID goes here
  var ms_calendarId = 'FIXME';
  var hs_calendarId = 'FIXME';
  var omtc_calendarId = 'FIXME';
  
  var eventsSheet = sheet.getRange('A0:O0'); // Adjust to your data range (include more rows)
  // Read the values in the row and assign to an array called events
  var events = eventsSheet.getValues();
  
  for (var i = 0; i < events.length; i++) {
    // Assign the leave type to the calendar event
    var leaveType = events[i][7];
    var listed_buildings = events[i][3].split(",");
    // Assign the data range in the active sheet to the eventsSheet variable
    const date = new Date();
    var day = date.getDate();
    var prevDay = date.getDate() - 1;
    var month = date.getMonth() + 1;
    var year = date.getFullYear();
    let todayDate = new Date(`${month}/${day}/${year}`);
    let yesterdayDate = new Date(`${month}/${prevDay}/${year}`);
    var title;
      // If the event column 8 has the word "no" in it then append "No Sub Needed!" to the event title
      if (events[i][8].toLowerCase().includes("no")) {
        title = events[i][2].concat(" - No Sub Needed!");
      } else {
        // If a sub is needed then append the sub name and leave type to the event title
        title = events[i][2].concat(" - ", events[i][14], ", Type: ", leaveType);
      }

    Logger.log('INFO: Todays Date: ' + todayDate)
    startDate = new Date(events[i][4]);
    endDate = new Date(events[i][4]);
    startDate.setHours(0,0,0,0)
    endDate.setHours(0,0,0,0)
    for (var x = 0; x < listed_buildings.length; x++) {
          console.log(listed_buildings[x]);
          var building = listed_buildings[x];
          // Assign calander id based on building
          if (building.toLowerCase().includes("high") || building.toLowerCase().includes("technology")) {
            var calendar = CalendarApp.getCalendarById(hs_calendarId);
          } else if (building.toLowerCase().includes("middle")) {
            var calendar = CalendarApp.getCalendarById(ms_calendarId);
          } else if (building.toLowerCase().includes("omtc")) {
            var calendar = CalendarApp.getCalendarById(omtc_calendarId);
          } else if (building.toLowerCase().includes("elementary") || building.toLowerCase().includes("pre")) {
            var calendar = CalendarApp.getCalendarById(es_calendarId);
          }
      Logger.log(startDate + ', ' + endDate + ', ' + title)
      var eventId = calendar.getEventsForDay(startDate);
      Logger.log(eventId)
      for (var b in (eventId)){
        var id = eventId[b].getId();
        Logger.log(id)
        var tempTitle = eventId[b].getTitle()
        Logger.log(tempTitle)
        if (id !== undefined) {
          console.log(id)
          eventToDel = calendar.getEventById(id);
          eventToDel.deleteEvent()
          Logger.log('Event: ' + tempTitle + ' has been deleted')
        }
      }
    }
  }
}    


