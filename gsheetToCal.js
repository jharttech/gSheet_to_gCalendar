/* 
20241028 - This script is designed to be ran against a gSheet
The script then creates an event in the assigned Google Calendar for each row
*/

// Create main function for the script
function importDataToCalendar() {
  // Assign the active sheet to the sheet variable
  var sheet = SpreadsheetApp.getActiveSheet();
  // Assign the desired calander ID to the calendarId variable
  var es_calendarId = ''; // Google Calendar ID goes here
  var ms_calendarId = ''; // Google Calendar ID goes here
  var hs_calendarId = ''; // Google Calendar ID goes here
  var omtc_calendarId = ''; // Google Calendar ID goes here
  // Assign the data range in the active sheet to the eventsSheet variable
  var eventsSheet = sheet.getRange('A2:O'); // Adjust to your data range (include more rows)
  // Read the values in the row and assign to an array called events
  var events = eventsSheet.getValues();
  
  // Call the calendar by its ID and assign it to the calendar variable
  //var calendar = CalendarApp.getCalendarById(calendarId);
  
  // Loop through the data range
  for (var i = 0; i < events.length; i++) {
    // Assign the leave type to the calendar event
    var leaveType = events[i][7];
    var listed_buildings = events[i][3].split(",");

    var assigned_sub = events[i][14];
    
    if (assigned_sub != ""){
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

        // Create title based on conditions
        var title;
        // If the event column 8 has the word "no" in it then append "No Sub Needed!" to the event title
        if (events[i][8].toLowerCase().includes("no")) {
          title = events[i][2].concat(" - No Sub Needed!");
        } else {
          // If a sub is needed then append the sub name and leave type to the event title
          title = events[i][2].concat(" - ", events[i][14], ", Type: ", leaveType);
        }
        console.log("title:", title);
        
        // Grab the 12th column and assign its value to the decription variable
        var description = events[i][11];
        console.log("description:", description);
        
        // Set event start date
        var startDate = new Date(events[i][4]);
        
        // Set start time based on description of full day leave, A.M. Half day leave, P.M. Half day leave, or no sub needed
        if (description.toLowerCase().includes("am")) {
          startDate.setHours(7, 40, 0, 0);
        } else if (description.toLowerCase().includes("pm")) {
          startDate.setHours(11, 30, 0, 0);
        } else if (description.toLowerCase().includes("full") || description.toLowerCase().includes("no")) {
          startDate.setHours(7, 40, 0, 0);
        }
        console.log("Start Date:", startDate);
        
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
          }
        }
        console.log("End Date:", endDate);

        // Check for valid dates, if the event already exists, and if there is an error in creating the event
        if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime()) && assigned_sub != "") {
          // Check for existing events
          var existingEvents = calendar.getEvents(startDate, endDate, { search: title });
          if (existingEvents.length === 0) {
            // Create the event if it doesn't exist
            try {
              calendar.createEvent(title, startDate, endDate, { description: description });
              Logger.log('Event created: ' + title);
            } catch (e) {
              Logger.log('Error creating event: ' + e.message);
            }
          } else {
            Logger.log('Event already exists: ' + title + ' at ' + startDate);
          }
        } else {
          if (assigned_sub == "") {
            Logger.log('No sub listed but one is needed!!!');
          } else {
          Logger.log('Invalid date for event: ' + title);
          }
        }
      }
    }
  }
}

