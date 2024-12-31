var currentDate = new Date();
var currentYear = currentDate.getFullYear();

const eventColumns = {
  what: 0,
  date: 1,
  main: 9,
  backup: 10
};  

const dataRange = "A2:K500";
const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('eventconfig')
const calendarID = configSheet.getRange('G1').getValue();

/* Creates a a tab on the menu called "add to Google Calendar". Each month has a separate button
to add SC events. */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Add to Google Calendar')
    .addItem('Add August SC Events', 'createAugustEvents')
    .addItem('Add September SC Events', 'createSeptemberEvents')
    .addItem('Add October SC Events', 'createOctoberEvents')
    .addItem('Add November SC Events', 'createNovemberEvents')
    .addItem('Add December SC Events', 'createDecemberEvents')
    .addItem('Add January SC Events', 'createJanuaryEvents')
    .addItem('Add February SC Events', 'createFebruaryEvents')
    .addItem('Add March SC Events', 'createMarchEvents')
    .addItem('Add April SC Events', 'createAprilEvents')
    .addItem('Add May SC Events', 'createMayEvents')
    .addItem('Add June SC Events', 'createJuneEvents')
    .addItem('Add July SC Events', 'createJulyEvents')
    .addToUi();
}

function createAugustEvents() {createEvents(8, currentYear);}
function createSeptemberEvents() {createEvents(9, currentYear);}
function createOctoberEvents() {createEvents(10, currentYear);}
function createNovemberEvents() {createEvents(11, currentYear);}
function createDecemberEvents() {createEvents(12, currentYear);}
function createJanuaryEvents() {createEvents(1, currentYear+1);}
function createFebruaryEvents() {createEvents(2, currentYear+1);}
function createMarchEvents() {createEvents(3, currentYear+1);}
function createAprilEvents() {createEvents(4, currentYear+1);}
function createMayEvents() {createEvents(5, currentYear+1);}
function createJuneEvents() {createEvents(6, currentYear+1);}
function createJulyEvents() {createEvents(7, currentYear+1);}

/* Creates google calendar events for a specified month and year.
Looks through data of specified data range for "SC" events. Gets names from "Audio"
and "Visual" column. Currently, start time = 8:30am and end time = 11:30am.
Checks if the event has already been created (prevents duplicates from being added) and 
deletes any existing events at the same time with different information.
*/
function createEvents(month, year) {
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const events = spreadsheet.getRange(dataRange).getValues();
  const calendar = CalendarApp.getCalendarById(calendarID);

  const startCriteria = new Date(year, month-1, 1);
  const endCriteria = new Date(year, month, 0);
  endCriteria.setHours(23);


  for (let x = 0; x < events.length; x++) {
    var eventName = events[x][eventColumns.what];
    if (eventName == "SC") {
      var eventDateStart = new Date(events[x][eventColumns.date]);
      eventDateStart.setHours(8);
      eventDateStart.setMinutes(30);
      var eventDateEnd = new Date(events[x][eventColumns.date]);
      eventDateEnd.setHours(11);
      eventDateEnd.setMinutes(30);
      var mainPerson = events[x][eventColumns.main];
      var backupPerson = events[x][eventColumns.backup];
      var eventTitle = "SC: " + mainPerson + ", " + backupPerson;

        // check if date falls within the specified month
      if (eventDateStart >= startCriteria && eventDateStart <= endCriteria) {

        // checks for existing events with the same date/time.
        var existingEvents = calendar.getEvents(eventDateStart, eventDateEnd);
        var eventExists = false;
        for(let j = 0; j < existingEvents.length; j++) {
          if (existingEvents[j].getTitle() == eventTitle) {
            eventExists = true;
          }
          else {
            existingEvents[j].deleteEvent();
            Logger.log('Event Deleted: ' + existingEvents[j].getTitle() + existingEvents[j].getStartTime().toString());
          }
        }

        if (!eventExists) {
          calendar.createEvent(eventTitle, eventDateStart, eventDateEnd);
          Logger.log('Event Created: ' + eventTitle + eventDateStart.toString());
        }
        else{Logger.log('Event already exists: ' + eventTitle + eventDateStart.toString());}
      }
    }
  }
}

// clears all events. for testing purposes
function clearEvents() {
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const calendar = CalendarApp.getCalendarById(calendarID);
  const events = spreadsheet.getRange(dataRange).getValues();

  const earliestDay = events[0][eventColumns.date];
  const latestDay = new Date(2025, 7, 31);

  const createdEvents = calendar.getEvents(
    earliestDay,
    latestDay
  );

  Logger.log(`Deleting ${createdEvents.length} events from ${earliestDay} to ${latestDay}.`);

  for(const createdEvent of createdEvents) {
    createdEvent.deleteEvent();
  }
}


