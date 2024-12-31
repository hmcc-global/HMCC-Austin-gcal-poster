var currentDate = new Date();
var currentYear = currentDate.getFullYear();

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

const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('eventconfig')
const calendarID = configSheet.getRange('G1').getValue();
const dataRange = "AL1:ZZ3";

const eventRows = {
  date: 0,
  what: 1,
  prayer: 2
};

// Functions to create events for each month 
function createAugustEvents() {createEvents(8, 2024);}
function createSeptemberEvents() {createEvents(9, 2024);}
function createOctoberEvents() {createEvents(10, 2024);}
function createNovemberEvents() {createEvents(11, 2024);}
function createDecemberEvents() {createEvents(12, 2024);}
function createJanuaryEvents() {createEvents(1, 2025);}
function createFebruaryEvents() {createEvents(2, 2025);}
function createMarchEvents() {createEvents(3, 2025);}
function createAprilEvents() {createEvents(4, 2025);}
function createMayEvents() {createEvents(5, 2025);}
function createJuneEvents() {createEvents(6, 2025);}
function createJulyEvents() {createEvents(7, 2025);}

/* Creates google calendar events for a specified month and year.
Looks through data of specified data range for "SC" events. 
Currently, start time = 8:00am and end time = 8:30am.
Checks if the event has already been created (prevents duplicates from being added) and 
deletes any existing events at the same time with different information.
*/
function createEvents(month, year) {
  const calendar = CalendarApp.getCalendarById(calendarID);
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const events = spreadsheet.getRange(dataRange);
  const eventValues = events.getValues();

  const startCriteria = new Date(year, month - 1, 1);
  const endCriteria = new Date(year, month, 0);
  endCriteria.setHours(23);

  // go through the entire data range and create events
  for (let x = 0; x < events.getWidth(); x+=2) {
    var eventTitle = eventValues[eventRows.what][x];
    if (eventTitle == "SC") {
      var eventDateStart = new Date(eventValues[eventRows.date][x]);
      eventDateStart.setHours(8);
      var eventDateEnd = new Date(eventValues[eventRows.date][x]);
      eventDateEnd.setHours(8);
      eventDateEnd.setMinutes(30);
      var prayerPerson = eventValues[eventRows.prayer][x];
      var prayerPersonTitle = "Prayer/Ann: " + prayerPerson;

      // check if date falls within the specified month
      if (eventDateStart >= startCriteria && eventDateStart <= endCriteria) {

        // check if the event has already been created
        var existingEvents = calendar.getEvents(eventDateStart, eventDateEnd);
        var eventExists = false;
        for(let j = 0; j < existingEvents.length; j++) {
          if (existingEvents[j].getTitle() == prayerPersonTitle) {
            eventExists = true;
          }
          else {
            existingEvents[j].deleteEvent()
            Logger.log('Event Deleted: ' + existingEvents[j].getTitle() + existingEvents[j].getStartTime().toString());
          }
        }
        if (!eventExists) {
          calendar.createEvent(prayerPersonTitle, eventDateStart, eventDateEnd);
          Logger.log('Event Created: ' + prayerPersonTitle + eventDateStart.toString());
        }
        else{Logger.log('Event already exists: ' + prayerPersonTitle + eventDateStart.toString());}
      }
    }
  }
}

// clears all events. for testing purposes
function clearEvents() {
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const calendar = CalendarApp.getCalendarById(calendarID);
  const events = spreadsheet.getRange(dataRange).getValues();

  const earliestDay = events[eventRows.date][0];
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



