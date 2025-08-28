// key limitations: 4998 events on a year's sheet, 25 eventconfigs
// column configuration is hardcoded. So changes to the columns will break the script
// point person is always made at the same time for now, regardless of what event it's for

var spreadsheet = SpreadsheetApp.getActiveSheet();
var year = Number(spreadsheet.getName().substring(0,4));

const eventColumns = {
  name: 1,
  date: 2,
  main: 24,
  backup: 25
};  

const eventInfoColumns = {
  title: 0, 
  startTime: 1,
  endTime: 2,
  allDay: 3
}

const dataRange = "A3:AG500";
const eventInfoRange = "A2:Z3";
const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('eventconfig');
const calendarID = configSheet.getRange('G1').getValue();

/* Creates a a tab on the menu called "add to Google Calendar". Each month has a separate button
to add SC events. */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Add to Google Calendar')
    .addItem('Add August Events', 'createAugustEvents')
    .addItem('Add September Events', 'createSeptemberEvents')
    .addItem('Add October Events', 'createOctoberEvents')
    .addItem('Add November Events', 'createNovemberEvents')
    .addItem('Add December Events', 'createDecemberEvents')
    .addItem('Add January Events', 'createJanuaryEvents')
    .addItem('Add February Events', 'createFebruaryEvents')
    .addItem('Add March Events', 'createMarchEvents')
    .addItem('Add April Events', 'createAprilEvents')
    .addItem('Add May Events', 'createMayEvents')
    .addItem('Add June Events', 'createJuneEvents')
    .addItem('Add July Events', 'createJulyEvents')
    .addToUi();
}

function createAugustEvents() {createEvents(8, year);}
function createSeptemberEvents() {createEvents(9, year);}
function createOctoberEvents() {createEvents(10, year);}
function createNovemberEvents() {createEvents(11, year);}
function createDecemberEvents() {createEvents(12, year);}
function createJanuaryEvents() {createEvents(1, year+1);}
function createFebruaryEvents() {createEvents(2, year+1);}
function createMarchEvents() {createEvents(3, year+1);}
function createAprilEvents() {createEvents(4, year+1);}
function createMayEvents() {createEvents(5, year+1);}
function createJuneEvents() {createEvents(6, year+1);}
function createJulyEvents() {createEvents(7, year+1);}

/* Gets event title, start time, end time, and if it's all day from eventConfig sheet. 
   Returns an array with objects containing all event info. */
function getEventInfo() {
  const eventInfo = configSheet.getRange(eventInfoRange).getValues();
  let eventsArray = new Array();
  for (let x = 0; x < eventInfo.length; x++) {
    var eventInfoTitle = eventInfo[x][eventInfoColumns.title];
    var eventInfoStart = eventInfo[x][eventInfoColumns.startTime];
    var eventInfoEnd = eventInfo[x][eventInfoColumns.endTime];
    var eventInfoAllDay = eventInfo[x][eventInfoColumns.allDay];

    let eventInfoObject = {
      "title": eventInfoTitle,
      "startTime": new Date(eventInfoStart),
      "endTime": new Date(eventInfoEnd),
      "allDay": eventInfoAllDay
    }
    if (eventInfoTitle.length > 0) {
      eventsArray.push(eventInfoObject);
    }
  }
  return eventsArray;
}

var eventsArray = getEventInfo();

// Converts column letter to index (e.g., AE => 31)
function colLetterToIndex(col) {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index *= 26;
    index += col.charCodeAt(i) - 64;
  }
  return index;
}

// --- Point Person Events ---
const pointColLetter = configSheet.getRange('J2').getValue();
const pointStartTime = configSheet.getRange('K2').getValue(); // returns Date object with time
const pointEndTime = configSheet.getRange('L2').getValue();   // same
const pointColIndex = colLetterToIndex(pointColLetter) - 1; // 0-based

/* Creates google calendar events for a specified month and year.
Looks through data of specified data range for "SC" events. Gets names from "Audio"
and "Visual" column. Currently, start time = 8:30am and end time = 11:30am.
Checks if the event has already been created (prevents duplicates from being added) and 
deletes any existing events at the same time with different information.
*/
function createEvents(month, year) {
  const events = spreadsheet.getRange(dataRange).getValues();
  const calendar = CalendarApp.getCalendarById(calendarID);

  const startCriteria = new Date(year, month-1, 1);
  const endCriteria = new Date(year, month, 0);
  endCriteria.setHours(23);

  for (let x = 0; x < events.length; x++) {
    var eventName = events[x][eventColumns.name];
    var matchingEvent = eventsArray.find(event => event.title === eventName);
  
    if (matchingEvent) {
      var mainPerson = events[x][eventColumns.main];
      var backupPerson = events[x][eventColumns.backup];   
      var eventTitle = eventName + ": " + mainPerson + ", " + backupPerson;
      var eventDateStart = new Date(events[x][eventColumns.date]);
      var eventDateEnd = new Date(events[x][eventColumns.date]);
      

      // check if date falls within the specified month
      if (eventDateStart >= startCriteria && eventDateStart <= endCriteria) {
        // checks for existing events with the same date/time.
        var existingEvents = calendar.getEvents(eventDateStart, eventDateEnd);
        var eventExists = false;
        var pointExists = false;
        const pointPerson = events[x][pointColIndex];
        const pointTitle = eventName + " Point: " + pointPerson;

        for(let j = 0; j < existingEvents.length; j++) {
          if (existingEvents[j].getTitle() == eventTitle) {
            eventExists = true;
          }
        }

        if (!eventExists) {
          if (matchingEvent["allDay"]) {
            calendar.createAllDayEvent(eventTitle, eventDateStart);
          }

          else {
            eventDateStart.setHours(matchingEvent["startTime"].getHours());
            eventDateStart.setMinutes(matchingEvent["startTime"].getMinutes());
            eventDateEnd.setHours(matchingEvent["endTime"].getHours());
            eventDateEnd.setMinutes(matchingEvent["endTime"].getMinutes());
            calendar.createEvent(eventTitle, eventDateStart, eventDateEnd);
            Logger.log('Event Created: ' + eventTitle + eventDateStart.toString());
          }
          // calendar.createEvent(eventTitle, eventDateStart, eventDateEnd);
          // Logger.log('Event Created: ' + eventTitle + eventDateStart.toString());
        }
        else{Logger.log('Event already exists: ' + eventTitle + eventDateStart.toString());}

        if (pointPerson) {
          const pointStart = new Date(events[x][eventColumns.date]);
          const pointEnd = new Date(pointStart);
          pointStart.setHours(8, 0);
          pointEnd.setHours(8, 30);

          const pointEvents = calendar.getEvents(pointStart, pointEnd);
          let pointExists = false;
          for (let j = 0; j < pointEvents.length; j++) {
            if (pointEvents[j].getTitle() === pointTitle) {
              pointExists = true;
            } else {
              pointEvents[j].deleteEvent();
              Logger.log('Point Event Deleted: ' + pointEvents[j].getTitle());
            }
          }

          if (!pointExists) {
            calendar.createEvent(pointTitle, pointStart, pointEnd);
            Logger.log('Point Event Created: ' + pointTitle);
          } else {
            Logger.log('Point Event Already Exists: ' + pointTitle);
          }
        }
      }
    }
  }
}


// clears all events. for testing purposes
function clearEvents() {
  // const spreadsheet = SpreadsheetApp.getActiveSheet();
  const calendar = CalendarApp.getCalendarById(calendarID);
  const events = spreadsheet.getRange(dataRange).getValues();

  const earliestDay = events[0][eventColumns.date];
  const latestDay = new Date(year + 1, 7, 31);

  const createdEvents = calendar.getEvents(
    earliestDay,
    latestDay
  );

  Logger.log(`Deleting ${createdEvents.length} events from ${earliestDay} to ${latestDay}.`);

  for(const createdEvent of createdEvents) {
    createdEvent.deleteEvent();
  }
}



