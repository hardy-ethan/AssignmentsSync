const { google } = require('googleapis');
const { isEqual } = require('lodash');

const { CALENDAR_ID, SPREADSHEET_ID } = require('./config.json')

const SCOPES = [
 'https://www.googleapis.com/auth/calendar',
 'https://www.googleapis.com/auth/spreadsheets.readonly'
];
const RANGE = 'WN25!A2:I';

async function authorize() {
 const auth = new google.auth.GoogleAuth({
   keyFile: 'credentials.json',
   scopes: SCOPES,
 });
 return auth.getClient();
}

async function getSpreadsheetData(auth) {
 const sheets = google.sheets({ version: 'v4', auth });
 const response = await sheets.spreadsheets.values.get({
   spreadsheetId: SPREADSHEET_ID,
   range: RANGE,
 });

 return response.data.values.map(row => ({
   Origin: row[0],
   Name: row[1], 
   'Due Date': row[2],
   'Due Time': row[3],
   Status: row[4],
   Difficulty: row[5],
   Priority: row[6],
   Notes: row[7],
   UUID: row[8]
 }));
}

function getEventData(assignment) {
 const eventDateTime = `${assignment['Due Date']}T${assignment['Due Time']}:00-04:00`;

 return {
   summary: assignment.Name,
   description: `Difficulty: ${assignment.Difficulty}\nPriority: ${assignment.Priority}\nNotes: ${assignment.Notes}`,
   start: { 
     dateTime: eventDateTime,
     timeZone: 'America/New_York'
   },
   end: { 
     dateTime: eventDateTime,
     timeZone: 'America/New_York'
   },
   extendedProperties: {
     private: { uuid: assignment.UUID }
   }
 };
}

function eventsAreEqual(event1, event2) {
 return isEqual({
   summary: event1.summary,
   description: event1.description,
   start: event1.start,
   end: event1.end
 }, {
   summary: event2.summary,
   description: event2.description,
   start: event2.start,
   end: event2.end
 });
}

async function syncWithCalendar() {
 try {
   const auth = await authorize();
   const assignments = await getSpreadsheetData(auth);
   const calendar = google.calendar({ version: 'v3', auth });

   const existingEvents = await calendar.events.list({
     calendarId: CALENDAR_ID,
     timeMin: new Date().toISOString(),
   });

   const existingEventMap = new Map(
     existingEvents.data.items.map(event => [event.extendedProperties?.private?.uuid, event])
   );

   for (const assignment of assignments) {
     const eventData = getEventData(assignment);
     const existingEvent = existingEventMap.get(assignment.UUID);

     if (existingEvent) {
       if (!eventsAreEqual(existingEvent, eventData)) {
         await calendar.events.update({
           calendarId: CALENDAR_ID,
           eventId: existingEvent.id,
           requestBody: eventData,
         });
         console.log('Updated event:', assignment.Name);
       }
       existingEventMap.delete(assignment.UUID);
     } else {
       await calendar.events.insert({
         calendarId: CALENDAR_ID,
         requestBody: eventData,
       });
       console.log('Created event:', assignment.Name);
     }
   }

   for (const [_, event] of existingEventMap) {
     await calendar.events.delete({
       calendarId: CALENDAR_ID,
       eventId: event.id,
     });
     console.log('Deleted event:', event.summary);
   }
 } catch (error) {
   console.error('Error:', JSON.stringify(error.response?.data ?? {}) || error);
 }
}

syncWithCalendar();