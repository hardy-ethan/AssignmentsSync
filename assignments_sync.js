const { google } = require('googleapis');
const { isEqual } = require('lodash');
const crypto = require('crypto');
const moment = require('moment');

const { CALENDAR_ID, SPREADSHEET_ID } = require('./config.json')

const SCOPES = [
 'https://www.googleapis.com/auth/calendar',
 'https://www.googleapis.com/auth/spreadsheets'
];
const RANGE = 'WN25!A2:I';

async function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function retryWithBackoff(operation, maxRetries = 5, baseDelay = 1000) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await operation();
    } catch (error) {
      if (error.response?.status === 429 && attempt < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, attempt);
        const jitter = Math.random() * 1000;
        console.log(`Rate limited. Attempt ${attempt + 1}/${maxRetries}. Retrying in ${delay}ms`);
        await wait(delay + jitter);
        continue;
      }
      throw error;
    }
  }
}

async function authorize() {
  const auth = new google.auth.GoogleAuth({
    keyFile: 'credentials.json',
    scopes: SCOPES,
  });
  return auth.getClient();
}

async function getSpreadsheetData(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const response = await retryWithBackoff(() => 
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    })
  );

  const data = [];

  for (const [index, row] of Object.entries(response.data.values)) {
    const uuid = row[8] || crypto.randomUUID();

    // If UUID was generated, update spreadsheet
    if (!row[8]) {
      await retryWithBackoff(() => 
        sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `WN25!I${Number(index)+2}`, // We query starting at A2
          valueInputOption: 'RAW',
          requestBody: {
            values: [[uuid]]
          }
        })
      );
    }

    data.push({
      Origin: row[0] ?? "Unknown",
      Name: row[1] ?? "Unknown", 
      'Due Date': row[2] ?? "Unknown",
      'Due Time': row[3] ?? "Unknown",
      Status: row[4] ?? "Unknown",
      Difficulty: row[5] ?? "Unknown",
      Priority: row[6] ?? "Unknown",
      Notes: row[7] ?? "Unknown",
      UUID: uuid
    });
  }

  return data;
}

function getEventData(assignment) {
  const originalDateTimeString = `${assignment['Due Date']}|${assignment['Due Time']}`;

  const dueDateAndTime = moment(originalDateTimeString, 'L|LTS');

  if (!dueDateAndTime.isValid()) {
    throw new Error(`Moment could not parse time "${originalDateTimeString}"`)
  }
  
  const eventDateTime = dueDateAndTime.toISOString(true);

  return {
    summary: `${assignment.Origin}: ${assignment.Name}`,
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

    const existingEvents = await retryWithBackoff(() => 
      calendar.events.list({
        calendarId: CALENDAR_ID,
        timeMin: new Date().toISOString(),
      })
    );

    const existingEventMap = new Map(
      existingEvents.data.items.map(event => [event.extendedProperties?.private?.uuid, event])
    );

    for (const assignment of assignments) {
      const eventData = getEventData(assignment);
      const existingEvent = existingEventMap.get(assignment.UUID);

      if (existingEvent) {
        if (!eventsAreEqual(existingEvent, eventData)) {
          await retryWithBackoff(() => 
            calendar.events.update({
              calendarId: CALENDAR_ID,
              eventId: existingEvent.id,
              requestBody: eventData,
            })
          );
          console.log('Updated event:', assignment.Name);
        }
        existingEventMap.delete(assignment.UUID);
      } else {
        await retryWithBackoff(() => 
          calendar.events.insert({
            calendarId: CALENDAR_ID,
            requestBody: eventData,
          })
        );
        console.log('Created event:', assignment.Name);
      }
    }

    for (const [_, event] of existingEventMap) {
      await retryWithBackoff(() => 
        calendar.events.delete({
          calendarId: CALENDAR_ID,
          eventId: event.id,
        })
      );
      console.log('Deleted event:', event.summary);
    }
  } catch (error) {
    console.error('Error:', error.response?.data || error);
  }
}

syncWithCalendar();