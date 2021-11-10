 'use strict';

 const functions = require('firebase-functions');
 const {google} = require('googleapis');
 const {WebhookClient} = require('dialogflow-fulfillment');
 
 // Enter your calendar ID below and service account JSON below
 const calendarId = "<Add your calendar ID here>"
 const serviceAccount = {<Add your service account details here>}; // Starts with {"type": "service_account",... // Remove the complete block and paste the code from the service account Key here
 
 // Set up Google Calendar Service account credentials
 const serviceAccountAuth = new google.auth.JWT({
   email: serviceAccount.client_email,
   key: serviceAccount.private_key,
   scopes: 'https://www.googleapis.com/auth/calendar'
 });
 
 const calendar = google.calendar('v3');
 process.env.DEBUG = 'dialogflow:*'; // enables lib debugging statements
 
 const timeZone = 'Asia/Kolkata'; // you can mention any timezone like Africa/Asmara OR America/Buenos_Aire
 const timeZoneOffset = '+5:30'; // timeoffset will be with respect to the timezone you choose to make
 
 exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
   const agent = new WebhookClient({ request, response });
   const appointment_type = agent.parameters.AppointmentType
   function makeAppointment (agent) {
     // Calculate appointment start and end datetimes (end = +1hr from start)
    const AppointmentDate = agent.parameters.date.split('T')[0];
    const AppointmentTime = agent.parameters.time.split('T')[1].substr(0,8);
    const dateTimeStart = new Date(AppointmentDate + 'T' + AppointmentTime+timeZoneOffset);
    const dateTimeEnd = new Date(new Date(dateTimeStart).setHours(dateTimeStart.getHours() + 1));
    
     // Check the availibility of the time, and make an appointment if there is time on the calendar
    return createCalendarEvent(dateTimeStart, dateTimeEnd, appointment_type).then(() => {
       agent.add(`Let me see if we can fit you in on ${AppointmentDate} at ${AppointmentTime}! Yes It is fine!.`);
     }).catch(() => {
       agent.add(`I'm sorry, there are no slots available for ${AppointmentDate} at ${AppointmentTime}.`);
     });
   }
 
   let intentMap = new Map();
   intentMap.set('<Name of the intent which you are making the appointment>', makeAppointment); //Replace the <Text> with the intent name
   agent.handleRequest(intentMap);
 });

function createCalendarEvent (dateTimeStart, dateTimeEnd, appointment_type) {
   return new Promise((resolve, reject) => {
     calendar.events.list({
       auth: serviceAccountAuth, // List events for time period
       calendarId: calendarId,
       timeMin: dateTimeStart.toISOString(),
       timeMax: dateTimeEnd.toISOString()
     }, (err, calendarResponse) => {
       // Check if there is a event already on the Calendar
       if (err || calendarResponse.data.items.length > 0) {
         reject(err || new Error('Requested time conflicts with another appointment'));
       } else {
         // Create event for the requested time period
         calendar.events.insert({ auth: serviceAccountAuth,
           calendarId: calendarId,
           resource: {summary: appointment_type +' Appointment', description: appointment_type,
             start: {dateTime: dateTimeStart},
             end: {dateTime: dateTimeEnd}}
         }, (err, event) => {
           err ? reject(err) : resolve(event);
         }
         );
       }
     });
   });
 }
