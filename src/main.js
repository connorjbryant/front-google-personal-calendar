import $ from "jquery";
import Front from "@frontapp/plugin-sdk";
import ICAL from "ical.js";

const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;

const DISCOVERY = "https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest";

const SCOPES = "https://www.googleapis.com/auth/calendar.events";

let tokenClient;
let googleConnected = false;
let currentFrontContext = null;

gapi.load("client", async function(){
  await gapi.client.init({
    discoveryDocs: [DISCOVERY]
  });

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: function() {}
  });
});

$("#login").on("click", function(){
  tokenClient.callback = async function(resp){
    console.log("Google auth response: ", resp);

    if(resp.error){
      console.log("Google auth failed: ", resp);
      return;
    }

    googleConnected = true;
    console.log("Google Calendar connected");

    // If a Front convo is already open when Google becomes connected immediately inspect convo
    if (currentFrontContext?.type === "singleConversation"){
      await syncCalendarInvites(currentFrontContext);
    }

    const events = [];
    let pageToken = null;

    try {
      do {
        const response = await gapi.client.calendar.events.list({
          calendarId: "primary",
          singleEvents: true,
          orderBy: "startTime",
          timeMin: new Date().toISOString(),
          pageToken: pageToken
        });

        events.push(...(response.result.items || []));
        pageToken = response.result.nextPageToken || null;
      } while (pageToken);

      console.log("Upcoming events:");
      console.log(events);
    } catch (error) {
      console.log("Event loading error: ", error);
    }
  };

  tokenClient.requestAccessToken();
});

console.log("About to subscribe to Front context");
console.log("Front SDK ", Front);
Front.contextUpdates.subscribe(function(context){
  console.log('Front context changed: ', context);

  currentFrontContext = context;

  if (context.type !== "singleConversation"){
    console.log("No single convo selected");
    return;
  }

  console.log("Selected conversation: ", context.conversation);

  if (!googleConnected){
    console.log("Google Calendar not connected");
    return;
  }

  await syncCalendarInvites(context);
});

async function syncCalendarInvites(context){
  console.log("Ready to sync this convo for an ICS file");
  console.log(context.conversation);
}