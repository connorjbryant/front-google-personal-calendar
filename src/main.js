import $ from "jquery";
import Front from "@frontapp/plugin-sdk";
import ICAL from "ical.js";

const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;

const DISCOVERY =
  "https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest";

const SCOPES =
  "https://www.googleapis.com/auth/calendar.events";

let tokenClient = null;
let googleConnected = false;
let currentFrontContext = null;

/**
 * Update the visible status message.
 */
function setStatus(message) {
  $("#status").text(message);
}

/**
 * Display debugging information inside the plugin.
 */
function showDebug(data) {
  $("#debug").text(JSON.stringify(data, null, 2));
}

/**
 * Initialize the Google Calendar API.
 */
function initializeGoogle() {
  setStatus("Initializing Google Calendar...");

  if (typeof gapi === "undefined") {
    setStatus("The Google API script did not load.");
    console.error("The Google API script did not load.");
    return;
  }

  if (typeof google === "undefined") {
    setStatus("Google Identity Services did not load.");
    console.error("Google Identity Services did not load.");
    return;
  }

  if (!CLIENT_ID) {
    setStatus("The Google OAuth client ID is missing.");
    console.error("VITE_GOOGLE_CLIENT_ID is missing.");
    return;
  }

  gapi.load("client", async function () {
    try {
      await gapi.client.init({
        discoveryDocs: [DISCOVERY]
      });

      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: function () {}
      });

      console.log("Google API initialized");

      $("#login").prop("disabled", false);

      setStatus("Waiting for Front context...");
    } catch (error) {
      setStatus("Google Calendar initialization failed.");
      showDebug({
        stage: "Google initialization",
        error: error?.message || String(error)
      });

      console.error("Google initialization failed:", error);
    }
  });
}

/**
 * Connect the user to Google Calendar.
 */
$("#login").prop("disabled", true);

$("#login").on("click", function () {
  if (!tokenClient) {
    setStatus("Google Calendar is not ready yet.");
    return;
  }

  setStatus("Opening Google sign-in...");

  tokenClient.callback = async function (response) {
    console.log("Google auth response:", response);

    if (response.error) {
      setStatus("Google authorization failed.");

      showDebug({
        stage: "Google authorization",
        error: response.error,
        errorDescription: response.error_description || null
      });

      console.error("Google auth failed:", response);
      return;
    }

    googleConnected = true;

    $("#login")
      .prop("disabled", true)
      .text("Google Calendar Connected");

    console.log("Google Calendar connected");

    if (currentFrontContext?.type === "singleConversation") {
      setStatus(
        "Google Calendar connected. Preparing the selected conversation..."
      );

      await syncCalendarInvites(currentFrontContext);
    } else {
      setStatus(
        "Google Calendar connected. Open a Front conversation containing an ICS invitation."
      );
    }

    await loadUpcomingGoogleEvents();
  };

  tokenClient.requestAccessToken({
    prompt: ""
  });
});

/**
 * Load all upcoming Google Calendar events.
 */
async function loadUpcomingGoogleEvents() {
  const events = [];
  let pageToken = null;

  try {
    do {
      const response =
        await gapi.client.calendar.events.list({
          calendarId: "primary",
          singleEvents: true,
          orderBy: "startTime",
          timeMin: new Date().toISOString(),
          pageToken
        });

      events.push(...(response.result.items || []));

      pageToken =
        response.result.nextPageToken || null;
    } while (pageToken);

    console.log("Upcoming Google Calendar events:", events);
  } catch (error) {
    console.error("Event loading error:", error);
  }
}

/**
 * Listen for conversation changes inside Front.
 */
console.log("About to subscribe to Front context");
console.log("Front SDK:", Front);

try {
  Front.contextUpdates.subscribe(function (context) {
    console.log("Front context received:", context);

    currentFrontContext = context;

    showDebug({
      contextId: context.id || null,
      type: context.type || null,
      conversationId:
        context.conversation?.id || null,
      conversationSubject:
        context.conversation?.subject || null,
      googleConnected
    });

    if (context.type !== "singleConversation") {
      setStatus(
        `Front context: ${context.type || "unknown"}. Open one conversation to continue.`
      );

      return;
    }

    if (!googleConnected) {
      setStatus(
        "Conversation detected. Connect Google Calendar to continue."
      );

      return;
    }

    setStatus(
      "Conversation detected and Google Calendar connected."
    );

    syncCalendarInvites(context).catch(function (error) {
      setStatus("Could not process the selected conversation.");

      showDebug({
        stage: "Front conversation sync",
        error: error?.message || String(error)
      });

      console.error(
        "Front conversation sync failed:",
        error
      );
    });
  });
} catch (error) {
  setStatus("Could not subscribe to Front context.");

  showDebug({
    stage: "Front context subscription",
    error: error?.message || String(error)
  });

  console.error(
    "Front context subscription failed:",
    error
  );
}

/**
 * Temporary sync function.
 *
 * Once Front context works, this will be expanded to:
 * 1. Retrieve the conversation messages.
 * 2. Find the ICS attachment.
 * 3. Download and parse it.
 * 4. Save the event to Google Calendar.
 */
async function syncCalendarInvites(context) {
  console.log(
    "Ready to search this conversation for an ICS file:",
    context.conversation
  );

  setStatus(
    "Front conversation detected. Ready to search for an ICS invitation."
  );

  showDebug({
    stage: "Ready to sync",
    contextType: context.type,
    conversationId:
      context.conversation?.id || null,
    conversationSubject:
      context.conversation?.subject || null,
    googleConnected
  });
}

initializeGoogle();