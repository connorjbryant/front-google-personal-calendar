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
  if (context.type !== "singleConversation") {
    setStatus("Open one Front conversation to continue.");
    return;
  }

  setStatus("Searching this conversation for an ICS invitation...");

  try {
    const messages = await getAllFrontMessages();

    console.log("Front messages:", messages);

    const calendarAttachments = [];

    messages.forEach(function (message) {
      const attachments = message.attachments || [];

      attachments.forEach(function (attachment) {
        console.log("Front attachment:", attachment);

        if (isCalendarAttachment(attachment)) {
          calendarAttachments.push({
            message,
            attachment
          });
        }
      });
    });

    if (calendarAttachments.length === 0) {
      setStatus("No ICS invitation was found in this conversation.");

      showDebug({
        stage: "Attachment search",
        conversationId: context.conversation?.id || null,
        messagesChecked: messages.length,
        calendarAttachmentsFound: 0
      });

      return;
    }

    const match = calendarAttachments[0];

    const attachmentName =
      match.attachment.filename ||
      match.attachment.name ||
      "calendar invitation";

    setStatus(`Found ${attachmentName}. Downloading it from Front...`);

    const file = await Front.downloadAttachment(
      match.message.id,
      match.attachment.id
    );

    if (!file) {
      throw new Error(
        "Front found the attachment but did not return the file."
      );
    }

    console.log("Downloaded ICS file:", file);

    const icsText = await file.text();

    console.log("ICS file contents:");
    console.log(icsText);

    setStatus(
      `Successfully downloaded ${attachmentName}. Ready to import it into Google Calendar.`
    );

    showDebug({
      stage: "ICS downloaded",
      conversationId: context.conversation?.id || null,
      messageId: match.message.id,
      attachmentId: match.attachment.id,
      filename: file.name,
      fileType: file.type,
      fileSize: file.size,
      calendarAttachmentsFound: calendarAttachments.length
    });
  } catch (error) {
    setStatus("Could not retrieve the calendar invitation from Front.");

    showDebug({
      stage: "Front attachment retrieval",
      error: error?.message || String(error)
    });

    console.error("ICS attachment error:", error);
  }
}

async function getAllFrontMessages() {
  const messages = [];
  let pageToken = undefined;

  do {
    const page = await Front.listMessages(pageToken);

    messages.push(...(page.results || []));

    pageToken = page.nextPageToken || undefined;
  } while (pageToken);

  return messages;
}

function isCalendarAttachment(attachment) {
  const filename = String(
    attachment.filename ||
    attachment.name ||
    ""
  ).toLowerCase();

  const contentType = String(
    attachment.contentType ||
    attachment.content_type ||
    ""
  ).toLowerCase();

  return (
    filename.endsWith(".ics") ||
    contentType === "text/calendar" ||
    contentType.includes("calendar")
  );
}

initializeGoogle();