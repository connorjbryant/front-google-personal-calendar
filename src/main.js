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
 * Parse an ICS file and save its events to Google Calendar.
 */
async function importIcsIntoGoogleCalendar(
  icsText,
  frontSource
) {
  const parsedData = ICAL.parse(icsText);
  const calendarComponent =
    new ICAL.Component(parsedData);

  const eventComponents =
    calendarComponent.getAllSubcomponents("vevent");

  const importedEvents = [];

  for (const eventComponent of eventComponents) {
    const icalEvent = new ICAL.Event(eventComponent);

    const result = await saveIcsEventToGoogle(
      icalEvent,
      frontSource
    );

    importedEvents.push(result);
  }

  return importedEvents;
}

/**
 * Convert one ICAL event into a Google Calendar event.
 */
async function saveIcsEventToGoogle(
  icalEvent,
  frontSource
) {
  const iCalUid = icalEvent.uid || "";

  if (iCalUid) {
    const existingEvent =
      await findExistingGoogleEvent(iCalUid);

    if (existingEvent) {
      console.log(
        "Google Calendar already contains this event:",
        existingEvent
      );

      return {
        summary:
          existingEvent.summary ||
          icalEvent.summary ||
          "Calendar event",
        googleEventId: existingEvent.id,
        status: "already-existed"
      };
    }
  }

  const googleEvent = {
    summary:
      icalEvent.summary || "Calendar event",

    description:
      icalEvent.description || "",

    location:
      icalEvent.location || "",

    start: convertIcalDateToGoogle(
      icalEvent.startDate
    ),

    end: convertIcalDateToGoogle(
      icalEvent.endDate || icalEvent.startDate
    ),

    extendedProperties: {
      private: {
        frontConversationId: String(
          frontSource.conversationId || ""
        ),
        frontMessageId: String(
          frontSource.messageId || ""
        ),
        frontAttachmentId: String(
          frontSource.attachmentId || ""
        ),
        originalIcalUid: String(iCalUid)
      }
    }
  };

  console.log(
    "Sending event to Google Calendar:",
    googleEvent
  );

  const response =
    await gapi.client.calendar.events.insert({
      calendarId: "primary",
      resource: googleEvent
    });

  console.log(
    "Google Calendar event created:",
    response.result
  );

  return {
    summary:
      response.result.summary ||
      googleEvent.summary,
    googleEventId: response.result.id,
    htmlLink: response.result.htmlLink || null,
    status: "created"
  };
}

/**
 * Convert an ICAL.js date into Google's event format.
 */
function convertIcalDateToGoogle(icalDate) {
  if (!icalDate) {
    throw new Error(
      "The calendar event is missing a date."
    );
  }

  if (icalDate.isDate) {
    return {
      date: formatIcalDate(icalDate)
    };
  }

  return {
    dateTime:
      icalDate.toJSDate().toISOString()
  };
}

/**
 * Format an all-day ICAL date as YYYY-MM-DD.
 */
function formatIcalDate(icalDate) {
  const year = String(icalDate.year);

  const month = String(
    icalDate.month
  ).padStart(2, "0");

  const day = String(
    icalDate.day
  ).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

/**
 * Look for an event already saved with the same ICS UID.
 */
async function findExistingGoogleEvent(iCalUid) {
  const response =
    await gapi.client.calendar.events.list({
      calendarId: "primary",
      iCalUID: iCalUid,
      maxResults: 1,
      singleEvents: true
    });

  const matches =
    response.result.items || [];

  return matches[0] || null;
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
 * Connect to Google Calendar.
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
 * Listen for Front conversation changes.
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
        `Front context: ${
          context.type || "unknown"
        }. Open one conversation to continue.`
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
 * Find and download the first ICS attachment
 * in the selected Front conversation.
 */
async function syncCalendarInvites(context) {
  if (context.type !== "singleConversation") {
    setStatus("Open one Front conversation to continue.");
    return;
  }

  setStatus(
    "Searching this conversation for an ICS invitation..."
  );

  try {
    const messages = await getAllFrontMessages(context);

    console.log("Front messages:", messages);

    let matchingMessage = null;
    let matchingAttachment = null;

    for (const message of messages) {
      console.log("Full Front message:", message);
      console.log("Message keys:", Object.keys(message));
      console.log("Message content:", message.content);

      const attachments = [
        ...(message.content?.attachments || []),
        ...(message.attachments || [])
      ];

      console.log("Message attachments:", attachments);

      for (const attachment of attachments) {
        if (isCalendarAttachment(attachment)) {
          matchingMessage = message;
          matchingAttachment = attachment;
          break;
        }
      }

      if (matchingAttachment) {
        break;
      }
    }

    if (!matchingMessage || !matchingAttachment) {
      setStatus(
        "No ICS invitation was found in this conversation."
      );

      showDebug({
        stage: "Attachment search",
        conversationId:
          context.conversation?.id || null,
        messagesChecked: messages.length,
        calendarAttachmentFound: false
      });

      return;
    }

    const filename =
      matchingAttachment.name ||
      matchingAttachment.filename ||
      "calendar invitation";

    setStatus(
      `Found ${filename}. Downloading from Front...`
    );

    const file = await context.downloadAttachment(
      matchingMessage.id,
      matchingAttachment.id
    );

    if (!file) {
      throw new Error(
        "Front found the ICS attachment but could not download it."
      );
    }

    const icsText = await file.text();

    console.log("Downloaded ICS file:", file);
    console.log("ICS contents:", icsText);

    setStatus(
      `Downloaded ${filename}. Reading the calendar invitation...`
    );

    const importedEvents = await importIcsIntoGoogleCalendar(
      icsText,
      {
        conversationId: context.conversation?.id || "",
        messageId: matchingMessage.id,
        attachmentId: matchingAttachment.id
      }
    );

    if (importedEvents.length === 0) {
      setStatus(
        "The ICS file was downloaded, but it did not contain a calendar event."
      );

      showDebug({
        stage: "ICS parsing",
        filename,
        eventsFound: 0
      });

      return;
    }

    setStatus(
      importedEvents.length === 1
        ? `"${importedEvents[0].summary}" was saved to Google Calendar.`
        : `${importedEvents.length} events were saved to Google Calendar.`
    );

    showDebug({
      stage: "Google Calendar import complete",
      conversationId: context.conversation?.id || null,
      messageId: matchingMessage.id,
      attachmentId: matchingAttachment.id,
      filename,
      importedEvents
    });
  } catch (error) {
    setStatus(
      "Could not retrieve the ICS invitation from Front."
    );

    showDebug({
      stage: "Front attachment retrieval",
      error: error?.message || String(error)
    });

    console.error("ICS attachment error:", error);
  }
}

/**
 * Retrieve every message page from the selected conversation.
 */
async function getAllFrontMessages(context) {
  const messages = [];
  let pageToken;

  do {
    const page = await context.listMessages(pageToken);

    messages.push(...(page.results || []));

    pageToken =
      page.nextPageToken || undefined;
  } while (pageToken);

  return messages;
}

/**
 * Determine whether a Front attachment is an ICS calendar file.
 */
function isCalendarAttachment(attachment) {
  const filename = String(
    attachment.name ||
    attachment.filename ||
    ""
  ).toLowerCase();

  const contentType = String(
    attachment.contentType ||
    attachment.content_type ||
    attachment.type ||
    ""
  ).toLowerCase();

  return (
    filename.endsWith(".ics") ||
    contentType === "text/calendar" ||
    contentType.includes("calendar")
  );
}

initializeGoogle();