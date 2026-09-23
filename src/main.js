import $ from "jquery";
import Front from "@frontapp/plugin-sdk";
import ICAL from "ical.js";
import "./style.css";

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
        htmlLink: existingEvent.htmlLink || null,
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
  if (!iCalUid) {
    return null;
  }

  // First check Google's native iCalUID field. This catches events that
  // already exist in Google Calendar because they were imported/received
  // through another calendar workflow.
  const nativeResponse =
    await gapi.client.calendar.events.list({
      calendarId: "primary",
      iCalUID: iCalUid,
      maxResults: 1,
      singleEvents: true,
      showDeleted: false
    });

  const nativeMatches =
    nativeResponse.result.items || [];

  if (nativeMatches.length > 0) {
    return nativeMatches[0];
  }

  // Events created by this plugin use events.insert(). Google assigns those
  // events its own iCalUID, so we also search the private property where the
  // original invitation UID is stored. This prevents the same Front invite
  // from being inserted again when the conversation is reopened.
  const privateResponse =
    await gapi.client.calendar.events.list({
      calendarId: "primary",
      privateExtendedProperty:
        `originalIcalUid=${iCalUid}`,
      maxResults: 1,
      singleEvents: true,
      showDeleted: false
    });

  const privateMatches =
    privateResponse.result.items || [];

  return privateMatches[0] || null;
}


/**
 * Switch between the existing ICS importer and meeting creator.
 */
$(document).on("click", ".tab", function () {
  const panelId = $(this).data("panel");

  $(".tab").removeClass("is-active");
  $(this).addClass("is-active");

  $(".panel").removeClass("is-active").prop("hidden", true);
  $(`#${panelId}`).addClass("is-active").prop("hidden", false);
});

/**
 * Give the meeting form sensible defaults.
 */
function setMeetingDefaults() {
  const now = new Date();
  const start = new Date(now.getTime() + 60 * 60 * 1000);
  start.setMinutes(Math.ceil(start.getMinutes() / 15) * 15, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);

  const localDate = [
    start.getFullYear(),
    String(start.getMonth() + 1).padStart(2, "0"),
    String(start.getDate()).padStart(2, "0")
  ].join("-");

  const time = (date) =>
    `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;

  $("#meeting-date").val(localDate);
  $("#meeting-start").val(time(start));
  $("#meeting-end").val(time(end));
  $("#meeting-repeat").val("none");
  $("#meeting-repeat-end").val("never");
  $("#meeting-repeat-until").val("");
  $("#meeting-repeat-end-field").prop("hidden", true);
  $("#meeting-repeat-until-field").prop("hidden", true);
}

/**
 * Show recurrence controls only when a repeating schedule is selected.
 * A series can run forever or stop on a selected date.
 */
function updateRecurrenceEndFields() {
  const repeats = $("#meeting-repeat").val() !== "none";
  const endsOnDate = repeats && $("#meeting-repeat-end").val() === "date";

  $("#meeting-repeat-end-field").prop("hidden", !repeats);
  $("#meeting-repeat-until-field").prop("hidden", !endsOnDate);
  $("#meeting-repeat-until").prop("required", endsOnDate);

  if (!endsOnDate) {
    $("#meeting-repeat-until").val("");
  }
}

$("#meeting-repeat").on("change", updateRecurrenceEndFields);
$("#meeting-repeat-end").on("change", updateRecurrenceEndFields);

/**
 * Build the Google Calendar RRULE used for recurring meetings.
 */
function buildRecurrenceRule(repeat, repeatEnd, untilDate) {
  if (repeat === "none") {
    return null;
  }

  const rules = {
    weekly: "FREQ=WEEKLY;INTERVAL=1",
    biweekly: "FREQ=WEEKLY;INTERVAL=2",
    monthly: "FREQ=MONTHLY;INTERVAL=1"
  };

  const rule = rules[repeat];
  if (!rule) {
    return null;
  }

  // Omitting UNTIL/COUNT makes the Google Calendar recurrence continue indefinitely.
  if (repeatEnd === "never") {
    return `RRULE:${rule}`;
  }

  if (!untilDate) {
    return null;
  }

  // UNTIL is inclusive. Use the end of the selected day in UTC.
  const until = `${untilDate.replaceAll("-", "")}T235959Z`;
  return `RRULE:${rule};UNTIL=${until}`;
}

function recurrenceLabel(repeat, repeatEnd, untilDate) {
  const labels = {
    weekly: "Repeats weekly",
    biweekly: "Repeats every 2 weeks",
    monthly: "Repeats monthly"
  };

  if (!labels[repeat]) {
    return "";
  }

  if (repeatEnd === "never") {
    return `${labels[repeat]} indefinitely.`;
  }

  const formattedUntil = new Date(`${untilDate}T12:00:00`).toLocaleDateString(undefined, {
    year: "numeric",
    month: "long",
    day: "numeric"
  });

  return `${labels[repeat]} until ${formattedUntil}.`;
}

/**
 * Turn the guest field into a clean, unique list of email addresses.
 */
function parseGuestEmails(value) {
  return [...new Set(
    String(value || "")
      .split(/[\n,;]+/)
      .map((email) => email.trim().toLowerCase())
      .filter(Boolean)
  )];
}

/**
 * Pull recipients from the currently selected Front conversation.
 */
async function getFrontConversationRecipients() {
  if (currentFrontContext?.type !== "singleConversation") {
    throw new Error("Open one Front conversation first.");
  }

  const recipients = [];
  let pageToken;

  do {
    const page = await currentFrontContext.listRecipients(pageToken);
    recipients.push(...(page.results || []));
    pageToken = page.nextPageToken || undefined;
  } while (pageToken);

  return recipients;
}

$("#add-front-recipients").on("click", async function () {
  try {
    setStatus("Loading people from this Front conversation...");

    const recipients = await getFrontConversationRecipients();
    const foundEmails = recipients
      .map((recipient) => recipient.email || recipient.handle || recipient.address || "")
      .filter((value) => String(value).includes("@"));

    const currentEmails = parseGuestEmails($("#meeting-guests").val());
    const merged = [...new Set([...currentEmails, ...foundEmails])];

    if (!merged.length) {
      setStatus("No email recipients were available from this conversation.");
      return;
    }

    $("#meeting-guests").val(merged.join(", "));
    setStatus(`${foundEmails.length} conversation recipient(s) added to the guest list.`);
  } catch (error) {
    setStatus(error?.message || "Could not load conversation recipients.");
    console.error("Front recipient loading failed:", error);
  }
});

/**
 * Create a real event in the connected user's primary Google Calendar.
 * Guests receive normal Google Calendar invitations via sendUpdates=all.
 */
$("#meeting-form").on("submit", async function (event) {
  event.preventDefault();

  if (!googleConnected) {
    setStatus("Connect Google Calendar before creating a meeting.");
    return;
  }

  const title = $("#meeting-title").val().trim();
  const date = $("#meeting-date").val();
  const startTime = $("#meeting-start").val();
  const endTime = $("#meeting-end").val();
  const repeat = $("#meeting-repeat").val();
  const repeatEnd = $("#meeting-repeat-end").val();
  const repeatUntil = $("#meeting-repeat-until").val();
  const guestEmails = parseGuestEmails($("#meeting-guests").val());

  if (!title || !date || !startTime || !endTime) {
    setStatus("Add a title, date, start time, and end time.");
    return;
  }

  const start = new Date(`${date}T${startTime}:00`);
  const end = new Date(`${date}T${endTime}:00`);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) {
    setStatus("The meeting end time must be after the start time.");
    return;
  }

  if (repeat !== "none" && repeatEnd === "date" && !repeatUntil) {
    setStatus("Choose when the recurring meeting should stop.");
    return;
  }

  if (repeat !== "none" && repeatEnd === "date" && repeatUntil < date) {
    setStatus("The repeat-until date cannot be before the first meeting.");
    return;
  }

  const recurrenceRule = buildRecurrenceRule(repeat, repeatEnd, repeatUntil);

  const invalidEmails = guestEmails.filter(
    (email) => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
  );

  if (invalidEmails.length) {
    setStatus(`Check the guest email address${invalidEmails.length > 1 ? "es" : ""}: ${invalidEmails.join(", ")}`);
    return;
  }

  // Google requires an explicit IANA time zone for recurring events.
  // Use the browser's local time zone so recurrence also follows daylight-saving changes.
  const meetingTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";

  const resource = {
    summary: title,
    description: $("#meeting-description").val().trim(),
    location: $("#meeting-location").val().trim(),
    start: { dateTime: start.toISOString(), timeZone: meetingTimeZone },
    end: { dateTime: end.toISOString(), timeZone: meetingTimeZone },
    attendees: guestEmails.map((email) => ({ email })),
    extendedProperties: {
      private: {
        createdFromFrontPlugin: "true",
        frontConversationId: String(currentFrontContext?.conversation?.id || "")
      }
    }
  };

  if (recurrenceRule) {
    resource.recurrence = [recurrenceRule];
  }

  const request = {
    calendarId: "primary",
    sendUpdates: guestEmails.length ? "all" : "none",
    resource
  };

  if ($("#meeting-google-meet").is(":checked")) {
    request.conferenceDataVersion = 1;
    resource.conferenceData = {
      createRequest: {
        requestId: `front-${Date.now()}-${Math.random().toString(36).slice(2)}`,
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    };
  }

  try {
    $("#create-meeting").prop("disabled", true).text("Creating meeting...");
    setStatus("Creating the meeting in Google Calendar...");

    const response = await gapi.client.calendar.events.insert(request);
    const created = response.result;
    const meetLink = created.hangoutLink || "";

    const links = [];
    if (created.htmlLink) {
      links.push(`<a href="${created.htmlLink}" target="_blank" rel="noopener noreferrer">Open in Google Calendar</a>`);
    }
    if (meetLink) {
      links.push(`<a href="${meetLink}" target="_blank" rel="noopener noreferrer">Open Google Meet</a>`);
    }

    $("#meeting-result")
      .prop("hidden", false)
      .html(`
        <strong>Meeting created</strong>
        <span>${created.summary || title}</span>
        ${recurrenceRule ? `<span>${recurrenceLabel(repeat, repeatEnd, repeatUntil)}</span>` : ""}
        ${guestEmails.length ? `<span>Invitations sent to ${guestEmails.length} guest${guestEmails.length === 1 ? "" : "s"}.</span>` : ""}
        ${links.length ? `<div class="result-links">${links.join("")}</div>` : ""}
      `);

    setStatus(
      guestEmails.length
        ? "Meeting created in Google Calendar and invitations sent."
        : "Meeting created in Google Calendar."
    );

    showDebug({
      stage: "Google meeting created",
      googleEventId: created.id,
      htmlLink: created.htmlLink || null,
      hangoutLink: meetLink || null,
      attendees: guestEmails,
      recurrence: recurrenceRule || null,
      timeZone: meetingTimeZone,
      frontConversationId: currentFrontContext?.conversation?.id || null
    });
  } catch (error) {
    const message = error?.result?.error?.message || error?.message || String(error);
    setStatus(`Could not create the meeting: ${message}`);
    showDebug({ stage: "Google meeting creation", error: message });
    console.error("Google meeting creation failed:", error);
  } finally {
    $("#create-meeting").prop("disabled", false).text("Create & send invitations");
  }
});

setMeetingDefaults();

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
      .addClass("is-connected")
      .text("✓ Google Calendar Connected");

    $("#calendar-actions").prop("hidden", false);

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

    const createdEvents = importedEvents.filter(
      (event) => event.status === "created"
    );

    const existingEvents = importedEvents.filter(
      (event) => event.status === "already-existed"
    );

    // Show direct Google Calendar links for imported or already-existing events.
    const linkedEvents = importedEvents.filter((event) => event.htmlLink);

    if (linkedEvents.length) {
      $("#import-result")
        .prop("hidden", false)
        .html(
          linkedEvents.map((event) => `
            <div class="import-result-event">
              <strong>${$("<div>").text(event.summary || "Calendar event").html()}</strong>
              <span>${event.status === "created" ? "Added to Google Calendar" : "Already on Google Calendar"}</span>
              <div class="result-links">
                <a href="${event.htmlLink}" target="_blank" rel="noopener noreferrer">Open in Google Calendar</a>
              </div>
            </div>
          `).join("")
        );
    } else {
      $("#import-result").prop("hidden", true).empty();
    }

    if (
      importedEvents.length === 1 &&
      existingEvents.length === 1
    ) {
      setStatus(
        `"${existingEvents[0].summary}" is already on Google Calendar. Nothing was added.`
      );
    } else if (
      importedEvents.length === 1 &&
      createdEvents.length === 1
    ) {
      setStatus(
        `"${createdEvents[0].summary}" was saved to Google Calendar.`
      );
    } else {
      setStatus(
        `${createdEvents.length} event${createdEvents.length === 1 ? "" : "s"} added; ` +
        `${existingEvents.length} already on Google Calendar.`
      );
    }

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