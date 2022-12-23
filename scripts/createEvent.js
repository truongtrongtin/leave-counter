const GOOGLE_CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const GOOGLE_CLIENT_SECRET = "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";
const GOOGLE_REFRESH_TOKEN = "1//04Gges38qNV3dCgYIARAAGAQSNwF-L9IrZ1KdgBOiDZmWeGe6D7wY9iKLkrNvMavIjHIQ9bSIwkubMwGVi7v9TKJJin2XE-SJy-w";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";

async function create(startDateString, endDateString, summary) {
  const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    body: JSON.stringify({
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      grant_type: "refresh_token",
      refresh_token: GOOGLE_REFRESH_TOKEN,
    }),
  });
  const tokenObject = await tokenResponse.json();
  const accessToken = tokenObject.access_token;

  const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
  await fetch(endpoint, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify({
      start: { date: startDateString },
      end: { date: endDateString },
      summary,
      attendees: [
        {
          email: memberEmail,
          responseStatus: "accepted",
        },
      ],
      sendUpdates: "all",
      extendedProperties: {
        private: {
          message_ts: newMessage.message?.ts,
          ...(reason ? { reason } : {}),
        },
      },
      transparency: "transparent",
    }),
  });
}

create();
