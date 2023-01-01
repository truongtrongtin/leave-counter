const GOOGLE_CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const GOOGLE_CLIENT_SECRET = "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";
const GOOGLE_REFRESH_TOKEN = "1//04Gges38qNV3dCgYIARAAGAQSNwF-L9IrZ1KdgBOiDZmWeGe6D7wY9iKLkrNvMavIjHIQ9bSIwkubMwGVi7v9TKJJin2XE-SJy-w";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";

async function migrate(oldName, newName) {
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
  console.log(JSON.stringify(tokenObject, null, 2));
  const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
  const query = new URLSearchParams({
    q: oldName,
    orderBy: "startTime",
    singleEvents: "true",
    maxResults: "2500",
  });
  let events = [];
  do {
    const response = await fetch(`${endpoint}?${query}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const data = await response.json();
    events = events.concat(data.items || []);
    query.set("pageToken", data.nextPageToken || "");
  } while (query.get("pageToken"));
  const filtedEvents = events.filter((event) => event.summary.includes(oldName));

  const summaries = filtedEvents.map((event) => event.summary);
  console.log(summaries);
  console.log(summaries.length);

  for (const event of filtedEvents) {
    const summary = event.summary.replace(oldName, newName);
    const result = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events/${event.id}`, {
      method: "PUT",
      headers: { Authorization: `Bearer ${accessToken}` },
      body: JSON.stringify({ ...event, summary }),
    });
    const json = await result.json();
    console.log(event.summary, "-->", json.summary);
  }
}

migrate("Son Le", "Son");
