import * as functions from "@google-cloud/functions-framework";

let calendarAccessToken = "";

functions.http("calendarEvents", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");

  const query = new URLSearchParams(req.query);
  const clientAccessToken = query.get("access_token");
  query.delete("access_token");
  if (!clientAccessToken) {
    return res.status(400).json({
      error: "required",
      message: "Required parameter is missing",
    });
  }
  let userInfo, events;
  try {
    [userInfo, events] = await Promise.all([getUserInfo(clientAccessToken), getCalendarEvents(query)]);
  } catch (error) {
    return res.status(400).json(error);
  }

  console.log(userInfo.name);
  return res.status(200).json(events);
});

async function getUserInfo(accessToken) {
  try {
    const userInfoEndpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const userInfoResponse = await fetch(`${userInfoEndpoint}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const userInfo = await userInfoResponse.json();
    if (!userInfoResponse.ok) throw userInfo;
    return userInfo;
  } catch (error) {
    throw error;
  }
}

async function getCalendarAccessToken() {
  try {
    const accessTokenResponse = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      body: new URLSearchParams({
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: "refresh_token",
        refresh_token: process.env.REFRESH_TOKEN,
      }),
    });
    const tokenObject = await accessTokenResponse.json();
    if (!accessTokenResponse.ok) throw tokenObject;
    calendarAccessToken = tokenObject.access_token;
  } catch (error) {
    throw error;
  }
}

async function getCalendarEvents(query) {
  try {
    if (calendarAccessToken) {
      console.log("use cached access token");
    } else {
      console.log("request new access token");
      await getCalendarAccessToken();
    }
    const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${process.env.CALENDAR_ID}/events`;
    let events = [];
    do {
      const response = await fetch(`${endpoint}?${query}`, {
        headers: { Authorization: `Bearer ${calendarAccessToken}` },
      });
      const data = await response.json();
      if (!response.ok) throw data;
      events = events.concat(data.items);
      query.set("pageToken", data.nextPageToken || "");
    } while (query.get("pageToken"));
    return events;
  } catch (error) {
    await getCalendarAccessToken();
    await getCalendarEvents(query);
    throw error;
  }
}
