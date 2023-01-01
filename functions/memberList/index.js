import * as functions from "@google-cloud/functions-framework";

let sheetAccessToken = "";

functions.http("memberList", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");

  const query = new URLSearchParams(req.query);
  const clientAccessToken = query.get("access_token");
  if (!clientAccessToken) {
    return res.status(400).json({
      error: "required",
      message: "Required parameter is missing",
    });
  }
  try {
    const [userInfo, sheetValues] = await Promise.all([getUserInfo(clientAccessToken), getSheetValues()]);
    console.log(userInfo.name);
    const [header, ...rows] = sheetValues.values;
    const result = [];
    for (const rowValues of rows) {
      const obj = {};
      for (let i = 0; i < rowValues.length; i++) {
        const key = header[i];
        const value = rowValues[i];
        obj[key] = value;
      }
      result.push(obj);
    }
    return res.status(200).json(result);
  } catch (error) {
    return res.status(400).json(error);
  }
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

async function getSheetAccessToken() {
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
    sheetAccessToken = tokenObject.access_token;
  } catch (error) {
    throw error;
  }
}

async function getSheetValues() {
  try {
    if (sheetAccessToken) {
      console.log("use cached access token");
    } else {
      console.log("request new access token");
      await getSheetAccessToken();
    }
    const sheetName = "2023";
    const query = new URLSearchParams({
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "FORMATTED_STRING",
    });
    const sheetValuesResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${process.env.SPREADSHEET_ID}/values/${sheetName}?${query}`, {
      headers: { Authorization: `Bearer ${sheetAccessToken}` },
    });
    const sheetValues = await sheetValuesResponse.json();
    if (!sheetValuesResponse.ok) throw sheetValues;
    return sheetValues;
  } catch (error) {
    throw error;
  }
}
