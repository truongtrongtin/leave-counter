import * as functions from "@google-cloud/functions-framework";

let sheetAccessToken = "";

functions.http("availableLeaves", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");

  const query = new URLSearchParams(req.query);
  const clientAccessToken = query.get("access_token");
  if (!clientAccessToken) {
    return res.status(400).json({
      error: "required",
      message: "Required parameter is missing",
    });
  }
  let userInfo, sheetValues;
  try {
    [userInfo, sheetValues] = await Promise.all([getUserInfo(clientAccessToken), getSheetValues()]);
  } catch (error) {
    return res.status(400).json(error);
  }

  console.log(userInfo.name);
  const memberCodes = sheetValues.valueRanges[0].values[0];
  const allAvailableLeaves = sheetValues.valueRanges[1].values[0];
  const result = {};
  for (let i = 0; i < memberCodes.length; i++) {
    const email = memberCodes[i] + "@localizedirect.com";
    result[email] = Number(allAvailableLeaves[i]) || 0;
  }
  return res.status(200).json(result);
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
    const sheetValuesQuery = new URLSearchParams();
    sheetValuesQuery.append("ranges", "Sheet1!B1:1");
    sheetValuesQuery.append("ranges", "Sheet1!B19:19");
    const sheetValuesResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${process.env.SPREADSHEET_ID}/values:batchGet?${sheetValuesQuery}`, {
      headers: { Authorization: `Bearer ${sheetAccessToken}` },
    });
    const sheetValues = await sheetValuesResponse.json();
    if (!sheetValuesResponse.ok) throw sheetValues;
    return sheetValues;
  } catch (error) {
    await getSheetAccessToken();
    await getSheetValues();
    throw error;
  }
}
