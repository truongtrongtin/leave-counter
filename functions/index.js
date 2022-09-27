import * as functions from "@google-cloud/functions-framework";
import fetch from "node-fetch";

functions.http("availableLeaves", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");

  const { access_token } = req.query;
  const userInfoQuery = new URLSearchParams({ access_token });
  const userInfoEndpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
  const userInfoResponse = await fetch(`${userInfoEndpoint}?${userInfoQuery}`);
  const userInfo = await userInfoResponse.json();
  if (!userInfoResponse.ok) {
    res.status(userInfoResponse.status).json(userInfo);
    return;
  }
  console.log(userInfo.name);

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
  if (!accessTokenResponse.ok) {
    res.status(accessTokenResponse.status).json(tokenObject);
    return;
  }

  const sheetValuesQuery = new URLSearchParams();
  sheetValuesQuery.append("ranges", "1:1");
  sheetValuesQuery.append("ranges", "18:18");
  sheetValuesQuery.append("ranges", "22:22");
  const sheetValuesResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${process.env.SPREADSHEET_ID}/values:batchGet?${sheetValuesQuery}`, {
    headers: { Authorization: `Bearer ${tokenObject.access_token}` },
  });
  const sheetValues = await sheetValuesResponse.json();
  if (!sheetValuesResponse.ok) {
    res.status(sheetValuesResponse.status).json(tokenObject);
    return;
  }

  const memberCodes = sheetValues.valueRanges[0].values[0];
  const allAvailableLeaves = sheetValues.valueRanges[1].values[0];
  const allAvailableLeavesThisYear = sheetValues.valueRanges[2].values[0];
  const foundIndex = memberCodes.findIndex((code) => code === userInfo.email.split("@localizedirect.com")[0]);
  const availableLeaves = Number(allAvailableLeaves[foundIndex]) || 0;
  const availableLeavesThisYear = Number(allAvailableLeavesThisYear[foundIndex]) || 0;

  res.status(200).json({ email: userInfo.email, availableLeaves: availableLeaves + availableLeavesThisYear });
});
