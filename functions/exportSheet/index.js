import * as functions from "@google-cloud/functions-framework";
import * as fs from "fs";
import * as XLSX from "xlsx/xlsx.mjs";

XLSX.set_fs(fs);

const members = [
  { email: "ld@localizedirect.com", names: ["Lynh"] },
  { email: "tn@localizedirect.com", names: ["Truong"] },
  { email: "gn@localizedirect.com", names: ["Giang"], isAdmin: true },
  { email: "dng@localizedirect.com", names: ["Duong Nguyen"] },
  { email: "vtl@localizedirect.com", names: ["Trong"] },
  { email: "kp@localizedirect.com", names: ["Khanh Pham"] },
  { email: "th@localizedirect.com", names: ["Tan"] },
  { email: "tin@localizedirect.com", names: ["Tin"], isAdmin: true },
  { email: "hh@localizedirect.com", names: ["Hieu Huynh"] },
  { email: "sn@localizedirect.com", names: ["Sang"] },
  { email: "dp@localizedirect.com", names: ["Dung"] },
  { email: "qv@localizedirect.com", names: ["Quang Vo"] },
  { email: "pv@localizedirect.com", names: ["Phu"] },
  { email: "pia@localizedirect.com", names: ["Pia"], isAdmin: true },
  { email: "ldv@localizedirect.com", names: ["Long"] },
  { email: "hm@localizedirect.com", names: ["Huong"] },
  { email: "tc@localizedirect.com", names: ["Steve"] },
  { email: "tnn@localizedirect.com", names: ["Thy"] },
  { email: "nn@localizedirect.com", names: ["Andy", "Nha"] },
  { email: "nnc@localizedirect.com", names: ["Jason"] },
  { email: "cm@localizedirect.com", names: ["Chau Nguyen"] },
  { email: "sla@localizedirect.com", names: ["Son", "Son Le"] },
  { email: "qh@localizedirect.com", names: ["Quang Huynh"] },
  { email: "dpn@localizedirect.com", names: ["Duong Phung"] },
  { email: "kl@localizedirect.com", names: ["Khanh Le"] },
  { email: "tp@localizedirect.com", names: ["Thanh", "Thanh Phan"] },
  { email: "np@localizedirect.com", names: ["Ngan", "Ngan Phan"] },
  { email: "qt@localizedirect.com", names: ["Quoc", "Quoc Truong"] },
  { email: "cdm@localizedirect.com", names: ["Chau Dang"] },
  { email: "mn@localizedirect.com", names: ["Minh", "Minh Nguyen"] },
  { email: "cnp@localizedirect.com", names: ["Cuong Nguyen"] },
];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const initialMonthlyCounts = Array(12).fill(0);
let calendarAccessToken = "",
  spentObject = {
    // [email]: {
    //   totalCount: 10,
    //   monthlyCounts: [],
    //   events: [{ start: "", end: "", type: "", count: 0, description: "" }],
    // },
  };

functions.http("exportSheet", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");
  res.set("Access-Control-Expose-Headers", ["Content-Disposition"]);
  const { access_token } = req.query;
  const year = Number(req.query.year);
  if (!access_token || !year) {
    return res.status(400).json({
      error: "required",
      message: "Required parameter is missing",
    });
  }

  try {
    const userInfo = await getUserInfo(access_token);
    console.log(userInfo.name);
    const foundMember = members.find((member) => member.email === userInfo.email);
    if (!foundMember) {
      return res.status(403).json({
        error: "unauthorized",
        message: "No permission",
      });
    }

    await getCalendarEvents(year);

    const rows = [];
    // const width = { email: 0, name: 0, start: 0, end: 0, type: 0, count: 0, description: 0 };
    for (const member of members) {
      const array = [];
      const spentData = getSpentData(member.email);
      array.push(member.names[0]);
      for (const count of spentData.monthlyCounts) {
        array.push(count || "");
      }
      array.push(spentData.totalCount);
      rows.push(array);
    }

    var worksheet = XLSX.utils.json_to_sheet(rows);
    var workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

    // table header
    const headers = ["Name", ...MONTH_NAMES, "Total"];

    /* fix headers */
    XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: "A1" });

    /* calculate column width */
    // const max_width = rows.reduce((w, r) => Math.max(w, r.description.length), 10);
    // worksheet["!cols"] = [{ wch: width.email }, { wch: width.name }, { wch: width.start }, { wch: width.end }, { wch: width.type }, { wch: width.count }, { wch: width.description }];
    /* generate buffer */
    var buf = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    /* set headers */
    res.attachment(`Leave summary ${year}.xlsx`);
    /* respond with file data */
    res.status(200).send(buf);
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

async function getCalendarEvents(year) {
  if (calendarAccessToken) {
    console.log("use cached access token");
  } else {
    console.log("request new access token");
    await getCalendarAccessToken();
  }
  spentObject = {};
  try {
    const query = new URLSearchParams({
      access_token: calendarAccessToken,
      timeMin: new Date(year, 0, 1).toISOString(),
      timeMax: new Date(year + 1, 0, 1).toISOString(),
      q: "off",
      orderBy: "startTime",
      singleEvents: "true",
      maxResults: "2500",
    });
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
    for (const event of events) {
      let dayPartText = "",
        dayPartCount = 0;
      if (event.summary.includes("morning")) {
        dayPartText = "Morning";
        dayPartCount = 0.5;
      } else if (event.summary.includes("afternoon")) {
        dayPartText = "Afternoon";
        dayPartCount = 0.5;
      } else {
        dayPartText = "All day";
        dayPartCount = 1;
      }
      const startDate = new Date(event.start.date);
      const endDate = new Date(event.end.date);
      const monthlyCountObject = {};
      let eventDayCount = 0;
      // date range loop
      for (let d = new Date(startDate); d < endDate; d.setDate(d.getDate() + 1)) {
        const month = d.getMonth();
        if (!monthlyCountObject[month]) monthlyCountObject[month] = 0;
        monthlyCountObject[month] += dayPartCount;
        eventDayCount += dayPartCount;
      }

      endDate.setDate(endDate.getDate() - 1);
      const reason = event?.extendedProperties?.private?.reason || "";
      const newEvent = {
        start: generateTimeText(startDate),
        end: generateTimeText(endDate),
        type: dayPartText,
        count: eventDayCount,
        description: reason,
      };

      const eventMemberNames = event.summary
        .split("(off")[0]
        .split(",")
        .map((name) => name.trim());
      for (const memberName of eventMemberNames) {
        const foundMember = members.find((m) => m.names.includes(memberName));
        const email = foundMember ? foundMember.email : memberName;
        if (!spentObject[email]) spentObject[email] = { totalCount: 0, events: [], monthlyCounts: [...initialMonthlyCounts] };
        spentObject[email].totalCount += eventDayCount;
        spentObject[email].events.push(newEvent);
        for (const month in monthlyCountObject) {
          spentObject[email].monthlyCounts[month] += monthlyCountObject[month];
        }
      }
    }
  } catch (error) {
    throw error;
  }
}

function getSpentData(email) {
  return spentObject?.[email] || { totalCount: 0, events: [], monthlyCounts: [...initialMonthlyCounts] };
}

function generateTimeText(date) {
  const day = new Date(date).toLocaleString("en-US", { day: "numeric", month: "short" });
  const weekday = new Date(date).toLocaleString("en-US", { weekday: "short" });
  return `${day} (${weekday})`;
}
