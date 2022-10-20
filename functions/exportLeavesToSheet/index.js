import * as functions from "@google-cloud/functions-framework";
import * as fs from "fs";
import fetch from "node-fetch";
import * as XLSX from "xlsx/xlsx.mjs";

XLSX.set_fs(fs);

const members = [
  { email: "cm@localizedirect.com", names: ["Chau"] },
  { email: "dng@localizedirect.com", names: ["Duong Nguyen", "Duong"] },
  { email: "dp@localizedirect.com", names: ["Dung"] },
  { email: "dpn@localizedirect.com", names: ["Duong Phung"] },
  { email: "gn@localizedirect.com", names: ["Giang"], isAdmin: true },
  { email: "hh@localizedirect.com", names: ["Hieu Huynh"] },
  { email: "hm@localizedirect.com", names: ["Huong"] },
  { email: "kl@localizedirect.com", names: ["Khanh Le"] },
  { email: "kp@localizedirect.com", names: ["Khanh Pham", "Khanh"] },
  { email: "ld@localizedirect.com", names: ["Lynh"] },
  { email: "ldv@localizedirect.com", names: ["Long"] },
  { email: "nn@localizedirect.com", names: ["Andy", "Nha"] },
  { email: "nnc@localizedirect.com", names: ["Jason", "Cuong"] },
  { email: "np@localizedirect.com", names: ["Ngan Phan"] },
  { email: "pia@localizedirect.com", names: ["Pia", "Huyen"], isAdmin: true },
  { email: "pv@localizedirect.com", names: ["Phu"] },
  { email: "qh@localizedirect.com", names: ["Quang Huynh"] },
  { email: "qv@localizedirect.com", names: ["Quang Vo", "Quang"] },
  { email: "sla@localizedirect.com", names: ["Son"] },
  { email: "sn@localizedirect.com", names: ["Sang"] },
  { email: "tc@localizedirect.com", names: ["Steve", "Tri Truong"] },
  { email: "th@localizedirect.com", names: ["Tan"] },
  { email: "tin@localizedirect.com", names: ["Tin"], isAdmin: true },
  { email: "tp@localizedirect.com", names: ["Thanh Phan", "Thanh"] },
  { email: "tn@localizedirect.com", names: ["Truong"] },
  { email: "tnn@localizedirect.com", names: ["Thy"] },
  { email: "vtl@localizedirect.com", names: ["Trong"] },
];

function generateTimeText(date) {
  return new Date(date).toLocaleDateString("en-CA");
}

async function getGoogleUser(access_token) {
  try {
    const userInfoQuery = new URLSearchParams({ access_token });
    const userInfoEndpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const userInfoResponse = await fetch(`${userInfoEndpoint}?${userInfoQuery}`);
    const userInfo = await userInfoResponse.json();
    if (!userInfoResponse.ok) {
      throw new Error(JSON.stringify(userInfo));
    }
    console.log(userInfo.name);
  } catch (error) {
    console.log(error.message);
  }
}

functions.http("exportLeavesToSheet", async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");
  res.set("Access-Control-Expose-Headers", ["Content-Disposition"]);
  const { access_token } = req.query;
  if (!access_token) {
    return res.status(400).json({
      error: "required",
      message: "Required parameter is missing",
    });
  }

  getGoogleUser(access_token);

  const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${process.env.CALENDAR_ID}/events`;
  const query = new URLSearchParams({
    access_token,
    q: "off",
    orderBy: "startTime",
    singleEvents: true,
    maxResults: 2500,
  });
  let events = [];
  do {
    const response = await fetch(`${endpoint}?${query}`);
    const data = await response.json();
    if (!response.ok) {
      return res.status(response.status).json(data);
    }
    events = events.concat(data.items);
    query.set("pageToken", data.nextPageToken || "");
  } while (query.get("pageToken"));

  const rows = [];
  const width = { email: 0, name: 0, start: 0, end: 0, type: 0, count: 0, description: 0 };
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
    const diffDays = Math.ceil((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
    const count = diffDays * dayPartCount;
    endDate.setDate(endDate.getDate() - 1);
    const reason = event?.extendedProperties?.private?.reason || "";

    const eventMemberNames = event.summary
      .split("(off")[0]
      .split(",")
      .map((name) => name.trim());
    for (const memberName of eventMemberNames) {
      const foundMember = members.find((m) => m.names.includes(memberName));
      const email = foundMember ? foundMember.email : "";
      const newEvent = {
        email,
        name: memberName,
        start: generateTimeText(startDate),
        end: generateTimeText(endDate),
        type: dayPartText,
        count,
        description: reason,
      };
      rows.push(newEvent);
      for (const key in width) {
        width[key] = Math.max(width[key], newEvent[key].toString().length, key.length);
      }
    }
  }

  var worksheet = XLSX.utils.json_to_sheet(rows);
  var workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

  /* fix headers */
  XLSX.utils.sheet_add_aoa(worksheet, [["Email", "Name", "Start", "End", "Type", "Count", "Description"]], { origin: "A1" });

  /* calculate column width */
  // const max_width = rows.reduce((w, r) => Math.max(w, r.description.length), 10);
  worksheet["!cols"] = [{ wch: width.email }, { wch: width.name }, { wch: width.start }, { wch: width.end }, { wch: width.type }, { wch: width.count }, { wch: width.description }];
  /* generate buffer */
  var buf = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  /* set headers */
  res.attachment("VNLocalizeDirectLeave.xlsx");
  /* respond with file data */
  res.status(200).send(buf);
});
