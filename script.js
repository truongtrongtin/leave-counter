// Detect system theme
const theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
document.documentElement.setAttribute("data-theme", theme);

let currentUser = null,
  availableObject = {},
  spentObject = {
    // email: {
    //   year: {
    //     totalCount: 10,
    //     monthlyCounts: [],
    //     events: [{ start: "", end: "", type: "", count: 10, description: "" }],
    //   },
    // },
  };
const thisYear = new Date().getFullYear();
const CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";
const AVAILABLE_LEAVES_URL = "https://available-leaves-yaxjnhmzuq-as.a.run.app";
const EXPORT_SHEET_URL = "https://export-leaves-to-sheet-yaxjnhmzuq-as.a.run.app";
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
  { email: "tc@localizedirect.com", names: ["Steve", "Tri Truong"] },
  { email: "tnn@localizedirect.com", names: ["Thy"] },
  { email: "nn@localizedirect.com", names: ["Andy", "Nha"] },
  { email: "nnc@localizedirect.com", names: ["Jason", "Cuong"] },
  { email: "cm@localizedirect.com", names: ["Chau Nguyen"] },
  { email: "sla@localizedirect.com", names: ["Son Le"] },
  { email: "qh@localizedirect.com", names: ["Quang Huynh"] },
  { email: "dpn@localizedirect.com", names: ["Duong Phung"] },
  { email: "kl@localizedirect.com", names: ["Khanh Le"] },
  { email: "tp@localizedirect.com", names: ["Thanh Phan", "Thanh"] },
  { email: "np@localizedirect.com", names: ["Ngan Phan"] },
  { email: "qt@localizedirect.com", names: ["Quoc Truong", "Quoc"] },
  { email: "cdm@localizedirect.com", names: ["Chau Dang"] },
];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const initialMonthlyCounts = Array(12).fill(0);

const hash = location.hash.substring(1);
const params = Object.fromEntries(new URLSearchParams(hash));
if (localStorage.getItem("oauth2-params")) {
  // Has token case
  main();
} else if (params["state"] === localStorage.getItem("csrf-token")) {
  // Google oauth2 redirect case, check state to mitigate csrf
  localStorage.setItem("oauth2-params", JSON.stringify(params));
  localStorage.removeItem("csrf-token");
  history.replaceState({}, null, "/"); // Url cleanup
  main();
} else {
  document.getElementById("get-spent-leaves-btn").classList.remove("hidden");
}

async function main() {
  await getMe();
  if (!currentUser) return;
  buildModeSelect();
  buildYearSelect();
  buildMemberSelect();
  buildModeSelect();
  getAndShowRandomQuote();

  const remainCountEl = document.getElementById("available-leaves");
  let dots = "";
  const loadingTimerId = setInterval(() => {
    dots += ".";
    if (dots.length === 11) dots = "";
    remainCountEl.innerHTML = `${dots} loading ${dots}`;
  }, 100);
  await Promise.all([onYearChange(), getAvailableData()]);
  clearInterval(loadingTimerId);
  buildMultipleTable();
  showAvailableDays();
}

function buildModeSelect() {
  const modeSelectEl = document.getElementById("mode-select");
  const isAdmin = getLocalMember(currentUser.email)?.isAdmin;
  if (!isAdmin) return;
  modeSelectEl.classList.remove("hidden");
}

function buildMemberSelect() {
  const isAdmin = getLocalMember(currentUser.email)?.isAdmin;
  if (!isAdmin) return;
  const memberSelectEl = document.getElementById("member-select");
  memberSelectEl.classList.remove("hidden");
  for (const member of members) {
    const option = document.createElement("option");
    option.text = member.names[0];
    option.value = member.email;
    if (member.email === currentUser.email) option.selected = true;
    memberSelectEl.appendChild(option);
  }
}

function buildYearSelect() {
  document.getElementById("by-year").innerText = thisYear;
  const yearSelectEl = document.getElementById("year-select");
  const startYear = 2019;
  for (let year = startYear; year <= thisYear; year++) {
    const option = document.createElement("option");
    option.text = option.value = year;
    if (year === thisYear) option.selected = true;
    yearSelectEl.appendChild(option);
  }
}

function getAccessToken() {
  return JSON.parse(localStorage.getItem("oauth2-params"))?.["access_token"] || "";
}

async function getMe() {
  try {
    const accessToken = getAccessToken();
    const query = new URLSearchParams({ access_token: accessToken });
    const endpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const response = await fetch(`${endpoint}?${query}`);
    const userInfo = await response.json();
    if (response.status === 401) {
      oauth2SignOut();
      oauth2SignIn();
      return;
    }
    if (!response.ok) throw new Error(userInfo.error_description);
    currentUser = userInfo;
    document.getElementById("get-spent-leaves-btn").classList.add("hidden");
    document.getElementById("avatar").setAttribute("src", currentUser.picture);
    document.getElementById("welcome").innerHTML = `Welcome <b>${currentUser.given_name}</b>!`;
    document.getElementById("logged-in-section").classList.remove("hidden");
  } catch (error) {
    showError(error.message);
  }
}

function getLocalMember(email) {
  return members.find((member) => member.email === email);
}

function getSelectedEmail() {
  return document.getElementById("member-select").value || currentUser.email;
}

function getSelectedYear() {
  return Number(document.getElementById("year-select").value) || thisYear;
}

function getSpentData(year, email) {
  const selectedEmail = email || getSelectedEmail();
  const selectedYear = year || getSelectedYear();
  return spentObject?.[selectedYear]?.[selectedEmail] || { totalCount: 0, events: [] };
}

function getAvailableValue(email) {
  const selectedEmail = email || getSelectedEmail();
  return availableObject[selectedEmail] || 0;
}

function generateTimeText(date) {
  const day = new Date(date).toLocaleString("en-US", { day: "numeric", month: "short" });
  const weekday = new Date(date).toLocaleString("en-US", { weekday: "short" });
  return `${day} (${weekday})`;
}

function showSpentCount() {
  const spentCountEl = document.getElementById("spent-leaves");
  spentCountEl.innerHTML = getSpentData().totalCount;
}

function showAvailableDays() {
  const remainCountEl = document.getElementById("available-leaves");
  remainCountEl.innerHTML = getAvailableValue() - getSpentData(thisYear).totalCount;
}

function buildSingleSpentTable() {
  const table = document.getElementById("single-table");
  // clear table
  table.replaceChildren();
  // table body
  const memberEvents = getSpentData().events;
  for (const event of memberEvents) {
    const tr = table.insertRow();
    for (const key in event) {
      const td = tr.insertCell();
      td.innerText = event[key];
    }
  }
  // table header
  const thead = table.createTHead();
  const row = thead.insertRow(0);
  const headers = ["Start", "End", "Type", "Count", "Description"];
  for (let i = 0; i < headers.length; i++) {
    cell = row.insertCell(i);
    cell.outerHTML = `<th>${headers[i]}</th>`;
  }
}

function buildMultipleTable() {
  const table = document.getElementById("multiple-table");
  // clear table
  table.replaceChildren();
  // table body
  for (const member of members) {
    const spentData = getSpentData(null, member.email);
    const monthlyCounts = spentData.monthlyCounts || initialMonthlyCounts;
    const tr = table.insertRow();
    tr.insertCell().innerText = member.names[0];
    for (const count of monthlyCounts) {
      const td = tr.insertCell();
      td.innerText = count || "";
    }
    tr.insertCell().innerText = spentData.totalCount;
    tr.insertCell().innerText = getAvailableValue(member.email) - getSpentData(thisYear, member.email).totalCount;
  }
  // table header
  const thead = table.createTHead();
  const row = thead.insertRow(0);
  const headers = ["Name", ...MONTH_NAMES, "Spent", "Avail"];
  for (let i = 0; i < headers.length; i++) {
    cell = row.insertCell(i);
    cell.outerHTML = `<th>${headers[i]}</th>`;
  }
}

async function getCalendarEvents() {
  let loadingTimerId;
  try {
    const accessToken = getAccessToken();
    const selectedYear = getSelectedYear();
    if (spentObject[selectedYear]) return;
    const spentCountEl = document.getElementById("spent-leaves");

    let dots = "";
    loadingTimerId = setInterval(() => {
      dots += ".";
      if (dots.length === 11) dots = "";
      spentCountEl.innerHTML = `${dots} loading ${dots}`;
    }, 100);

    const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
    const query = new URLSearchParams({
      access_token: accessToken,
      timeMin: new Date(selectedYear, 0, 1).toISOString(),
      timeMax: new Date(selectedYear + 1, 0, 1).toISOString(),
      q: "off",
      orderBy: "startTime",
      singleEvents: "true",
      maxResults: "2500",
    });
    let events = [];
    do {
      const response = await fetch(`${endpoint}?${query}`);
      const data = await response.json();
      if (!response.ok) throw new Error(data.error.message);
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
        if (!spentObject[selectedYear]) spentObject[selectedYear] = {};
        if (!spentObject[selectedYear][email]) spentObject[selectedYear][email] = { totalCount: 0, events: [], monthlyCounts: [...initialMonthlyCounts] };
        spentObject[selectedYear][email].totalCount += eventDayCount;
        spentObject[selectedYear][email].events.push(newEvent);
        for (const month in monthlyCountObject) {
          spentObject[selectedYear][email].monthlyCounts[month] += monthlyCountObject[month];
        }
      }
    }

    clearInterval(loadingTimerId);
  } catch (error) {
    clearInterval(loadingTimerId);
    showError(error.message);
  }
}

function onModeChange() {
  const mode = document.getElementById("mode-select").value;
  const memberSelect = document.getElementById("member-select");
  const availableSection = document.getElementById("available-section");
  const singleSpentSection = document.getElementById("single-spent-section");
  const multipleSpentSection = document.getElementById("multiple-spent-section");
  switch (mode) {
    case "single":
      memberSelect.disabled = false;
      availableSection.classList.remove("hidden");
      singleSpentSection.classList.remove("hidden");
      multipleSpentSection.classList.add("hidden");
      break;
    case "multiple":
      memberSelect.disabled = true;
      availableSection.classList.add("hidden");
      singleSpentSection.classList.add("hidden");
      multipleSpentSection.classList.remove("hidden");
      break;
    default:
      break;
  }
}

async function onYearChange() {
  await getCalendarEvents();
  showSpentCount();
  buildSingleSpentTable();
  buildMultipleTable();
}

function onMemberChange() {
  showSpentCount();
  buildSingleSpentTable();
  showAvailableDays();
}

async function getAndShowRandomQuote() {
  const quoteContent = document.getElementById("quote-content");
  const quoteAuthor = document.getElementById("quote-author");
  try {
    const response = await fetch("https://api.quotable.io/random");
    const quote = await response.json();
    quoteContent.innerHTML = quote.content;
    quoteAuthor.innerHTML = quote.author;
  } catch (error) {
    quoteContent.innerHTML = error.message;
  }
}

async function getAvailableData() {
  try {
    const accessToken = getAccessToken();
    const query = new URLSearchParams({ access_token: accessToken });
    const response = await fetch(`${AVAILABLE_LEAVES_URL}?${query}`);
    availableObject = await response.json();
  } catch (error) {
    throw error;
  }
}

async function downloadSheet() {
  const accessToken = getAccessToken();
  const downloadBtn = document.getElementById("download");
  downloadBtn.disabled = true;
  downloadBtn.textContent = "Exporting...";
  try {
    const response = await fetch(`${EXPORT_SHEET_URL}?access_token=${accessToken}`);
    const filename = response.headers.get("Content-Disposition").split('"')[1];
    const fileBlob = await response.blob();
    const url = URL.createObjectURL(fileBlob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
  } catch (error) {
    showError(error.message);
  }
  downloadBtn.disabled = false;
  downloadBtn.textContent = "Export";
}

function oauth2SignIn() {
  // https://gist.github.com/6174/6062387
  const csrfToken = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
  localStorage.setItem("csrf-token", csrfToken);
  const endpoint = "https://accounts.google.com/o/oauth2/v2/auth";
  const query = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: location.href.replace(/\/$/, ""),
    scope: "profile email https://www.googleapis.com/auth/calendar.events.readonly",
    state: csrfToken,
    response_type: "token",
    hd: "localizedirect.com",
  });
  location = `${endpoint}?${query}`;
}

function oauth2SignOut() {
  currentUser = null;
  localStorage.removeItem("oauth2-params");
  document.getElementById("get-spent-leaves-btn").classList.remove("hidden");
  document.getElementById("logged-in-section").classList.add("hidden");
}

function showError(message) {
  document.getElementById("dialog").showModal();
  document.getElementById("error-message").innerHTML = message;
}
