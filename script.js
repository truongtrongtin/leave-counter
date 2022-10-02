// Detect system theme
const theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
document.documentElement.setAttribute("data-theme", theme);

let currentUser = null,
  availableLeaves = [],
  spentObject = {
    // email: { year: { totalCount: 10, events: [{ start: "", end: "", type: "", count: 10, description: "" }] } }
  };
const thisYear = new Date().getFullYear();
const CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";
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
  if (!currentUser) {
    oauth2SignOut();
    oauth2SignIn();
    return;
  }
  buildYearSelect();
  buildMemberSelect();
  getAndShowRandomQuote();

  const remainCountEl = document.getElementById("available-leaves");
  let dots = "";
  const loadingTimerId = setInterval(() => {
    dots += ".";
    if (dots.length === 11) dots = "";
    remainCountEl.innerHTML = `${dots} loading ${dots}`;
  }, 100);
  try {
    await Promise.all([onYearChange(), getAvailableLeaves()]);
    clearInterval(loadingTimerId);
    showAvailableDays();
  } catch (error) {
    clearInterval(loadingTimerId);
    console.log(error.message);
  }
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
  yearSelectEl.classList.remove("hidden");
  const startYear = 2019;
  for (let year = startYear; year <= thisYear; year++) {
    const option = document.createElement("option");
    option.text = option.value = year;
    if (year === thisYear) option.selected = true;
    yearSelectEl.appendChild(option);
  }
}

async function getMe() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  try {
    const query = new URLSearchParams({ access_token: params["access_token"] });
    const endpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const response = await fetch(`${endpoint}?${query}`);
    const userInfo = await response.json();
    if (!response.ok) throw new Error(userInfo.error_description);
    currentUser = userInfo;
    document.getElementById("get-spent-leaves-btn").classList.add("hidden");
    document.getElementById("avatar").setAttribute("src", currentUser.picture);
    document.getElementById("welcome").innerHTML = `Welcome <b>${currentUser.given_name}</b>!`;
    document.getElementById("logged-in-section").classList.remove("hidden");
  } catch (error) {
    console.log(error.message);
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

function generateTimeText(date) {
  const day = new Date(date).toLocaleString("en-US", { day: "numeric", month: "short" });
  const weekday = new Date(date).toLocaleString("en-US", { weekday: "short" });
  return `${day} (${weekday})`;
}

function showSpentCount() {
  const spentCountEl = document.getElementById("spent-leaves");
  spentCountEl.innerHTML = getSpentData().totalCount;
}

function showSpentTable() {
  const table = document.getElementById("detail-table");
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

async function getCalendarEvents() {
  const selectedYear = getSelectedYear();
  if (spentObject[selectedYear]) return;
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  const spentCountEl = document.getElementById("spent-leaves");
  const errorEl = document.getElementById("error-message");

  let dots = "";
  const loadingTimerId = setInterval(() => {
    dots += ".";
    if (dots.length === 11) dots = "";
    spentCountEl.innerHTML = `${dots} loading ${dots}`;
  }, 100);

  try {
    const endpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
    const query = new URLSearchParams({
      access_token: params["access_token"],
      timeMin: new Date(selectedYear, 0, 1).toISOString(),
      timeMax: new Date(selectedYear + 1, 0, 1).toISOString(),
      q: "off",
      orderBy: "startTime",
      singleEvents: true,
    });
    const response = await fetch(`${endpoint}?${query}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error.message);

    for (const event of data.items) {
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
      const reason = event.description ? JSON.parse(event.description).reason : "";

      const newEvent = {
        start: generateTimeText(startDate),
        end: generateTimeText(endDate),
        type: dayPartText,
        count,
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
        if (!spentObject[selectedYear][email]) spentObject[selectedYear][email] = { totalCount: 0, events: [] };
        spentObject[selectedYear][email].totalCount += count;
        spentObject[selectedYear][email].events.push(newEvent);
      }
    }

    clearInterval(loadingTimerId);
    errorEl.innerHTML = "";
  } catch (error) {
    console.log(error.message);
    clearInterval(loadingTimerId);
    errorEl.innerHTML = error.message;
    spentCountEl.innerHTML = "";
    errorEl.style.color = "red";
  }
}

async function onYearChange() {
  await getCalendarEvents();
  showSpentCount();
  showSpentTable();
}

function onMemberChange() {
  showSpentCount();
  showSpentTable();
  showAvailableDays();
}

async function getAndShowRandomQuote() {
  try {
    const response = await fetch("https://api.quotable.io/random");
    const quote = await response.json();
    document.getElementById("quote-content").innerHTML = quote.content;
    document.getElementById("quote-author").innerHTML = quote.author;
  } catch (error) {
    console.log(error.message);
  }
}

async function getAvailableLeaves() {
  try {
    const params = JSON.parse(localStorage.getItem("oauth2-params"));
    const query = new URLSearchParams({ access_token: params["access_token"] });
    const endpoint = "https://available-leaves-yaxjnhmzuq-as.a.run.app";
    const response = await fetch(`${endpoint}?${query}`);
    availableLeaves = await response.json();
  } catch (error) {
    throw error;
  }
}

function showAvailableDays() {
  const memberEmail = document.getElementById("member-select").value || currentUser.email;
  const count = availableLeaves.find((leave) => leave.email === memberEmail).value;
  const remainCountEl = document.getElementById("available-leaves");
  remainCountEl.innerHTML = count - getSpentData(thisYear).totalCount;
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
