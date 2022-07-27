// Detect system theme
const theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
document.documentElement.setAttribute("data-theme", theme);

let currentUser, lastLeaveCount, loadingTimerId;
const CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";
const members = [
  { email: "cm@localizedirect.com", names: ["Chau"] },
  { email: "dng@localizedirect.com", names: ["Duong Nguyen", "Duong"] },
  { email: "dpn@localizedirect.com", names: ["Duong Phung"] },
  { email: "dp@localizedirect.com", names: ["Dung"] },
  { email: "gn@localizedirect.com", names: ["Giang"] },
  { email: "hh@localizedirect.com", names: ["Hieu H.", "Hieu Huynh", "Hieu H"] },
  { email: "hm@localizedirect.com", names: ["Huong"] },
  { email: "kl@localizedirect.com", names: ["Khanh Le"] },
  { email: "kp@localizedirect.com", names: ["Khanh Pham", "Khanh", "Khanh P", "Khanh P."] },
  { email: "ld@localizedirect.com", names: ["Lynh"] },
  { email: "ldv@localizedirect.com", names: ["Long"] },
  { email: "nn@localizedirect.com", names: ["Andy", "Nha"] },
  { email: "nnc@localizedirect.com", names: ["Jason", "Cuong"] },
  { email: "pia@localizedirect.com", names: ["Pia", "Huyen"] },
  { email: "pv@localizedirect.com", names: ["Phu"] },
  { email: "qh@localizedirect.com", names: ["Quang Huynh"] },
  { email: "qv@localizedirect.com", names: ["Quang Vo", "Quang"] },
  { email: "sla@localizedirect.com", names: ["Son"] },
  { email: "sn@localizedirect.com", names: ["Sang"] },
  { email: "tc@localizedirect.com", names: ["Steve", "Tri Truong", "Tri T."] },
  { email: "th@localizedirect.com", names: ["Tan"] },
  { email: "tin@localizedirect.com", names: ["Tin"] },
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
  generateYearOptions();
  getRandomQuote();
  await getSpentLeaves();
  getAvailableLeaves();
}

function generateYearOptions() {
  const yearSelectEl = document.getElementById("year-select");
  const startYear = 2019;
  const currentYear = new Date().getFullYear();
  for (let y = startYear; y <= currentYear; y++) {
    const option = document.createElement("option");
    option.text = option.value = y;
    if (y === currentYear) option.selected = true;
    yearSelectEl.appendChild(option);
  }
}

async function getMe() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  try {
    const userInfoQuery = new URLSearchParams({ access_token: params["access_token"] });
    const userInfoEndpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const userInfoResponse = await fetch(`${userInfoEndpoint}?${userInfoQuery}`);
    const userInfoJson = await userInfoResponse.json();
    if (!userInfoResponse.ok) throw new Error(userInfoJson.error_description);
    currentUser = userInfoJson;
    document.getElementById("get-spent-leaves-btn").classList.add("hidden");
    document.getElementById("avatar").setAttribute("src", currentUser.picture);
    document.getElementById("welcome").innerHTML = `Welcome <b>${currentUser.given_name}</b>!`;
    document.getElementById("logged-in-section").classList.remove("hidden");
  } catch (error) {
    console.log(error.message);
  }
}

function getUserNames(email = currentUser.email) {
  const foundMember = members.find((member) => member.email === email);
  if (!foundMember) throw new Error("User not found");
  return foundMember.names;
}

function generateTimeText(date) {
  const day = new Date(date).toLocaleString("en-US", { day: "numeric", month: "short" });
  const weekday = new Date(date).toLocaleString("en-US", { weekday: "short" });
  return `${day} (${weekday})`;
}

function renderLeaveTable(events) {
  const table = document.getElementById("detail-table");
  // clear table
  table.replaceChildren();

  // table body
  events.forEach((event) => {
    const tr = table.insertRow();
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
    const diffDays = Math.ceil((new Date(event.end.date).getTime() - new Date(event.start.date).getTime()) / (24 * 60 * 60 * 1000));
    const count = diffDays * dayPartCount;
    const endDate = new Date(event.end.date);
    endDate.setDate(endDate.getDate() - 1);
    const reason = event.description ? JSON.parse(event.description).reason : "";

    const columns = [];
    columns.push(generateTimeText(event.start.date));
    columns.push(generateTimeText(endDate));
    columns.push(dayPartText);
    columns.push(count);
    columns.push(reason);

    columns.forEach((column) => {
      const td = tr.insertCell();
      td.innerText = column;
    });
  });

  // table header
  const thead = table.createTHead();
  const row = thead.insertRow(0);
  ["Start", "End", "Type", "Count", "Description"].forEach((header, index) => {
    cell = row.insertCell(index);
    cell.outerHTML = `<th>${header}</th>`;
  });
}

async function getSpentLeaves() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  const selectedYear = Number(document.getElementById("year-select").value);
  const resultEl = document.getElementById("spent-leaves");
  const errorEl = document.getElementById("error-message");
  let dots = "";
  loadingTimerId = setInterval(() => {
    dots += ".";
    if (dots.length === 11) dots = "";
    resultEl.innerHTML = `${dots} calculating ${dots}`;
  }, 100);

  try {
    const eventListEndpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
    const eventListQuery = new URLSearchParams({
      access_token: params["access_token"],
      timeMin: new Date(selectedYear, 0, 1).toISOString(),
      timeMax: new Date(selectedYear + 1, 0, 1).toISOString(),
      q: "off",
    });
    const eventListResponse = await fetch(`${eventListEndpoint}?${eventListQuery}`);
    const eventListData = await eventListResponse.json();
    if (!eventListResponse.ok) throw new Error(eventListData.error.message);

    let leaveCount = 0;
    const events = eventListData.items || [];
    const userEvents = [];
    const userNames = getUserNames();
    for (const event of events) {
      const dayPart = /morning|afternoon/.test(event.summary) ? 0.5 : 1;
      const startDate = new Date(event.end.date).getTime();
      const endDate = new Date(event.start.date).getTime();
      const diffDays = Math.ceil((startDate - endDate) / (24 * 60 * 60 * 1000));
      const count = diffDays * dayPart;

      // Parse all members mentioned in event summary
      const eventMembers = event.summary.split("(off")[0].split(",");
      // Count leaves
      for (const member of eventMembers) {
        const memberName = member.trim().toLowerCase();
        if (userNames.some((name) => name.toLowerCase() === memberName)) {
          userEvents.push(event);
          leaveCount += count;
        }
      }
    }
    renderLeaveTable(userEvents);
    if (selectedYear === new Date().getFullYear()) {
      lastLeaveCount = leaveCount;
    }
    clearInterval(loadingTimerId);
    errorEl.innerHTML = "";
    resultEl.innerHTML = leaveCount;
  } catch (error) {
    console.log(error.message);
    clearInterval(loadingTimerId);
    errorEl.innerHTML = error.message;
    resultEl.innerHTML = "";
    errorEl.style.color = "red";
  }
}

async function getRandomQuote() {
  try {
    const quoteResponse = await fetch("https://api.quotable.io/random");
    const quote = await quoteResponse.json();
    document.getElementById("quote-content").innerHTML = quote.content;
    document.getElementById("quote-author").innerHTML = quote.author;
  } catch (error) {
    console.log(error.message);
  }
}

async function getAvailableLeaves() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  try {
    const resultEl = document.getElementById("available-leaves");
    let dots = "";
    loadingTimerId = setInterval(() => {
      dots += ".";
      if (dots.length === 11) dots = "";
      resultEl.innerHTML = `${dots} calculating ${dots}`;
    }, 100);
    const availableLeavesQuery = new URLSearchParams({ access_token: params["access_token"] });
    const availableLeavesEndpoint = "https://asia-southeast1-my-project-1540367072726.cloudfunctions.net/availableLeaves";
    const availableLeavesResponse = await fetch(`${availableLeavesEndpoint}?${availableLeavesQuery}`);
    const availableLeavesJson = await availableLeavesResponse.json();
    clearInterval(loadingTimerId);
    resultEl.innerHTML = Number((availableLeavesJson.availableLeaves - lastLeaveCount).toFixed(1));
  } catch (error) {
    console.log(error.message);
    clearInterval(loadingTimerId);
  }
}

function oauth2SignIn() {
  // https://gist.github.com/6174/6062387
  const csrfToken = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
  localStorage.setItem("csrf-token", csrfToken);
  const oauth2Endpoint = "https://accounts.google.com/o/oauth2/v2/auth";
  const oauth2Query = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: location.href.replace(/\/$/, ""),
    scope: "profile email https://www.googleapis.com/auth/calendar.events.readonly",
    state: csrfToken,
    response_type: "token",
    hd: "localizedirect.com",
  });
  location = `${oauth2Endpoint}?${oauth2Query}`;
}

function oauth2SignOut() {
  currentUser = null;
  localStorage.removeItem("oauth2-params");
  document.getElementById("get-spent-leaves-btn").classList.remove("hidden");
  document.getElementById("logged-in-section").classList.add("hidden");
}
