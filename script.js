// Detect system theme
const theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
document.documentElement.setAttribute("data-theme", theme);

let currentUser;
let lastLeaveCount;
const CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";
const members = [
  { email: "cm@localizedirect.com", possibleNames: ["chau"] },
  { email: "dng@localizedirect.com", possibleNames: ["duong"] },
  { email: "dp@localizedirect.com", possibleNames: ["dung"] },
  { email: "gn@localizedirect.com", possibleNames: ["giang"] },
  { email: "hh@localizedirect.com", possibleNames: ["hieu huynh", "hieu h", "hieu h."] },
  { email: "hm@localizedirect.com", possibleNames: ["huong"] },
  { email: "kp@localizedirect.com", possibleNames: ["khanh", "khanh p", "khanh ph"] },
  { email: "ld@localizedirect.com", possibleNames: ["lynh"] },
  { email: "ldv@localizedirect.com", possibleNames: ["long"] },
  { email: "nn@localizedirect.com", possibleNames: ["nha", "andy"] },
  { email: "nnc@localizedirect.com", possibleNames: ["cuong", "jason"] },
  { email: "pia@localizedirect.com", possibleNames: ["pia", "huyen"] },
  { email: "pv@localizedirect.com", possibleNames: ["phu"] },
  { email: "qv@localizedirect.com", possibleNames: ["quang"] },
  { email: "sn@localizedirect.com", possibleNames: ["sang"] },
  { email: "tc@localizedirect.com", possibleNames: ["tri truong", "tri t.", "steve"] },
  { email: "th@localizedirect.com", possibleNames: ["tan"] },
  { email: "tin@localizedirect.com", possibleNames: ["tin"] },
  { email: "tlv@localizedirect.com", possibleNames: ["win"] },
  { email: "tn@localizedirect.com", possibleNames: ["truong"] },
  { email: "tnn@localizedirect.com", possibleNames: ["thy"] },
  { email: "vtl@localizedirect.com", possibleNames: ["trong"] },
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
    const userInfoQuery = new URLSearchParams({ access_token: params["access_token"] }).toString();
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

async function getSpentLeaves() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  const selectedYear = Number(document.getElementById("year-select").value);
  const resultEl = document.getElementById("spent-leaves");
  const errorEl = document.getElementById("error-message");
  let dots = "";
  const loadingTimerId = setInterval(() => {
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
    }).toString();
    const eventListResponse = await fetch(`${eventListEndpoint}?${eventListQuery}`);
    const eventListData = await eventListResponse.json();
    if (!eventListResponse.ok) throw new Error(eventListData.error.message);

    // Object to track leave count of each member
    const leaveCountMap = {};
    const events = eventListData.items || [];
    for (const event of events) {
      const dayPart = /morning|afternoon/.test(event.summary) ? 0.5 : 1;
      const startDate = new Date(event.end.date).getTime();
      const endDate = new Date(event.start.date).getTime();
      const diffDays = Math.ceil((startDate - endDate) / (24 * 60 * 60 * 1000));
      const count = diffDays * dayPart;

      // Parse all members mentioned in event summary
      const eventMembers = event.summary.split("(off")[0].split(",");
      // Accumulate leave day in leaveCountMap
      for (const member of eventMembers) {
        const memberName = member.trim().toLowerCase();
        if (leaveCountMap.hasOwnProperty(memberName)) {
          leaveCountMap[memberName] += count;
        } else {
          leaveCountMap[memberName] = count;
        }
      }
    }

    let result = 0;
    const foundMember = members.find((member) => member.email === currentUser.email);
    if (foundMember) {
      // Accumulate leave day by current user 's possible names
      for (const possibleName of foundMember.possibleNames) {
        if (leaveCountMap.hasOwnProperty(possibleName)) {
          result += leaveCountMap[possibleName];
        }
      }
    } else {
      result = leaveCountMap[currentUser.given_name.toLowerCase()];
    }
    if (selectedYear === new Date().getFullYear()) {
      lastLeaveCount = result;
    }
    clearInterval(loadingTimerId);
    errorEl.innerHTML = "";
    resultEl.innerHTML = result;
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
    const loadingTimerId = setInterval(() => {
      dots += ".";
      if (dots.length === 11) dots = "";
      resultEl.innerHTML = `${dots} calculating ${dots}`;
    }, 100);
    const availableLeavesQuery = new URLSearchParams({ access_token: params["access_token"] }).toString();
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
  }).toString();
  location = `${oauth2Endpoint}?${oauth2Query}`;
}

function oauth2SignOut() {
  currentUser = null;
  localStorage.removeItem("oauth2-params");
  document.getElementById("get-spent-leaves-btn").classList.remove("hidden");
  document.getElementById("logged-in-section").classList.add("hidden");
}
