// Detect system theme
const theme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
document.documentElement.setAttribute("data-theme", theme);

let currentUser;
const CLIENT_ID = "81206403759-o2s2tkv3cl58c86njqh90crd8vnj6b82.apps.googleusercontent.com";
const CALENDAR_ID = "localizedirect.com_jeoc6a4e3gnc1uptt72bajcni8@group.calendar.google.com";
const members = [
  { email: "dng@localizedirect.com", possibleNames: ["duong"] },
  { email: "dp@localizedirect.com", possibleNames: ["dung"] },
  { email: "gn@localizedirect.com", possibleNames: ["giang"] },
  { email: "hh@localizedirect.com", possibleNames: ["hieu huynh", "hieu h", "hieu h."] },
  { email: "hm@localizedirect.com", possibleNames: ["huong"] },
  { email: "hn@localizedirect.com", possibleNames: ["hieu nguyen", "hieu ng", "hieu ng."] },
  { email: "kp@localizedirect.com", possibleNames: ["khanh", "khanh p"] },
  { email: "ld@localizedirect.com", possibleNames: ["lynh"] },
  { email: "ldv@localizedirect.com", possibleNames: ["long"] },
  { email: "tnn@localizedirect.com", possibleNames: ["thy"] },
  { email: "pia@localizedirect.com", possibleNames: ["pia"] },
  { email: "pv@localizedirect.com", possibleNames: ["phu"] },
  { email: "qv@localizedirect.com", possibleNames: ["quang"] },
  { email: "sn@localizedirect.com", possibleNames: ["sang"] },
  { email: "tc@localizedirect.com", possibleNames: ["tri truong", "tri t.", "steve"] },
  { email: "th@localizedirect.com", possibleNames: ["tan"] },
  { email: "tin@localizedirect.com", possibleNames: ["tin"] },
  { email: "tlv@localizedirect.com", possibleNames: ["win"] },
  { email: "tn@localizedirect.com", possibleNames: ["truong"] },
  { email: "tri@localizedirect.com", possibleNames: ["tri"] },
  { email: "vtl@localizedirect.com", possibleNames: ["trong"] },
];

const isLoggedIn = Boolean(localStorage.getItem("oauth2-params"));
if (isLoggedIn) {
  getMe();
  getLeaveCount();
  getRandomQuote();
} else {
  document.getElementById("get-leave-count-btn").classList.remove("hidden");
}

const hash = location.hash.substring(1);
const params = Object.fromEntries(new URLSearchParams(hash));
// Google oauth2 redirect handler, check state to mitigate csrf
if (params && params["state"] === localStorage.getItem("csrf-token")) {
  localStorage.setItem("oauth2-params", JSON.stringify(params));
  getMe();
  getLeaveCount();
  getRandomQuote();
  localStorage.removeItem("csrf-token");
  history.replaceState({}, null, "/"); // Url cleanup
}

async function getMe() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  try {
    const userInfoQuery = new URLSearchParams({ access_token: params["access_token"] }).toString();
    const userInfoEndpoint = "https://www.googleapis.com/oauth2/v3/userinfo";
    const userInfoResponse = await fetch(`${userInfoEndpoint}?${userInfoQuery}`);
    currentUser = await userInfoResponse.json();
    if (!userInfoResponse.ok) throw new Error(currentUser.error_description);
    document.getElementById("get-leave-count-btn").classList.add("hidden");
    document.getElementById("avatar").setAttribute("src", currentUser.picture);
    document.getElementById("welcome").innerHTML = `Welcome <b>${currentUser.given_name}</b>!`;
    document.getElementById("logged-in-section").classList.remove("hidden");
  } catch (error) {
    console.log(error.message);
    oauth2SignOut();
  }
}

async function getLeaveCount() {
  const params = JSON.parse(localStorage.getItem("oauth2-params"));
  if (!params || !params["access_token"]) return oauth2SignIn();

  const errorMessageElement = document.getElementById("error-message");
  let dots = "";
  const loadingTimerId = setInterval(() => {
    dots += ".";
    if (dots.length === 11) dots = "";
    errorMessageElement.innerHTML = `${dots} calculating ${dots}`;
  }, 100);
  try {
    const eventListEndpoint = `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`;
    const currentYear = new Date().getFullYear();
    const eventListQuery = new URLSearchParams({
      access_token: params["access_token"],
      timeMin: new Date(currentYear, 0, 1).toISOString(),
      timeMax: new Date(currentYear + 1, 0, 1).toISOString(),
      q: "off",
    }).toString();
    const eventListResponse = await fetch(`${eventListEndpoint}?${eventListQuery}`);
    const eventListData = await eventListResponse.json();
    if (!eventListResponse.ok) throw new Error(eventListData.error.message);

    // Object to track leave count of each member
    const leaveCountMap = {};
    const events = eventListData.items || [];
    for (const event of events) {
      // Parse day part from event summary
      const dayPart = event.summary.match(/morning|afternoon/g)?.[0] || "full";
      const dayPartByNumber = dayPart === "full" ? 1 : 0.5;
      const startDate = new Date(event.end.date).getTime();
      const endDate = new Date(event.start.date).getTime();
      const diffTime = Math.abs(endDate - startDate);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      const count = diffDays * dayPartByNumber;

      // Parse all members mentioned in event summary
      const eventMembers = event.summary.split("(off")[0].split(",");
      // Accumulate leave day in leaveCountMap
      for (const member of eventMembers) {
        const memberName = member.trim().toLowerCase();
        if (memberName in leaveCountMap) {
          leaveCountMap[memberName] += count;
        } else {
          leaveCountMap[memberName] = count;
        }
      }
    }
    console.log(leaveCountMap);

    let result = 0;
    const foundMember = members.find((member) => member.email === currentUser.email);
    // Accumulate leave day by current user 's possible names
    if (foundMember) {
      for (const possibleName of foundMember.possibleNames) {
        result += leaveCountMap[possibleName];
      }
    }
    clearInterval(loadingTimerId);
    errorMessageElement.innerHTML = "";
    document.getElementById("leave-count").innerHTML = `Your total leaves in ${currentYear} is <b>${result}</b> days`;
  } catch (error) {
    console.log(error.message);
    clearInterval(loadingTimerId);
    errorMessageElement.innerHTML = error.message;
    errorMessageElement.style.color = "red";
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
  localStorage.removeItem("oauth2-params");
  document.getElementById("get-leave-count-btn").classList.remove("hidden");
  document.getElementById("logged-in-section").classList.add("hidden");
}
