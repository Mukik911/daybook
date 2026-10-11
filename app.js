(() => {
"use strict";

// ───────────────────────── constants ─────────────────────────
const CFG = window.DAYBOOK_CONFIG || {};
const SCOPES = ["User.Read", "Calendars.ReadWrite"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const CATEGORY = "Daybook";
const DONE_CATEGORY = "Daybook: done";
// Custom MAPI property that carries Daybook's own fields on each Outlook event.
const PROP_ID = "String {8d3f1a62-5b7c-4e19-9a0d-2c6e4b7f9013} Name DaybookMeta";
const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
const REDIRECT_URI = new URL("./", location.href).href;

const QUAD = { q1: "Do now", q2: "Schedule", q3: "Delegate", q4: "Later / maybe" };
const ENERGY = { deep: "Deep focus", light: "Light", admin: "Admin" };
const REPEAT = { none: "", daily: "Daily", weekdays: "Weekdays", weekly: "Weekly", monthly: "Monthly", custom: "Repeats" };
const DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
const GROUPS = ["Overdue", "Today", "Tomorrow", "Next 7 days", "Later", "Done today"];
const LOG_CAP = 30; // history entries kept per task (stored on the Outlook event)
const FIELD_LABEL = { title: "Title", quad: "Priority", cat: "Area", energy: "Energy", date: "Date", time: "Start",
  dur: "Duration", remind: "Reminder", repeat: "Repeat", dod: "Done when", notes: "Notes", busy: "Show as" };

// ───────────────────────── helpers ─────────────────────────
const $ = (s) => document.querySelector(s);
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};
const pad = (n) => String(n).padStart(2, "0");
const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const todayStr = () => ymd(new Date());
const addDays = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return ymd(d); };
const localIso = (d) => `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
function startOf(t) {
  const [y, m, d] = t.date.split("-").map(Number);
  const [h, mi] = (t.time || "00:00").split(":").map(Number);
  return new Date(y, m - 1, d, h, mi);
}
function endTime(t) { const e = new Date(startOf(t).getTime() + t.dur * 60000); return `${pad(e.getHours())}:${pad(e.getMinutes())}`; }
function fmtDate(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: "short", day: "numeric", month: "short" }); }
const newId = () => "l" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

function el(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") n.className = v;
    else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
    else if (v !== false && v != null) n.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) n.append(c.nodeType ? c : document.createTextNode(c));
  return n;
}

let toastTimer;
function toast(msg, isError = false) {
  const t = $("#toast");
  t.textContent = msg; t.classList.toggle("error", isError); t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), isError ? 6000 : 3200);
}

// ───────────────────────── state ─────────────────────────
const settings = {
  clientId: ls.get("daybook.clientId", "") || CFG.clientId || "",
  tenant: ls.get("daybook.tenant", "") || CFG.tenant || "organizations",
  linkAcct: ls.get("daybook.linkAcct", "work"),
};
const state = {
  account: null,
  outlook: ls.get("daybook.cache", []),   // tasks mirrored from Outlook (offline copy)
  local: ls.get("daybook.local", []),     // tasks that live only on this device
  filter: "open",
  query: "",
  editing: null,                          // task being edited
  pendingDelete: new Set(),
  syncing: false,
  lastSync: 0,
  view: "tasks",                          // tasks | archive | stats | log
  period: 7,                              // analytics window in days
  logFilter: "all",
  archQuery: "",
  openHistory: new Set(),                 // archive rows with history expanded
};
const signedIn = () => !!state.account;
const visibleTasks = () => (signedIn() ? state.outlook : state.local);
const saveLocal = () => ls.set("daybook.local", state.local);
const saveCache = () => { if (!DEMO) ls.set("daybook.cache", state.outlook); };

// ───────────────────────── auth (MSAL) ─────────────────────────
let pca = null;
async function initAuth() {
  if (DEMO) { state.account = { username: "demo calendar" }; state.outlook = []; seedDemo(); return; }
  if (!settings.clientId || !window.msal) return;
  pca = new msal.PublicClientApplication({
    auth: {
      clientId: settings.clientId,
      authority: `https://login.microsoftonline.com/${settings.tenant || "organizations"}`,
      redirectUri: REDIRECT_URI,
      postLogoutRedirectUri: REDIRECT_URI,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await pca.initialize();
  const res = await pca.handleRedirectPromise();
  if (res && res.account) pca.setActiveAccount(res.account);
  const acct = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  if (acct) { pca.setActiveAccount(acct); state.account = acct; }
}

const PERSONAL_DOMAIN = /@(outlook|hotmail|live|msn)\.[a-z.]+$/i;

async function signIn(email) {
  if (!pca) throw new Error("Daybook isn't connected to Microsoft yet. Open Connection settings and add the Application (client) ID (see SETUP.md).");
  if (email && PERSONAL_DOMAIN.test(email) && settings.tenant === "organizations") {
    throw new Error('This is a personal Microsoft account. To use it, set Tenant to "common" in Connection settings.');
  }
  // Redirect works in installed iOS / Windows apps, where pop-ups don't.
  // loginHint pre-fills the email on Microsoft's sign-in page.
  await pca.loginRedirect(email ? { scopes: SCOPES, loginHint: email } : { scopes: SCOPES, prompt: "select_account" });
}

async function signOut() {
  if (!pca || !state.account) return;
  // Clears this app's session only; you stay signed in to Outlook elsewhere.
  await pca.logoutRedirect({ account: state.account, onRedirectNavigate: () => false });
  state.account = null; state.outlook = []; saveCache();
  ls.set("daybook.localMode", false);
  closeSettings(); renderChrome(); render();
  showGate();
}

// ── sign-in screen ──
const needsGate = () => !signedIn() && !DEMO && !ls.get("daybook.localMode", false);
function showGate(error) {
  document.body.classList.add("gated");
  $("#gate").hidden = false;
  $("#g-email").value = ls.get("daybook.email", "");
  const err = $("#gate-err"); err.hidden = !error; err.textContent = error || "";
  const btn = $("#gate-submit"); btn.disabled = false; btn.textContent = "Continue";
  setTimeout(() => $("#g-email").focus(), 50);
}
function hideGate() {
  document.body.classList.remove("gated");
  $("#gate").hidden = true;
}

async function accessToken() {
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account: state.account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect({ scopes: SCOPES, account: state.account });
      return new Promise(() => {}); // page is navigating away
    }
    throw e;
  }
}

// ───────────────────────── demo calendar (?demo) ─────────────────────────
// Open the app with ?demo to try the calendar features with a made-up calendar.
// Nothing is sent to Microsoft in this mode.
const DEMO = new URLSearchParams(location.search).has("demo");
const demoEvents = [];
function demoDay(date) {
  const [y, m, d] = date.split("-").map(Number);
  const wd = new Date(y, m - 1, d).getDay();
  if (wd === 0 || wd === 6) return [];
  const ev = (id, subject, s, e, showAs = "busy") => ({
    id: `demo-${date}-${id}`, subject, showAs, isAllDay: false, categories: [],
    start: { dateTime: `${date}T${s}:00`, timeZone: TZ }, end: { dateTime: `${date}T${e}:00`, timeZone: TZ } });
  if (d % 5 === 0) return [ev("ws", "Quarterly planning workshop", "09:00", "18:00")];
  return [ev("su", "Daily stand-up", "09:30", "09:45"), ev("ps", "Product sync", "11:00", "12:00"),
          ev("lu", "Lunch", "13:00", "14:00", "free"), ev("oo", "1:1 with manager", "15:00", "15:30"),
          ev("rv", "Design review", "15:15", "16:00", "tentative")];
}
// Sample Daybook tasks spread over the past three weeks and the next ten days.
function seedDemo() {
  const at = (date, time) => new Date(`${date}T${time}:00`).toISOString();
  const today = todayStr();
  const rows = [
    // [title, quad, cat, energy, dueOffset, time, dur, doneOffset|null, doneTime]
    ["Review slot provider contracts", "q2", "Work", "deep", -19, "10:00", 90, -19, "11:40"],
    ["Pay office internet invoice", "q1", "Finance", "admin", -17, "12:00", 15, -17, "12:10"],
    ["Prepare sprint demo", "q1", "Work", "light", -15, "14:00", 60, -14, "16:20"],
    ["Book dentist appointment", "q2", "Health", "admin", -12, "", 15, -12, "09:05"],
    ["Draft Q4 OKRs", "q2", "Work", "deep", -10, "09:30", 120, -9, "11:50"],
    ["Hand over CRM export to analytics", "q3", "Work", "admin", -8, "15:00", 30, -8, "15:25"],
    ["Read chapter 3 of Team Topologies", "q4", "Learning", "light", -6, "18:30", 45, -3, "19:20"],
    ["Fix banner copy for promo page", "q1", "Work", "light", -4, "11:30", 30, -4, "12:05"],
    ["Plan quarterly review deck", "q2", "Work", "deep", -2, "10:00", 90, -1, "10:40"],
    ["Send weekly KPI summary", "q1", "Work", "light", -1, "16:00", 30, null],
    ["1:1 notes for team leads", "q2", "Meetings", "light", 1, "16:30", 30, null],
    ["Draft Q1 roadmap priorities", "q2", "Work", "deep", 1, "09:00", 120, null],
    ["Renew car insurance", "q3", "Finance", "admin", 2, "", 30, null],
    ["Interview: senior PM candidate", "q1", "Meetings", "light", 2, "12:00", 60, null],
    ["Write test plan for live casino lobby", "q2", "Work", "deep", 3, "14:00", 180, null],
    ["Gym — legs day", "q2", "Health", "light", 4, "18:00", 60, null],
    ["Prepare budget forecast", "q2", "Finance", "deep", 2, "14:00", 150, null],
    ["Review design mocks", "q1", "Work", "light", 2, "16:00", 90, null],
  ];
  for (const [title, quad, cat, energy, dueOff, time, dur, doneOff, doneTime] of rows) {
    const date = dateShift(today, dueOff);
    const created = dateShift(date, -(2 + (title.length % 4)));
    let t = { id: newId(), title, quad, cat, energy, date, time, dur, remind: 15, repeat: "none", busy: true,
              dod: "", notes: "", done: doneOff != null, doneAt: doneOff != null ? at(dateShift(today, doneOff), doneTime) : null,
              createdAt: at(created, "09:00"), log: [{ at: at(created, "09:00"), a: "created" }] };
    if (title.startsWith("Draft Q4")) t.log.push({ at: at(dateShift(date, -1), "17:10"), a: "edited", c: [{ f: "dur", from: 60, to: 120 }, { f: "quad", from: "q1", to: "q2" }] });
    if (t.done) t.log.push({ at: t.doneAt, a: "completed" });
    const ev = { ...taskToEvent(t), id: "demo-" + t.id, webLink: "" };
    demoEvents.push(ev);
  }
}

async function demoGraph(method, path, body) {
  await new Promise((r) => setTimeout(r, 150));
  const url = new URL(path, GRAPH + "/");
  if (method === "GET" && url.pathname.endsWith("/me/calendarView")) {
    const from = ymd(new Date(url.searchParams.get("startDateTime")));
    const to = ymd(new Date(url.searchParams.get("endDateTime")));
    const value = [];
    for (let date = from; date < to; date = dateShift(date, 1)) {
      value.push(...demoDay(date), ...demoEvents.filter((e) => e.start.dateTime.startsWith(date)));
    }
    return { value };
  }
  if (method === "GET") return { value: demoEvents };
  const id = decodeURIComponent(url.pathname.split("/").pop());
  if (method === "POST") { const e = { ...body, id: "demo-" + newId(), webLink: "" }; demoEvents.push(e); return e; }
  const i = demoEvents.findIndex((e) => e.id === id);
  if (method === "PATCH" && i >= 0) { demoEvents[i] = { ...demoEvents[i], ...body }; return demoEvents[i]; }
  if (method === "DELETE" && i >= 0) { demoEvents.splice(i, 1); return null; }
  const err = new Error("Not found"); err.status = 404; throw err;
}

// ───────────────────────── Microsoft Graph ─────────────────────────
async function graph(method, path, body) {
  if (DEMO) return demoGraph(method, path, body);
  const token = await accessToken();
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      Prefer: `outlook.timezone="${TZ}"`,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const err = new Error((data && data.error && data.error.message) || res.statusText);
    err.status = res.status; err.code = data && data.error && data.error.code;
    throw err;
  }
  return data;
}

function metaOf(t) {
  return { v: 1, quad: t.quad, cat: t.cat, energy: t.energy, dod: t.dod, notes: t.notes,
           done: !!t.done, doneAt: t.doneAt || null, dur: t.dur, date: t.date, createdAt: t.createdAt || null,
           log: (t.log || []).slice(-LOG_CAP) };
}

// ── activity log ──
function withLog(t, entry) {
  return { ...t, log: [...(t.log || []), { at: new Date().toISOString(), ...entry }].slice(-LOG_CAP) };
}
function diffTask(a, b) {
  return Object.keys(FIELD_LABEL)
    .filter((f) => String(a[f] ?? "") !== String(b[f] ?? ""))
    .map((f) => (f === "notes" ? { f } : { f, from: a[f] ?? "", to: b[f] ?? "" }));
}
function fmtField(f, v) {
  if (v === "" || v == null) return f === "time" ? "All day" : "—";
  switch (f) {
    case "quad": return QUAD[v] || v;
    case "energy": return ENERGY[v] || v;
    case "date": return fmtDate(v);
    case "dur": return fmtDur(Number(v));
    case "remind": return Number(v) < 0 ? "None" : Number(v) === 0 ? "At start" : `${fmtDur(Number(v))} before`;
    case "repeat": return REPEAT[v] || "Doesn't repeat";
    case "busy": return v === true || v === "true" ? "Busy" : "Free";
    default: return String(v);
  }
}
function recordDeletion(t) {
  const list = ls.get("daybook.deletedLog", []);
  list.push({ at: new Date().toISOString(), a: "deleted", title: t.title, quad: t.quad, date: t.date });
  ls.set("daybook.deletedLog", list.slice(-300));
}

function recurrenceFor(t) {
  if (!t.repeat || t.repeat === "none" || t.repeat === "custom") return null;
  const d = startOf(t);
  const pattern = {
    daily: { type: "daily", interval: 1 },
    weekdays: { type: "weekly", interval: 1, daysOfWeek: DAYS.slice(1, 6), firstDayOfWeek: "monday" },
    weekly: { type: "weekly", interval: 1, daysOfWeek: [DAYS[d.getDay()]], firstDayOfWeek: "monday" },
    monthly: { type: "absoluteMonthly", interval: 1, dayOfMonth: d.getDate() },
  }[t.repeat];
  return { pattern, range: { type: "noEnd", startDate: t.date, recurrenceTimeZone: TZ } };
}

function repeatFrom(r) {
  if (!r || !r.pattern) return "none";
  const p = r.pattern;
  if (p.interval !== 1) return "custom";
  if (p.type === "daily") return "daily";
  if (p.type === "weekly" && p.daysOfWeek && p.daysOfWeek.length === 5 && !p.daysOfWeek.includes("saturday") && !p.daysOfWeek.includes("sunday")) return "weekdays";
  if (p.type === "weekly" && p.daysOfWeek && p.daysOfWeek.length === 1) return "weekly";
  if (p.type === "absoluteMonthly") return "monthly";
  return "custom";
}

function eventBody(t) {
  return [
    `Priority: ${QUAD[t.quad]}`,
    `Area: ${t.cat} · Energy: ${ENERGY[t.energy] || t.energy}`,
    t.dod ? `Done when: ${t.dod}` : "",
    t.notes ? `\n${t.notes}` : "",
    "\n— Added from Daybook",
  ].filter(Boolean).join("\n");
}

function taskToEvent(t, { includeRecurrence = true } = {}) {
  const ev = {
    subject: t.title,
    body: { contentType: "text", content: eventBody(t) },
    isAllDay: !t.time,
    showAs: t.busy ? "busy" : "free",
    isReminderOn: t.remind >= 0,
    categories: t.done ? [CATEGORY, DONE_CATEGORY] : [CATEGORY],
    singleValueExtendedProperties: [{ id: PROP_ID, value: JSON.stringify(metaOf(t)) }],
  };
  if (t.remind >= 0) ev.reminderMinutesBeforeStart = t.remind;
  const s = startOf(t);
  if (t.time) {
    const e = new Date(s.getTime() + t.dur * 60000);
    ev.start = { dateTime: localIso(s), timeZone: TZ };
    ev.end = { dateTime: localIso(e), timeZone: TZ };
  } else {
    const e = new Date(s); e.setDate(e.getDate() + 1);
    ev.start = { dateTime: ymd(s) + "T00:00:00", timeZone: TZ };
    ev.end = { dateTime: ymd(e) + "T00:00:00", timeZone: TZ };
  }
  if (includeRecurrence) ev.recurrence = recurrenceFor(t);
  return ev;
}

function parseGraphDate(dt) {
  const s = dt.dateTime.replace(/\.\d+$/, "");
  return dt.timeZone === "UTC" && TZ !== "UTC" ? new Date(s + "Z") : new Date(s);
}

function eventToTask(ev) {
  let meta = {};
  try { meta = JSON.parse((ev.singleValueExtendedProperties && ev.singleValueExtendedProperties[0] && ev.singleValueExtendedProperties[0].value) || "{}"); } catch {}
  let date, time = "", dur;
  if (ev.isAllDay) {
    date = ev.start.timeZone === "UTC" && meta.date ? meta.date : ev.start.dateTime.slice(0, 10);
    dur = meta.dur || 30;
  } else {
    const s = parseGraphDate(ev.start), e = parseGraphDate(ev.end);
    date = ymd(s); time = `${pad(s.getHours())}:${pad(s.getMinutes())}`;
    dur = Math.max(5, Math.round((e - s) / 60000));
  }
  return {
    id: ev.id, eventId: ev.id, webLink: ev.webLink || "",
    title: ev.subject || "(no title)",
    quad: QUAD[meta.quad] ? meta.quad : "q2",
    cat: meta.cat || "Work", energy: meta.energy || "light",
    date, time, dur,
    dod: meta.dod || "", notes: meta.notes || "",
    remind: ev.isReminderOn ? ev.reminderMinutesBeforeStart : -1,
    repeat: repeatFrom(ev.recurrence),
    busy: ev.showAs !== "free",
    done: !!meta.done || (ev.categories || []).includes(DONE_CATEGORY),
    doneAt: meta.doneAt || null, createdAt: meta.createdAt || null,
    log: Array.isArray(meta.log) ? meta.log : [],
  };
}

async function fetchOutlookTasks() {
  const filter = encodeURIComponent(`categories/any(c:c eq '${CATEGORY}')`);
  const expand = encodeURIComponent(`singleValueExtendedProperties($filter=id eq '${PROP_ID}')`);
  const select = "subject,start,end,isAllDay,recurrence,isReminderOn,reminderMinutesBeforeStart,showAs,webLink,categories";
  let url = `/me/events?$filter=${filter}&$expand=${expand}&$select=${select}&$top=100`;
  const out = [];
  while (url && out.length < 3000) {
    const page = await graph("GET", url);
    out.push(...page.value);
    url = page["@odata.nextLink"] || null;
  }
  return out.map(eventToTask);
}

// Outlook web "new event" link — used when not signed in.
function composeLink(t) {
  const base = settings.linkAcct === "personal" ? "https://outlook.live.com" : "https://outlook.office.com";
  const p = new URLSearchParams({ path: "/calendar/action/compose", rru: "addevent", subject: t.title, body: eventBody(t) });
  const s = startOf(t);
  if (t.time) {
    const e = new Date(s.getTime() + t.dur * 60000);
    p.set("startdt", s.toISOString()); p.set("enddt", e.toISOString()); p.set("allday", "false");
  } else {
    const e = new Date(s); e.setDate(e.getDate() + 1);
    p.set("startdt", ymd(s)); p.set("enddt", ymd(e)); p.set("allday", "true");
  }
  return `${base}/calendar/0/deeplink/compose?${p}`;
}

function describeError(e) {
  if (e instanceof TypeError || !navigator.onLine) return "You're offline. Try again when you're connected.";
  if (e.status === 401 || e.status === 403) return "Outlook refused access. Sign out and sign in again, or ask IT to approve the Calendars permission.";
  if (e.status === 429) return "Outlook is busy. Wait a minute and try again.";
  return `Outlook error: ${e.message || "unknown problem"}`;
}

// ───────────────────────── sync ─────────────────────────
async function refresh({ quiet = false } = {}) {
  if (!signedIn() || state.syncing) return;
  state.syncing = true; renderChrome();
  try {
    state.outlook = await fetchOutlookTasks();
    state.lastSync = Date.now();
    saveCache(); render();
    if (!quiet) toast("Up to date with Outlook");
  } catch (e) {
    toast(describeError(e), true);
  } finally {
    state.syncing = false; renderChrome();
  }
}

async function createTask(t) {
  if (!signedIn()) {
    state.local.push(t); saveLocal(); render();
    return { ok: true, local: true };
  }
  try {
    const ev = await graph("POST", "/me/events", taskToEvent(t));
    const saved = { ...t, id: ev.id, eventId: ev.id, webLink: ev.webLink || "" };
    state.outlook.push(saved); saveCache(); dayCache.clear(); render();
    return { ok: true };
  } catch (e) {
    if (e instanceof TypeError) {
      // Offline: keep it on the device and offer to upload later.
      state.local.push(t); saveLocal(); render(); renderChrome();
      return { ok: true, offline: true };
    }
    return { ok: false, error: describeError(e) };
  }
}

async function updateTask(orig, t) {
  if (!orig.eventId) {
    state.local = state.local.map((x) => (x.id === orig.id ? { ...t, id: orig.id } : x));
    saveLocal(); render();
    return { ok: true };
  }
  try {
    // A series' range start must follow its new date, so resend recurrence then too.
    const repeatChanged = orig.repeat !== t.repeat || (t.repeat !== "none" && t.repeat !== "custom" && orig.date !== t.date);
    await graph("PATCH", `/me/events/${encodeURIComponent(orig.eventId)}`, taskToEvent(t, { includeRecurrence: repeatChanged }));
    const saved = { ...t, id: orig.id, eventId: orig.eventId, webLink: orig.webLink };
    state.outlook = state.outlook.map((x) => (x.id === orig.id ? saved : x));
    saveCache(); dayCache.clear(); render();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: describeError(e) };
  }
}

async function setDone(t, done) {
  const next = withLog({ ...t, done, doneAt: done ? new Date().toISOString() : null }, { a: done ? "completed" : "reopened" });
  if (!t.eventId) {
    state.local = state.local.map((x) => (x.id === t.id ? next : x)); saveLocal(); render();
    return;
  }
  state.outlook = state.outlook.map((x) => (x.id === t.id ? next : x)); render();
  try {
    await graph("PATCH", `/me/events/${encodeURIComponent(t.eventId)}`, {
      categories: done ? [CATEGORY, DONE_CATEGORY] : [CATEGORY],
      singleValueExtendedProperties: [{ id: PROP_ID, value: JSON.stringify(metaOf(next)) }],
    });
    saveCache();
  } catch (e) {
    state.outlook = state.outlook.map((x) => (x.id === t.id ? t : x)); render();
    toast(describeError(e), true);
  }
}

async function deleteTask(t) {
  state.pendingDelete.delete(t.id);
  if (!t.eventId) {
    recordDeletion(t);
    state.local = state.local.filter((x) => x.id !== t.id); saveLocal(); render(); renderChrome();
    toast("Task deleted");
    return;
  }
  try {
    await graph("DELETE", `/me/events/${encodeURIComponent(t.eventId)}`);
    recordDeletion(t);
    state.outlook = state.outlook.filter((x) => x.id !== t.id); saveCache(); dayCache.clear(); render();
    toast("Deleted from Outlook");
  } catch (e) {
    if (e.status === 404) { state.outlook = state.outlook.filter((x) => x.id !== t.id); saveCache(); render(); return; }
    render(); toast(describeError(e), true);
  }
}

async function uploadLocal() {
  const btn = $("#banner-btn"); btn.disabled = true;
  let done = 0;
  for (const t of [...state.local]) {
    try {
      const ev = await graph("POST", "/me/events", taskToEvent(t));
      state.outlook.push({ ...t, id: ev.id, eventId: ev.id, webLink: ev.webLink || "" });
      state.local = state.local.filter((x) => x.id !== t.id);
      done++;
    } catch (e) { toast(describeError(e), true); break; }
  }
  saveLocal(); saveCache(); btn.disabled = false; render(); renderChrome();
  if (done) toast(`${done} task${done > 1 ? "s" : ""} added to Outlook`);
}

// ───────────────────────── rendering ─────────────────────────
const doneDay = (t) => (t.doneAt ? ymd(new Date(t.doneAt)) : t.date);
const allTasks = () => (signedIn() ? [...state.outlook, ...state.local] : state.local);

function bucket(t) {
  if (t.done) return "Done today";
  const today = todayStr();
  if (t.date < today && t.repeat === "none") return "Overdue";
  if (t.date <= today) return "Today";
  if (t.date === addDays(1)) return "Tomorrow";
  if (t.date <= addDays(7)) return "Next 7 days";
  return "Later";
}

function matches(t) {
  const f = state.filter;
  // Finished tasks stay on the list for the rest of the day, then live in the Archive.
  if (t.done && doneDay(t) !== todayStr()) return false;
  if (f.startsWith("q") && t.quad !== f) return false;
  if (state.query) {
    const q = state.query.toLowerCase();
    if (![t.title, t.notes, t.dod, t.cat].some((s) => (s || "").toLowerCase().includes(q))) return false;
  }
  return true;
}

function taskRow(t, example = false) {
  const late = !t.done && t.repeat === "none" && t.date < todayStr();
  const timeTxt = t.time ? `${t.time}–${endTime(t)}` : "All day";
  const syncLabel = signedIn()
    ? (t.eventId ? el("span", { class: "sync ok" }, "✓ In Outlook") : el("span", { class: "sync" }, "On this device"))
    : null;

  const acts = example ? el("div") : el("div", { class: "t-acts" },
    el("button", { type: "button", class: "btn btn-primary btn-sm", "aria-label": `Edit "${t.title}"`, onclick: () => startEdit(t) }, "Edit"),
    t.eventId && t.webLink
      ? el("a", { class: "btn btn-ghost btn-sm", href: t.webLink, target: "_blank", rel: "noopener" }, "Open in Outlook")
      : (!signedIn() ? el("a", { class: "btn btn-ghost btn-sm", href: composeLink(t), target: "_blank", rel: "noopener" }, "Add to Outlook →") : null),
    state.pendingDelete.has(t.id)
      ? el("button", { type: "button", class: "btn btn-danger btn-sm", onclick: () => deleteTask(t) },
          t.eventId ? (t.repeat !== "none" ? "Delete whole series" : "Delete from Outlook") : "Confirm delete")
      : el("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: () => {
          state.pendingDelete.add(t.id); render();
          setTimeout(() => { if (state.pendingDelete.delete(t.id)) render(); }, 4000);
        } }, "Delete")
  );

  return el("li", { class: `task ${t.quad}${t.done ? " done" : ""}${example ? " example" : ""}${state.editing && state.editing.id === t.id ? " editing" : ""}` },
    el("input", { type: "checkbox", class: "check", "aria-label": `Mark "${t.title}" done`, checked: !!t.done, disabled: example,
                  onchange: (e) => setDone(t, e.target.checked) }),
    el("div", { class: "t-main", role: example ? null : "button", tabindex: example ? null : "0", title: example ? null : "Edit task",
                onclick: () => !example && startEdit(t),
                onkeydown: (e) => { if (!example && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); startEdit(t); } } },
      el("div", { class: "t-title" }, example ? `Example — ${t.title}` : t.title),
      el("div", { class: "t-meta" },
        el("span", { class: "tag" }, QUAD[t.quad]),
        el("span", { class: "cat" }, t.cat),
        el("span", { class: late ? "late" : "" }, t.repeat !== "none" ? `${REPEAT[t.repeat]} from ${fmtDate(t.date)}` : fmtDate(t.date)),
        el("span", { class: "t-time" }, timeTxt),
        el("span", {}, ENERGY[t.energy] || ""),
        syncLabel
      ),
      t.dod ? el("div", { class: "t-dod" }, `Done when: ${t.dod}`) : null
    ),
    acts
  );
}

function examples() {
  return [
    { id: "ex1", title: "Send weekly KPI summary to the team", quad: "q1", cat: "Work", energy: "light", date: todayStr(), time: "10:00", dur: 30, dod: "Summary posted in the team channel", repeat: "none" },
    { id: "ex2", title: "Draft Q4 roadmap priorities", quad: "q2", cat: "Work", energy: "deep", date: addDays(1), time: "09:30", dur: 90, dod: "One-page draft shared for comments", repeat: "none" },
    { id: "ex3", title: "Book annual health check-up", quad: "q2", cat: "Health", energy: "admin", date: addDays(3), time: "", dur: 15, dod: "Appointment confirmed", repeat: "none" },
  ];
}

function render() {
  wl.at = 0; // task data may have changed: recompute workload on next look
  renderTasks();
  renderView();
}

function renderTasks() {
  const all = allTasks();
  const open = all.filter((t) => !t.done);
  const today = todayStr();
  $("#stats").textContent = `${open.length} open · ${open.filter((t) => bucket(t) === "Today").length} today · ${open.filter((t) => t.date < today && t.repeat === "none").length} overdue`;

  const list = $("#list"); list.replaceChildren();
  if (!all.length) {
    list.append(el("p", { class: "empty" }, signedIn()
      ? "No Daybook tasks in your Outlook calendar yet. Tap + or use the form to add one."
      : "No tasks yet — here's what a planned list looks like."));
    if (!signedIn()) list.append(el("div", { class: "group" }, el("ul", { class: "tasks" }, examples().map((t) => taskRow(t, true)))));
    return;
  }
  const shown = all.filter(matches);
  if (!shown.length) { list.append(el("p", { class: "empty" }, state.query ? "No tasks match your search." : "Nothing in this view.")); return; }

  const groups = {};
  for (const t of shown) (groups[bucket(t)] ||= []).push(t);
  for (const name of GROUPS) {
    const g = groups[name]; if (!g) continue;
    if (name === "Done today") g.sort((a, b) => (b.doneAt || b.date).localeCompare(a.doneAt || a.date));
    else g.sort((a, b) => a.date.localeCompare(b.date) || (a.time || "99").localeCompare(b.time || "99") || a.quad.localeCompare(b.quad));
    const mins = g.filter((t) => !t.done && t.time).reduce((s, t) => s + Number(t.dur), 0);
    list.append(el("div", { class: `group${name === "Overdue" ? " overdue" : ""}` },
      el("h3", {}, name, el("span", {}, `${g.length}${mins ? ` · ${Math.floor(mins / 60)}h ${pad(mins % 60)}m planned` : ""}`)),
      el("ul", { class: "tasks" }, g.map((t) => taskRow(t)))
    ));
  }
}

function renderChrome() {
  const st = $("#status");
  if (state.syncing) { st.dataset.kind = "busy"; st.textContent = "Syncing…"; }
  else if (signedIn()) { st.dataset.kind = "outlook"; st.textContent = DEMO ? "Demo calendar" : state.account.username || "Outlook connected"; }
  else { st.dataset.kind = "local"; st.textContent = "Local only"; }
  st.title = signedIn() ? state.account.username : "Tasks are saved on this device only";
  $("#btn-signin").hidden = signedIn();
  $("#btn-sync").hidden = !signedIn();
  $("#btn-sync").classList.toggle("spin", state.syncing);
  $("#save-hint").textContent = signedIn()
    ? "Saving adds the event to your Outlook calendar. No start time = all-day."
    : "Saved on this device. Sign in to send tasks to Outlook automatically.";

  const banner = $("#banner");
  if (signedIn() && state.local.length) {
    banner.hidden = false;
    $("#banner-text").textContent = `${state.local.length} task${state.local.length > 1 ? "s are" : " is"} saved only on this device.`;
    $("#banner-btn").textContent = "Add to Outlook";
  } else banner.hidden = true;
}

// ───────────────────────── day view & free-time suggestions ─────────────────────────
const work = {
  start: ls.get("daybook.workStart", "09:00"),
  end: ls.get("daybook.workEnd", "18:00"),
  weekends: ls.get("daybook.weekends", false),
};
const dayCache = new Map(); // "YYYY-MM-DD" → { at, events }
const PX_PER_MIN = 0.8;     // 48 px per hour
const toMin = (hhmm) => { const [h, m] = hhmm.split(":").map(Number); return h * 60 + m; };
const fmtMin = (m) => `${pad(Math.floor(m / 60) % 24)}:${pad(m % 60)}`;
const fmtDur = (m) => (m >= 60 ? `${Math.floor(m / 60)}h${m % 60 ? ` ${pad(m % 60)}m` : ""}` : `${m}m`);
const blocks = (ev) => !["free", "workingElsewhere"].includes(ev.showAs);
const dateShift = (date, n) => { const [y, m, d] = date.split("-").map(Number); return ymd(new Date(y, m - 1, d + n)); };
const isWeekend = (date) => { const [y, m, d] = date.split("-").map(Number); const w = new Date(y, m - 1, d).getDay(); return w === 0 || w === 6; };

async function dayEvents(date) {
  const c = dayCache.get(date);
  if (c && Date.now() - c.at < 120000) return c.events;
  const [y, m, d] = date.split("-").map(Number);
  const s = new Date(y, m - 1, d), e = new Date(y, m - 1, d + 1);
  let url = `/me/calendarView?startDateTime=${encodeURIComponent(s.toISOString())}&endDateTime=${encodeURIComponent(e.toISOString())}` +
    `&$select=subject,start,end,isAllDay,showAs,isCancelled,seriesMasterId,categories&$orderby=start/dateTime&$top=200`;
  const raw = [];
  while (url) { const page = await graph("GET", url); raw.push(...page.value); url = page["@odata.nextLink"] || null; }
  const events = raw.filter((ev) => !ev.isCancelled).map((ev) => {
    const base = { id: ev.id, master: ev.seriesMasterId || null, title: ev.subject || "Busy", showAs: ev.showAs,
                   daybook: (ev.categories || []).includes(CATEGORY) };
    if (ev.isAllDay) return { ...base, allDay: true };
    const st = parseGraphDate(ev.start), en = parseGraphDate(ev.end);
    return { ...base, allDay: false,
             start: Math.max(0, Math.round((st - s) / 60000)), end: Math.min(1440, Math.round((en - s) / 60000)) };
  });
  dayCache.set(date, { at: Date.now(), events });
  return events;
}

// Free intervals inside working hours (minutes from midnight).
function freeGaps(events, date) {
  let ws = toMin(work.start);
  const we = toMin(work.end);
  if (date === todayStr()) { const n = new Date(); ws = Math.max(ws, Math.ceil((n.getHours() * 60 + n.getMinutes()) / 15) * 15); }
  if (events.some((e) => e.allDay && e.showAs === "oof")) return [];
  const busy = events.filter((e) => !e.allDay && blocks(e)).map((e) => [e.start, e.end]).sort((a, b) => a[0] - b[0]);
  const gaps = [];
  let cur = ws;
  for (const [s, e] of busy) {
    if (e <= cur) continue;
    if (s >= we) break;
    if (s > cur) gaps.push([cur, Math.min(s, we)]);
    cur = Math.max(cur, e);
  }
  if (cur < we) gaps.push([cur, we]);
  return gaps;
}

// Up to n start times (minutes) that fit `dur`, nearest to `wanted` first.
function suggestStarts(gaps, dur, wanted, n = 3) {
  const out = new Set();
  for (const [gs, ge] of gaps) {
    const lo = Math.ceil(gs / 15) * 15, hi = Math.floor((ge - dur) / 15) * 15;
    if (hi < lo) continue;
    out.add(lo);
    if (wanted != null) out.add(Math.min(hi, Math.max(lo, Math.round(wanted / 15) * 15)));
  }
  const list = [...out];
  list.sort(wanted != null ? (a, b) => Math.abs(a - wanted) - Math.abs(b - wanted) || a - b : (a, b) => a - b);
  return list.slice(0, n).sort((a, b) => a - b);
}

async function nextFreeDay(fromDate, dur, wanted, exclude) {
  for (let i = 1; i <= 14; i++) {
    const date = dateShift(fromDate, i);
    if (!work.weekends && isWeekend(date)) continue;
    const evs = (await dayEvents(date)).filter(exclude);
    const starts = suggestStarts(freeGaps(evs, date), dur, wanted, 1);
    if (starts.length) return { date, start: starts[0] };
  }
  return null;
}

function useSlot(date, start) {
  form.date.value = date;
  form.time.value = fmtMin(start);
  updateDayView();
}

function slotButton(date, start, dur, withDate = false) {
  return el("button", { type: "button", class: "slot", onclick: () => useSlot(date, start) },
    `${withDate ? fmtDate(date) + ", " : ""}${fmtMin(start)}–${fmtMin(start + dur)}`);
}

let dvSeq = 0, dvTimer = 0;
function scheduleDayView() { clearTimeout(dvTimer); dvTimer = setTimeout(updateDayView, 200); }

async function updateDayView() {
  const dv = $("#dayview"), verdict = $("#dv-verdict"), sug = $("#dv-suggest"), note = $("#dv-note");
  const date = form.date.value;
  if (!date) { dv.hidden = true; return; }
  dv.hidden = false;
  $("#dv-title").textContent = `Your calendar · ${fmtDate(date)}`;
  verdict.hidden = true; sug.hidden = true; sug.replaceChildren(); note.textContent = "";

  if (!signedIn()) {
    $("#dv-load").textContent = "";
    $("#dv-allday").replaceChildren(); $("#dv-scroll").hidden = true;
    note.textContent = "Sign in with Microsoft to see this day's meetings here and get free-time suggestions.";
    return;
  }

  const seq = ++dvSeq;
  $("#dv-load").textContent = "Loading…";
  const eid = state.editing && state.editing.eventId;
  const exclude = (ev) => !eid || (ev.id !== eid && ev.master !== eid);
  let events;
  try { events = (await dayEvents(date)).filter(exclude); }
  catch (e) { if (seq === dvSeq) { $("#dv-load").textContent = ""; note.textContent = describeError(e); } return; }
  if (seq !== dvSeq) return;

  const dur = Number(form.dur.value);
  const wanted = form.time.value ? toMin(form.time.value) : null;
  const ws = toMin(work.start), we = toMin(work.end);
  const gaps = freeGaps(events, date);
  const freeMin = gaps.reduce((s, [a, b]) => s + (b - a), 0);
  const booked = events.filter((e) => !e.allDay && blocks(e))
    .reduce((s, e) => s + Math.max(0, Math.min(e.end, we) - Math.max(e.start, ws)), 0);
  $("#dv-load").textContent = `${fmtDur(booked)} booked · ${fmtDur(freeMin)} free`;

  const oof = events.find((e) => e.allDay && e.showAs === "oof");
  const clashes = wanted == null ? [] :
    events.filter((e) => !e.allDay && blocks(e) && e.start < wanted + dur && e.end > wanted);
  const starts = suggestStarts(gaps, dur, wanted, 3);

  if (wanted != null && (clashes.length || oof)) {
    verdict.className = "dv-verdict bad";
    verdict.replaceChildren(
      el("b", {}, "Busy at this time. "),
      oof ? `You're out of office (${oof.title}).`
          : `Clashes with ${clashes.map((c) => `"${c.title}" ${fmtMin(c.start)}–${fmtMin(c.end)}`).join(", ")}.`
    );
    verdict.hidden = false;
  } else if (wanted != null) {
    const outside = wanted < ws || wanted + dur > we;
    verdict.className = "dv-verdict ok";
    verdict.replaceChildren(el("b", {}, "You're free "), `${fmtMin(wanted)}–${fmtMin(wanted + dur)}.`,
      outside ? " It's outside your working hours." : "");
    verdict.hidden = false;
  }

  const needSuggestions = wanted == null || clashes.length || oof;
  if (needSuggestions && starts.length) {
    sug.append(el("span", {}, wanted == null ? `Free ${fmtDur(dur)} slots:` : "Try instead:"),
      ...starts.map((s) => slotButton(date, s, dur)));
    sug.hidden = false;
  } else if (needSuggestions) {
    verdict.className = "dv-verdict bad";
    verdict.replaceChildren(el("b", {}, "This day is full. "), `No free ${fmtDur(dur)} slot within ${work.start}–${work.end}.`);
    verdict.hidden = false;
    sug.append(el("span", {}, "Looking for the next free day…")); sug.hidden = false;
    nextFreeDay(date, dur, wanted, exclude).then((nf) => {
      if (seq !== dvSeq) return;
      sug.replaceChildren(nf ? el("span", {}, "Next free slot:") : el("span", {}, "No free slot in the next two weeks."),
        nf ? slotButton(nf.date, nf.start, dur, true) : null);
    }).catch(() => { if (seq === dvSeq) sug.hidden = true; });
  }

  drawDay(events, date, wanted, dur, clashes.length > 0 || !!oof);
}

function drawDay(events, date, wanted, dur, clash) {
  const allday = $("#dv-allday");
  allday.replaceChildren(...events.filter((e) => e.allDay).map((e) =>
    el("span", { class: e.showAs === "oof" ? "oof" : "", title: e.title }, `All day · ${e.title}`)));

  const timed = events.filter((e) => !e.allDay).sort((a, b) => a.start - b.start || b.end - a.end);
  const ws = toMin(work.start), we = toMin(work.end);
  let from = Math.min(ws, ...timed.map((e) => e.start), wanted ?? ws);
  let to = Math.max(we, ...timed.map((e) => e.end), wanted != null ? wanted + dur : we);
  from = Math.max(0, Math.floor(from / 60) * 60); to = Math.min(1440, Math.ceil(to / 60) * 60);
  const y = (m) => (m - from) * PX_PER_MIN;

  const grid = $("#dv-grid");
  grid.style.height = `${(to - from) * PX_PER_MIN}px`;
  const parts = [];
  if (ws > from) parts.push(el("div", { class: "dv-off", style: `top:0;height:${y(ws)}px` }));
  if (we < to) parts.push(el("div", { class: "dv-off", style: `top:${y(we)}px;height:${(to - we) * PX_PER_MIN}px` }));
  for (let h = from; h < to; h += 60) parts.push(el("div", { class: "dv-hour", style: `top:${y(h)}px` }, el("span", {}, fmtMin(h))));

  // Side-by-side lanes, counted per group of overlapping events.
  const placed = [];
  let cluster = [], laneEnds = [], clusterEnd = -1;
  const closeCluster = () => { for (const p of cluster) p.lanes = laneEnds.length; cluster = []; laneEnds = []; };
  for (const e of timed) {
    if (e.start >= clusterEnd) closeCluster();
    let lane = laneEnds.findIndex((end) => end <= e.start);
    if (lane === -1) { lane = laneEnds.length; laneEnds.push(e.end); } else laneEnds[lane] = e.end;
    const p = { e, lane, lanes: 1 };
    cluster.push(p); placed.push(p);
    clusterEnd = Math.max(clusterEnd, e.end);
  }
  closeCluster();
  // While a time is chosen, meetings use the left part and the new task the right column.
  const area = wanted != null ? "(100% - 44px) * 0.62" : "(100% - 44px)";
  for (const { e, lane, lanes } of placed) {
    const isClash = wanted != null && blocks(e) && e.start < wanted + dur && e.end > wanted;
    const cls = ["dv-ev", e.daybook ? "daybook" : "", e.showAs === "free" ? "free" : "", e.showAs === "tentative" ? "tentative" : "", isClash ? "clash" : ""].join(" ");
    const h = Math.max(16, (e.end - e.start) * PX_PER_MIN - 2);
    parts.push(el("div", {
      class: cls,
      style: `top:${y(e.start) + 1}px;height:${h}px;left:calc(40px + ${area} * ${lane / lanes});width:calc(${area} / ${lanes} - 2px)`,
      title: `${e.title} · ${fmtMin(e.start)}–${fmtMin(e.end)}`,
    }, e.title, h > 30 ? el("small", {}, `${fmtMin(e.start)}–${fmtMin(e.end)}`) : null));
  }
  if (wanted != null) {
    parts.push(el("div", { class: `dv-new${clash ? " clash" : ""}`, style: `top:${y(wanted)}px;height:${Math.max(16, dur * PX_PER_MIN)}px` },
      form.title.value.trim() || "New task"));
  }
  if (date === todayStr()) {
    const n = new Date(), nm = n.getHours() * 60 + n.getMinutes();
    if (nm > from && nm < to) parts.push(el("div", { class: "dv-now", style: `top:${y(nm)}px` }));
  }
  grid.replaceChildren(...parts);
  grid.dataset.from = from;

  const scroll = $("#dv-scroll");
  scroll.hidden = false;
  const focus = wanted ?? (date === todayStr() ? new Date().getHours() * 60 : ws);
  scroll.scrollTop = Math.max(0, y(focus) - 60);
}

// Tap an empty spot on the timeline to pick that start time.
$("#dv-grid").addEventListener("click", (e) => {
  const grid = e.currentTarget;
  const rect = grid.getBoundingClientRect();
  const m = Number(grid.dataset.from || 0) + (e.clientY - rect.top) / PX_PER_MIN;
  const start = Math.max(0, Math.min(1440 - 15, Math.floor(m / 15) * 15));
  useSlot(form.date.value, start);
});

// ───────────────────────── form ─────────────────────────
const form = $("#task-form");

function openForm() { document.body.classList.add("form-open"); setTimeout(() => $("#f-title").focus(), 50); }
function closeForm() { document.body.classList.remove("form-open"); }

function resetForm() {
  state.editing = null;
  form.reset();
  $("#f-date").value = todayStr();
  $("#form-h").textContent = "New task";
  $("#btn-save").textContent = "Save task";
  $("#btn-cancel-edit").hidden = true;
  $("#btn-clear").hidden = false;
  $("#form-err").hidden = true;
  const custom = $("#f-repeat").querySelector('option[value="custom"]');
  if (custom) custom.remove();
  $("#f-history").hidden = true;
  updateDayView();
  render();
}

function startEdit(t) {
  resetForm();
  state.editing = t;
  form.title.value = t.title;
  form.querySelector(`input[name="quad"][value="${t.quad}"]`).checked = true;
  form.cat.value = t.cat; form.energy.value = t.energy;
  form.date.value = t.date; form.time.value = t.time || "";
  if (![...form.dur.options].some((o) => o.value === String(t.dur))) form.dur.append(el("option", { value: String(t.dur) }, `${t.dur} min`));
  form.dur.value = String(t.dur);
  const remind = t.remind == null ? 15 : t.remind;
  if (![...form.remind.options].some((o) => o.value === String(remind))) form.remind.append(el("option", { value: String(remind) }, `${remind} min before`));
  form.remind.value = String(remind);
  if (t.repeat === "custom") form.repeat.append(el("option", { value: "custom" }, "Custom (set in Outlook)"));
  form.repeat.value = t.repeat || "none";
  form.dod.value = t.dod || ""; form.notes.value = t.notes || "";
  form.busy.checked = t.busy !== false;
  $("#form-h").textContent = "Edit task";
  $("#btn-save").textContent = signedIn() && t.eventId ? "Save changes to Outlook" : "Save changes";
  $("#btn-cancel-edit").hidden = false;
  $("#btn-clear").hidden = true;
  const hist = $("#f-history");
  hist.hidden = !(t.log && t.log.length);
  $("#f-history-list").replaceChildren(...historyItems(t.log || []));
  updateDayView();
  render();
  openForm();
  if (window.matchMedia("(min-width: 861px)").matches) $("#form-panel").scrollIntoView({ behavior: "smooth", block: "start" });
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const err = $("#form-err"); err.hidden = true;
  const title = form.title.value.trim();
  if (!title) { err.textContent = "Give the task a title — start with a verb."; err.hidden = false; form.title.focus(); return; }
  if (!form.date.value) { err.textContent = "Pick a date so the task lands in your calendar."; err.hidden = false; form.date.focus(); return; }

  const t = {
    id: state.editing ? state.editing.id : newId(),
    title, quad: form.quad.value, cat: form.cat.value, energy: form.energy.value,
    date: form.date.value, time: form.time.value || "", dur: Number(form.dur.value),
    remind: Number(form.remind.value), repeat: form.repeat.value, busy: form.busy.checked,
    dod: form.dod.value.trim(), notes: form.notes.value.trim(),
    done: state.editing ? !!state.editing.done : false,
    doneAt: state.editing ? state.editing.doneAt || null : null,
    createdAt: state.editing ? state.editing.createdAt || null : new Date().toISOString(),
    eventId: null,
    log: state.editing ? state.editing.log || [] : [],
  };
  let logged = t;
  if (state.editing) {
    const changes = diffTask(state.editing, t);
    if (changes.length) logged = withLog(t, { a: "edited", c: changes });
  } else {
    logged = withLog(t, { a: "created" });
  }

  const btn = $("#btn-save"); btn.disabled = true;
  const res = state.editing ? await updateTask(state.editing, logged) : await createTask(logged);
  btn.disabled = false;
  if (!res.ok) { err.textContent = res.error; err.hidden = false; return; }

  toast(state.editing
    ? "Changes saved"
    : res.offline ? "You're offline — saved on this device. Add it to Outlook when you're back online."
    : res.local ? "Saved on this device"
    : "Added to your Outlook calendar");
  const keepDate = t.date;
  resetForm(); form.date.value = keepDate; updateDayView();
  closeForm(); renderChrome();
});

form.addEventListener("reset", () => setTimeout(() => { if (!state.editing) $("#f-date").value = todayStr(); updateDayView(); }, 0));
for (const id of ["#f-date", "#f-time", "#f-dur"]) $(id).addEventListener("change", scheduleDayView);
$("#f-time").addEventListener("input", scheduleDayView);
$("#f-title").addEventListener("input", scheduleDayView);
$("#btn-cancel-edit").addEventListener("click", () => { resetForm(); closeForm(); });
$("#btn-close-form").addEventListener("click", () => { resetForm(); closeForm(); });
$("#fab").addEventListener("click", () => { resetForm(); openForm(); });

// ───────────────────────── settings ─────────────────────────
const dlg = $("#settings");
function openSettings(message) {
  $("#s-client").value = settings.clientId;
  $("#s-tenant").value = settings.tenant;
  $("#s-acct").value = settings.linkAcct;
  $("#s-work-start").value = work.start;
  $("#s-work-end").value = work.end;
  $("#s-weekends").checked = work.weekends;
  $("#s-redirect").textContent = REDIRECT_URI;
  $("#acct-line").textContent = signedIn() ? `Signed in as ${state.account.username}` : "Not signed in.";
  $("#btn-signout").hidden = !signedIn();
  const pw = $("#btn-change-pw");
  pw.hidden = !signedIn() || DEMO;
  if (signedIn()) pw.href = CHANGE_URL[accountKind(state.account.username)];
  const se = $("#settings-err"); se.hidden = !message; se.textContent = message || "";
  if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
}
function closeSettings() { if (dlg.open) dlg.close(); }

$("#btn-settings-save").addEventListener("click", () => {
  const clientId = $("#s-client").value.trim();
  const tenant = $("#s-tenant").value.trim() || "organizations";
  const se = $("#settings-err");
  if (clientId && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId)) {
    se.textContent = "The client ID should look like 00000000-0000-0000-0000-000000000000."; se.hidden = false; return;
  }
  const wsv = $("#s-work-start").value || "09:00", wev = $("#s-work-end").value || "18:00";
  if (toMin(wev) <= toMin(wsv)) { se.textContent = "Working hours must end after they start."; se.hidden = false; return; }
  work.start = wsv; work.end = wev; work.weekends = $("#s-weekends").checked;
  ls.set("daybook.workStart", work.start); ls.set("daybook.workEnd", work.end); ls.set("daybook.weekends", work.weekends);
  settings.linkAcct = $("#s-acct").value; ls.set("daybook.linkAcct", settings.linkAcct);
  const authChanged = clientId !== settings.clientId || tenant !== settings.tenant;
  ls.set("daybook.clientId", clientId); ls.set("daybook.tenant", tenant);
  if (authChanged) { location.reload(); return; }
  closeSettings(); render(); updateDayView(); toast("Settings saved");
});

// ───────────────────────── views: archive · activity · analytics ─────────────────────────
const ACTION = { created: "Created", edited: "Edited", completed: "Completed", reopened: "Restored", deleted: "Deleted" };
const fmtWhen = (iso) => new Date(iso).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
const fmtDateTime = (iso) => `${fmtDate(ymd(new Date(iso)))}, ${fmtWhen(iso)}`;
const onTime = (t) => doneDay(t) <= t.date;
const dayDiff = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000); // "YYYY-MM-DD" strings

function changeNodes(e) {
  if (e.a !== "edited" || !e.c) return null;
  return e.c.map((c) => el("span", { class: "log-change" }, `${FIELD_LABEL[c.f] || c.f}: `,
    c.f === "notes" ? "updated" : [el("s", {}, fmtField(c.f, c.from)), " → ", fmtField(c.f, c.to)]));
}
function historyItems(log) {
  return [...log].reverse().map((e) => el("li", {},
    el("span", { class: "log-when" }, fmtDateTime(e.at)),
    el("span", { class: "log-what" }, el("span", { class: `act ${e.a}` }, ACTION[e.a] || e.a), changeNodes(e))));
}

function setView(v) {
  if (!["tasks", "archive", "stats", "log"].includes(v)) v = "tasks";
  state.view = v;
  document.body.dataset.view = v;
  ls.set("daybook.view", v);
  document.querySelectorAll(".tabs [role=tab]").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.view === v)));
  for (const id of ["archive", "stats", "log"]) $(`#view-${id}`).hidden = v !== id;
  if (v !== "tasks") closeForm();
  renderView();
}
function renderView() {
  if (state.view === "archive") renderArchive();
  else if (state.view === "stats") renderStats();
  else if (state.view === "log") renderLog();
}

// ── archive ──
const archived = () => allTasks().filter((t) => t.done)
  .sort((a, b) => (b.doneAt || b.date).localeCompare(a.doneAt || a.date));

function renderArchive() {
  const all = archived();
  const q = state.archQuery.toLowerCase();
  const shown = q ? all.filter((t) => [t.title, t.notes, t.dod, t.cat].some((s) => (s || "").toLowerCase().includes(q))) : all;
  const ot = all.filter(onTime).length;
  $("#arch-stats").textContent = all.length ? `${all.length} done · ${Math.round((ot / all.length) * 100)}% on time` : "";
  $("#arch-export").disabled = !all.length;
  const list = $("#arch-list"); list.replaceChildren();
  if (!shown.length) {
    list.append(el("p", { class: "empty" }, all.length ? "No completed tasks match your search." : "Nothing here yet. Tick a task done and it appears here."));
    return;
  }
  const months = new Map();
  for (const t of shown) {
    const key = (t.doneAt ? new Date(t.doneAt) : startOf(t)).toLocaleDateString(undefined, { month: "long", year: "numeric" });
    if (!months.has(key)) months.set(key, []);
    months.get(key).push(t);
  }
  for (const [name, items] of months) {
    const mins = items.reduce((s, t) => s + Number(t.dur || 0), 0);
    list.append(el("div", { class: "month" },
      el("h3", {}, name, el("span", {}, `${items.length} done · ${fmtDur(mins)}`)), items.map(archRow)));
  }
}

function archRow(t) {
  const lateBy = onTime(t) ? 0 : dayDiff(t.date, doneDay(t));
  const open = state.openHistory.has(t.id);
  const confirming = state.pendingDelete.has(t.id);
  return el("div", { class: `arch ${t.quad}` },
    el("div", {},
      el("div", { class: "t-title" }, t.title),
      el("div", { class: "t-meta" },
        el("span", { class: "tag" }, QUAD[t.quad]),
        el("span", { class: "cat" }, t.cat),
        el("span", {}, `Due ${fmtDate(t.date)}`),
        el("span", {}, `Done ${t.doneAt ? fmtDateTime(t.doneAt) : "—"}`),
        el("span", { class: `badge ${lateBy ? "late" : "ontime"}` }, lateBy ? `${lateBy} day${lateBy > 1 ? "s" : ""} late` : "On time"))),
    el("div", { class: "t-acts" },
      el("button", { type: "button", class: "btn btn-ghost btn-sm",
        onclick: () => setDone(t, false).then(() => toast("Back on your task list")) }, "Restore"),
      el("button", { type: "button", class: "btn btn-ghost btn-sm", "aria-expanded": String(open),
        onclick: () => { if (open) state.openHistory.delete(t.id); else state.openHistory.add(t.id); renderArchive(); } },
        open ? "Hide history" : "History"),
      confirming
        ? el("button", { type: "button", class: "btn btn-danger btn-sm", onclick: () => deleteTask(t) }, t.eventId ? "Delete from Outlook" : "Confirm delete")
        : el("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: () => {
            state.pendingDelete.add(t.id); render();
            setTimeout(() => { if (state.pendingDelete.delete(t.id)) render(); }, 4000);
          } }, "Delete")),
    open ? el("ol", { class: "log-list" }, t.log && t.log.length ? historyItems(t.log)
      : el("li", {}, el("span", { class: "log-when" }, ""), el("span", { class: "log-what" }, "No history recorded for this task."))) : null);
}

function exportCsv() {
  const safe = (v) => { const s = String(v ?? ""); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  const rows = [["Title", "Priority", "Area", "Energy", "Due date", "Start", "Duration (min)", "Completed", "On time", "Done when", "Notes"]];
  for (const t of archived()) {
    rows.push([t.title, QUAD[t.quad], t.cat, ENERGY[t.energy] || t.energy, t.date, t.time || "All day", t.dur,
      t.doneAt ? `${doneDay(t)} ${fmtWhen(t.doneAt)}` : "", onTime(t) ? "Yes" : "No", t.dod || "", t.notes || ""]);
  }
  const csv = rows.map((r) => r.map((v) => `"${safe(v).replace(/"/g, '""')}"`).join(",")).join("\r\n");
  const url = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }));
  const a = el("a", { href: url, download: `daybook-archive-${todayStr()}.csv` });
  document.body.append(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
}

// ── activity ──
function activity() {
  const out = [];
  for (const t of allTasks()) for (const e of t.log || []) out.push({ ...e, title: t.title });
  for (const d of ls.get("daybook.deletedLog", [])) out.push(d);
  return out.sort((x, y) => y.at.localeCompare(x.at));
}

function renderLog() {
  const all = activity();
  const shown = state.logFilter === "all" ? all : all.filter((e) => e.a === state.logFilter);
  const weekAgo = dateShift(todayStr(), -6);
  $("#log-stats").textContent = `${all.filter((e) => ymd(new Date(e.at)) >= weekAgo).length} actions in the last 7 days`;
  const list = $("#log-list"); list.replaceChildren();
  if (!shown.length) { list.append(el("p", { class: "empty" }, "No activity yet.")); return; }
  const days = new Map();
  for (const e of shown.slice(0, 500)) {
    const k = ymd(new Date(e.at));
    if (!days.has(k)) days.set(k, []);
    days.get(k).push(e);
  }
  const today = todayStr();
  for (const [day, items] of days) {
    list.append(
      el("div", { class: "feed-day" }, day === today ? "Today" : day === dateShift(today, -1) ? "Yesterday" : fmtDate(day)),
      el("ol", { class: "feed log-list" }, items.map((e) => el("li", {},
        el("span", { class: "log-when" }, fmtWhen(e.at)),
        el("span", { class: "log-what" }, el("span", { class: `act ${e.a}` }, ACTION[e.a] || e.a), el("b", {}, e.title), changeNodes(e))))));
  }
}

// ── analytics ──
let wl = { at: 0, days: null, signed: null };

async function loadWorkload() {
  if (wl.days && Date.now() - wl.at < 120000 && wl.signed === signedIn()) return wl.days;
  const today = todayStr();
  const days = Array.from({ length: 14 }, (_, i) => {
    const date = dateShift(today, i);
    const weekend = isWeekend(date);
    return { date, weekend, meet: 0, task: 0, meetings: 0, tasks: 0, oof: false,
             cap: work.weekends || !weekend ? toMin(work.end) - toMin(work.start) : 0, iv: [] };
  });
  const byDate = new Map(days.map((d) => [d.date, d]));

  if (signedIn()) {
    const [y, m, d] = today.split("-").map(Number);
    const s = new Date(y, m - 1, d), e = new Date(y, m - 1, d + 14);
    let url = `/me/calendarView?startDateTime=${encodeURIComponent(s.toISOString())}&endDateTime=${encodeURIComponent(e.toISOString())}` +
      `&$select=subject,start,end,isAllDay,showAs,isCancelled,seriesMasterId,categories&$top=500`;
    const raw = [];
    while (url) { const page = await graph("GET", url); raw.push(...page.value); url = page["@odata.nextLink"] || null; }
    const tasksById = new Map(state.outlook.map((t) => [t.id, t]));
    for (const ev of raw) {
      if (ev.isCancelled) continue;
      const cats = ev.categories || [];
      const st = ev.isAllDay ? null : parseGraphDate(ev.start);
      const day = byDate.get(ev.isAllDay ? ev.start.dateTime.slice(0, 10) : ymd(st));
      if (!day) continue;
      if (cats.includes(CATEGORY)) {
        if (cats.includes(DONE_CATEGORY)) continue;
        const t = tasksById.get(ev.id) || tasksById.get(ev.seriesMasterId);
        day.task += ev.isAllDay ? Number((t && t.dur) || 30) : Math.round((parseGraphDate(ev.end) - st) / 60000);
        day.tasks++;
      } else if (ev.isAllDay) {
        if (ev.showAs === "oof") day.oof = true;
      } else if (blocks(ev)) {
        const startM = st.getHours() * 60 + st.getMinutes();
        day.iv.push([startM, Math.min(1440, startM + Math.round((parseGraphDate(ev.end) - st) / 60000))]);
        day.meetings++;
      }
    }
    for (const day of days) { // union of meeting intervals, so overlaps aren't counted twice
      let cur = -1;
      for (const [a, b] of day.iv.sort((p, q) => p[0] - q[0])) {
        if (b <= cur) continue;
        day.meet += b - Math.max(a, cur); cur = b;
      }
    }
  }
  for (const t of state.local) { // tasks not in Outlook yet
    const day = byDate.get(t.date);
    if (day && !t.done) { day.task += Number(t.dur); day.tasks++; }
  }
  wl = { at: Date.now(), days, signed: signedIn() };
  return days;
}

function kpi(label, value, sub, id) {
  return el("div", { class: "kpi", id }, el("span", { class: "k-label" }, label),
    el("span", { class: "k-value" }, String(value)), el("span", { class: "k-sub" }, sub));
}

function hbars(container, rows) {
  const max = Math.max(1, ...rows.map((r) => r.value));
  const total = rows.reduce((s, r) => s + r.value, 0);
  container.replaceChildren(...(total ? rows.map((r) => el("div", {
      class: `hbar ${r.cls || ""}`,
      "data-tip": `${r.name}\n${fmtDur(r.value)} · ${Math.round((r.value / total) * 100)}%${r.count != null ? ` · ${r.count} task${r.count === 1 ? "" : "s"}` : ""}`,
    },
    el("span", { class: "hb-name" }, r.name),
    el("div", { class: "hb-track" }, el("div", { class: "hb-fill", style: `width:${(r.value / max) * 100}%` })),
    el("span", { class: "hb-val" }, fmtDur(r.value))))
    : [el("p", { class: "empty" }, "No tasks in this period yet.")]));
}

function gridlines(chart, maxVal, H, step, fmt) {
  for (let v = 0; v <= maxVal + 1e-9; v += step) {
    chart.append(el("div", { class: "gridline", style: `bottom:${(v / maxVal) * H}px` }, el("span", {}, fmt(v))));
  }
}

let statsSeq = 0;
async function renderStats() {
  const seq = ++statsSeq;
  const P = state.period, today = todayStr();
  const since = dateShift(today, -(P - 1)), prevSince = dateShift(since, -P);
  const all = allTasks();
  const doneIn = (a, b) => all.filter((t) => t.done && doneDay(t) >= a && doneDay(t) <= b);
  const completed = doneIn(since, today);
  const prev = doneIn(prevSince, dateShift(since, -1));
  const onTimeN = completed.filter(onTime).length;
  const overdue = all.filter((t) => !t.done && t.repeat === "none" && t.date < today).sort((a, b) => a.date.localeCompare(b.date));
  const delta = completed.length - prev.length;

  $("#kpis").replaceChildren(
    kpi("Completed", completed.length, !prev.length && !completed.length ? `in the last ${P} days`
      : delta === 0 ? `same as the previous ${P} days` : `${delta > 0 ? "▲" : "▼"} ${Math.abs(delta)} vs the previous ${P} days`),
    kpi("On time", completed.length ? `${Math.round((onTimeN / completed.length) * 100)}%` : "—",
      completed.length ? `${onTimeN} of ${completed.length} by their due date` : "no completed tasks yet"),
    kpi("Overdue now", overdue.length, overdue.length ? `oldest due ${fmtDate(overdue[0].date)}` : "nothing overdue"),
    kpi("Next 7 days booked", "…", "loading calendar", "kpi-load"));

  // Added vs completed
  const weekly = P > 30;
  const buckets = [];
  if (!weekly) {
    for (let i = 0; i < P; i++) {
      const d = dateShift(since, i);
      const [y, m, dd] = d.split("-").map(Number);
      const label = P === 7 ? new Date(y, m - 1, dd).toLocaleDateString(undefined, { weekday: "short" }) : (i % 5 === 0 || i === P - 1 ? String(dd) : "");
      buckets.push({ from: d, to: d, label, tip: fmtDate(d) });
    }
  } else {
    for (let from = since, i = 0; from <= today; from = dateShift(from, 7), i++) {
      const to = dateShift(from, 6) > today ? today : dateShift(from, 6);
      buckets.push({ from, to, label: i % 2 === 0 ? fmtDate(from).replace(/^\S+\s/, "") : "", tip: `${fmtDate(from)} – ${fmtDate(to)}` });
    }
  }
  for (const b of buckets) {
    b.added = all.filter((t) => t.createdAt && ymd(new Date(t.createdAt)) >= b.from && ymd(new Date(t.createdAt)) <= b.to).length;
    b.done = all.filter((t) => t.done && doneDay(t) >= b.from && doneDay(t) <= b.to).length;
  }
  $("#flow-sub").textContent = weekly ? "per week" : "per day";
  const FH = 140, fmax = Math.max(2, Math.ceil(Math.max(...buckets.map((b) => Math.max(b.added, b.done))) / 2) * 2);
  const flow = el("div", { class: "vchart", style: `height:${FH}px` });
  gridlines(flow, fmax, FH, fmax / 2, (v) => String(v));
  for (const b of buckets) {
    flow.append(el("div", { class: "pair", "data-tip": `${b.tip}\nAdded: ${b.added}\nCompleted: ${b.done}` },
      el("i", { class: "added", style: `height:${(b.added / fmax) * FH}px` }),
      el("i", { class: "done", style: `height:${(b.done / fmax) * FH}px` })));
  }
  $("#flow-chart").replaceChildren(flow, el("div", { class: "xlabels" },
    buckets.map((b) => el("span", { class: b.to === today ? "today" : "" }, b.label))));

  // Planned time by priority / area (tasks due in the period), completed by energy
  const planned = all.filter((t) => t.date >= since && t.date <= today);
  const sumBy = (list, key, val) => list.filter((t) => t[key] === val);
  const quadRows = Object.keys(QUAD).map((q) => {
    const items = sumBy(planned, "quad", q);
    return { name: QUAD[q], cls: q, value: items.reduce((s, t) => s + Number(t.dur), 0), count: items.length };
  });
  hbars($("#quad-chart"), quadRows);
  const qTotal = quadRows.reduce((s, r) => s + r.value, 0);
  const share = (q) => (qTotal ? Math.round((quadRows.find((r) => r.cls === q).value / qTotal) * 100) : 0);
  $("#quad-insight").replaceChildren(...(qTotal ? [
    el("b", {}, `${share("q2")}%`), ` of planned time went to "Schedule" work (important, not urgent). `,
    share("q1") > 40 ? `"Do now" takes ${share("q1")}% — a sign of firefighting; block Schedule time earlier in the week.`
      : share("q2") >= 40 ? "That's a healthy balance." : "Aim for 40% or more to stay ahead of deadlines.",
  ] : []));
  const areas = [...new Set(planned.map((t) => t.cat))];
  hbars($("#area-chart"), areas.map((a) => {
    const items = sumBy(planned, "cat", a);
    return { name: a, value: items.reduce((s, t) => s + Number(t.dur), 0), count: items.length };
  }).sort((a, b) => b.value - a.value));
  hbars($("#energy-chart"), Object.keys(ENERGY).map((k) => {
    const items = sumBy(completed, "energy", k);
    return { name: ENERGY[k], value: items.reduce((s, t) => s + Number(t.dur), 0), count: items.length };
  }));

  // Workload (needs the calendar)
  $("#wl-sub").textContent = signedIn() ? (DEMO ? "demo calendar" : "meetings and tasks from Outlook") : "tasks only — sign in to include meetings";
  let days;
  try { days = await loadWorkload(); }
  catch (e) { if (seq === statsSeq) { $("#wl-chart").replaceChildren(el("p", { class: "empty" }, describeError(e))); } return; }
  if (seq !== statsSeq) return;
  renderWorkload(days);
}

function renderWorkload(days) {
  const today = todayStr();
  const H = 170;
  const maxBooked = Math.max(...days.map((d) => Math.max(d.meet + d.task, d.cap)));
  const maxH = Math.max(120, Math.ceil(maxBooked / 120) * 120); // whole 2-hour steps
  const chart = el("div", { class: "vchart", style: `height:${H}px` });
  gridlines(chart, maxH, H, maxH > 600 ? 240 : 120, (v) => `${v / 60}h`);
  for (const d of days) {
    const booked = d.meet + d.task;
    const pct = d.cap ? Math.round((booked / d.cap) * 100) : null;
    const over = d.cap ? booked > d.cap : booked > 0;
    const [y, m, dd] = d.date.split("-").map(Number);
    const tip = [fmtDate(d.date),
      `Meetings: ${fmtDur(d.meet)}${d.meetings ? ` (${d.meetings})` : ""}`,
      `Tasks: ${fmtDur(d.task)}${d.tasks ? ` (${d.tasks})` : ""}`,
      d.cap ? `${pct}% of ${fmtDur(d.cap)} working time` : "Day off",
      d.oof ? "Out of office" : ""].filter(Boolean).join("\n");
    const segs = [];
    if (d.meet) segs.push(el("i", { class: "meet", style: `height:${(d.meet / maxH) * H}px` }));
    if (d.task) segs.push(el("i", { class: "task", style: `height:${(d.task / maxH) * H}px` }));
    chart.append(el("div", { class: `col${d.weekend ? " weekend" : ""}`, "data-tip": tip },
      el("div", { class: "stack" }, segs),
      d.cap ? el("div", { class: "cap", style: `bottom:${(d.cap / maxH) * H}px` }) : null,
      booked && d.cap && (over || pct >= 85 || d.date === today)
        ? el("div", { class: `col-label${over ? " over" : ""}`, style: `bottom:${(booked / maxH) * H + 4}px` }, `${over ? "▲ " : ""}${pct}%`)
        : null));
    d.label = new Date(y, m - 1, dd).toLocaleDateString(undefined, { weekday: "narrow" }) + " " + dd;
  }
  $("#wl-chart").replaceChildren(chart, el("div", { class: "xlabels" },
    days.map((d) => el("span", { class: d.date === today ? "today" : "" }, d.label))));

  // KPI: next 7 days
  const week = days.slice(0, 7);
  const wBooked = week.reduce((s, d) => s + d.meet + d.task, 0), wCap = week.reduce((s, d) => s + d.cap, 0);
  const k = $("#kpi-load");
  if (k) k.replaceChildren(el("span", { class: "k-label" }, "Next 7 days booked"),
    el("span", { class: "k-value" }, wCap ? `${Math.round((wBooked / wCap) * 100)}%` : "—"),
    el("span", { class: "k-sub" }, `${fmtDur(wBooked)} of ${fmtDur(wCap)} working time`));

  // Insight
  const work14 = days.filter((d) => d.cap);
  const overDays = work14.filter((d) => d.meet + d.task > d.cap);
  const pctOf = (d) => Math.round(((d.meet + d.task) / d.cap) * 100);
  const ins = $("#wl-insight");
  if (!work14.length) ins.textContent = "";
  else if (overDays.length) {
    ins.replaceChildren(el("b", {}, `▲ ${overDays.length} overbooked day${overDays.length > 1 ? "s" : ""}: `),
      overDays.slice(0, 3).map((d) => `${fmtDate(d.date)} (${pctOf(d)}%)`).join(", "),
      ". Move Schedule or Later tasks to lighter days.");
  } else {
    const busiest = work14.reduce((a, b) => (pctOf(b) > pctOf(a) ? b : a));
    const lightest = work14.filter((d) => d.date !== today).reduce((a, b) => (pctOf(b) < pctOf(a) ? b : a), work14[work14.length - 1]);
    ins.replaceChildren("Busiest: ", el("b", {}, `${fmtDate(busiest.date)} — ${fmtDur(busiest.meet + busiest.task)} (${pctOf(busiest)}%)`),
      ". Lightest: ", el("b", {}, `${fmtDate(lightest.date)} (${pctOf(lightest)}%)`), " — a good day for deep-focus work.");
  }

  // Table view
  $("#wl-table").replaceChildren(el("table", {},
    el("thead", {}, el("tr", {}, ["Day", "Meetings", "Tasks", "Booked", "Working time", "Load"].map((h) => el("th", {}, h)))),
    el("tbody", {}, days.map((d) => el("tr", {},
      el("td", {}, fmtDate(d.date)), el("td", {}, fmtDur(d.meet)), el("td", {}, fmtDur(d.task)),
      el("td", {}, fmtDur(d.meet + d.task)), el("td", {}, d.cap ? fmtDur(d.cap) : "Day off"),
      el("td", {}, d.cap ? `${pctOf(d)}%` : "—"))))));
}

// ── chart tooltip ──
const tip = $("#tip");
function showTip(e) {
  const m = e.target.closest && e.target.closest("[data-tip]");
  if (!m) { tip.hidden = true; return; }
  if (tip.textContent !== m.dataset.tip) tip.textContent = m.dataset.tip;
  tip.hidden = false;
  const r = tip.getBoundingClientRect();
  let left = e.clientX + 12, top = e.clientY + 14;
  if (left + r.width > innerWidth - 8) left = e.clientX - r.width - 12;
  if (top + r.height > innerHeight - 8) top = e.clientY - r.height - 12;
  tip.style.left = `${Math.max(8, left)}px`; tip.style.top = `${Math.max(8, top)}px`;
}
document.addEventListener("pointermove", showTip);
document.addEventListener("pointerdown", showTip);
document.addEventListener("scroll", () => (tip.hidden = true), true);

// ───────────────────────── wiring ─────────────────────────
document.querySelectorAll(".tabs [role=tab]").forEach((b) => b.addEventListener("click", () => setView(b.dataset.view)));
document.querySelectorAll("#period-chips .chip").forEach((b) => b.addEventListener("click", () => {
  state.period = Number(b.dataset.p);
  document.querySelectorAll("#period-chips .chip").forEach((x) => x.setAttribute("aria-pressed", String(x === b)));
  renderStats();
}));
document.querySelectorAll("#log-chips .chip").forEach((b) => b.addEventListener("click", () => {
  state.logFilter = b.dataset.a;
  document.querySelectorAll("#log-chips .chip").forEach((x) => x.setAttribute("aria-pressed", String(x === b)));
  renderLog();
}));
$("#arch-search").addEventListener("input", (e) => { state.archQuery = e.target.value.trim(); renderArchive(); });
$("#arch-export").addEventListener("click", exportCsv);
$("#btn-signin").addEventListener("click", () => { ls.set("daybook.localMode", false); showGate(); });
$("#gate-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const email = $("#g-email").value.trim();
  const err = $("#gate-err");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    err.textContent = "Enter your email address, like name@company.com."; err.hidden = false; $("#g-email").focus(); return;
  }
  err.hidden = true;
  ls.set("daybook.email", email);
  const btn = $("#gate-submit"); btn.disabled = true; btn.textContent = "Opening Microsoft sign-in…";
  try { await signIn(email); }
  catch (ex) { err.textContent = ex.message; err.hidden = false; btn.disabled = false; btn.textContent = "Continue"; }
});
$("#gate-local").addEventListener("click", () => {
  ls.set("daybook.localMode", true);
  hideGate(); renderChrome();
  toast("Tasks are saved on this device only. Sign in any time from the top bar.");
});
$("#gate-settings").addEventListener("click", () => openSettings());

// Password reset happens at Microsoft: work/school accounts use self-service password reset,
// personal accounts (Outlook.com, Hotmail…) use the Microsoft account recovery page.
const RESET_URL = { work: "https://passwordreset.microsoftonline.com/", personal: "https://account.live.com/password/reset" };
const CHANGE_URL = { work: "https://mysignins.microsoft.com/security-info/password/change", personal: "https://account.live.com/password/Change" };
const accountKind = (email) => (PERSONAL_DOMAIN.test(email || "") ? "personal" : "work");
function updateResetLink() {
  const kind = accountKind($("#g-email").value.trim());
  $("#gate-reset-link").href = RESET_URL[kind];
  $("#gate-reset-note").hidden = kind === "personal";
}
$("#gate-forgot").addEventListener("click", () => {
  const panel = $("#gate-reset");
  panel.hidden = !panel.hidden;
  $("#gate-forgot").setAttribute("aria-expanded", String(!panel.hidden));
  updateResetLink();
});
$("#g-email").addEventListener("input", updateResetLink);
$("#btn-signout").addEventListener("click", () => signOut().catch((e) => toast(e.message, true)));
$("#btn-settings").addEventListener("click", () => openSettings());
$("#btn-sync").addEventListener("click", () => refresh());
$("#banner-btn").addEventListener("click", uploadLocal);
$("#search").addEventListener("input", (e) => { state.query = e.target.value.trim(); render(); });
document.querySelectorAll("#task-chips .chip").forEach((b) => b.addEventListener("click", () => {
  state.filter = b.dataset.f;
  document.querySelectorAll("#task-chips .chip").forEach((x) => x.setAttribute("aria-pressed", String(x === b)));
  render();
}));
const guide = $("#guide");
guide.open = ls.get("daybook.guide", true);
guide.addEventListener("toggle", () => ls.set("daybook.guide", guide.open));
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && Date.now() - state.lastSync > 60000) refresh({ quiet: true });
});
window.addEventListener("online", () => { renderChrome(); if (signedIn()) refresh({ quiet: true }); });

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}

// ───────────────────────── boot ─────────────────────────
(async () => {
  resetForm();
  renderChrome(); render();
  let authError = null;
  try { await initAuth(); }
  catch (e) {
    authError = e.errorCode === "user_cancelled" || /AADB2C90091|access_denied/.test(e.message || "")
      ? "Sign-in was cancelled. Enter your email to try again."
      : `Microsoft sign-in didn't work: ${e.errorMessage || e.message}`;
  }
  if (needsGate()) showGate(authError);
  else if (authError) toast(authError, true);
  if (location.hash === "#new") { history.replaceState(null, "", location.pathname); openForm(); }
  renderChrome(); render(); updateDayView();
  setView(ls.get("daybook.view", "tasks"));
  if (signedIn()) refresh({ quiet: true });
})();
})();
