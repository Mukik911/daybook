# Daybook — setup guide

Daybook is a Progressive Web App (PWA): one app that installs on **iPhone/iPad** and **Windows**.
It has no server of its own. Your device talks directly to Microsoft (sign-in + Outlook calendar
through Microsoft Graph). Each task is a normal Outlook event, tagged with the category **Daybook**.

Getting it running takes three steps:

1. **Host** the `daybook-app` folder on an HTTPS address (it's just static files).
2. **Register** the app in Microsoft Entra ID so it can write to your calendar.
3. **Install** it on your iPhone and Windows PC.

Until step 2 is done, the app still works in **Local only** mode: tasks are saved on the device and
each one gets an "Add to Outlook →" link.

---

## 1. Host the app (HTTPS)

Any static host works. You need HTTPS, because iOS and Windows install PWAs only from secure addresses.

**Option A — GitHub Pages (free, ~5 min)**
1. Create a new repository, for example `daybook`, and upload the *contents* of `daybook-app/`.
2. Go to **Settings → Pages → Build and deployment**, choose **Deploy from a branch**, branch `main`, folder `/ (root)`.
3. Your app address will be `https://<your-github-name>.github.io/daybook/`.

**Option B — Azure Static Web Apps (free tier, stays inside Microsoft)**
1. In the Azure portal, create a **Static Web App** (Free plan) and point it at the repository or upload the folder.
2. The address will look like `https://<name>.azurestaticapps.net/`.

> Whenever you deploy changed files, bump `VERSION` in `sw.js` (for example `daybook-v2`) so
> installed apps pick up the update.

## 2. Register the app in Microsoft Entra ID

You need an account that can create app registrations. At Betsson this may need IT, see the note below.

1. Open <https://entra.microsoft.com> → **Applications → App registrations → New registration**.
2. **Name:** `Daybook`.
3. **Supported account types:** *Accounts in this organizational directory only* (recommended for a work account).
4. **Redirect URI:** platform **Single-page application (SPA)**, value = your app address from step 1,
   with a trailing slash, for example `https://<your-github-name>.github.io/daybook/`.
   - Optional: also add `http://localhost:8080/` for local testing.
5. Click **Register**. On the **Overview** page, copy the:
   - **Application (client) ID**
   - **Directory (tenant) ID**
6. Go to **API permissions → Add a permission → Microsoft Graph → Delegated** and add:
   - `User.Read` (usually already there)
   - `Calendars.ReadWrite`
7. If your organization requires it, click **Grant admin consent** (or ask IT to).

Then put the IDs into the app in either of these ways:
- edit `config.js` (`clientId` and `tenant`) before you deploy, **or**
- open the app → ⚙ **Settings** → paste the **Application (client) ID** and the **tenant ID** → **Save**.

Daybook now opens on its **Sign in** screen:

1. Enter your work email and press **Continue**.
2. Microsoft's sign-in page opens with that email filled in. Enter your password there (and MFA if
   your organization uses it). Daybook never sees or stores your password.
3. You're back in Daybook, connected to Outlook. The top bar shows your email.

Daybook remembers your email and stays signed in on that device. **⚙ Settings → Sign out** brings
the sign-in screen back. **Use without an account** skips sign-in and keeps tasks on the device only;
press **Sign in** in the top bar to connect later.

**Forgot your password?** on the sign-in screen opens Microsoft's password reset: self-service
password reset (`passwordreset.microsoftonline.com`) for work accounts, account recovery
(`account.live.com`) for personal ones. When signed in, **⚙ Settings → Change password** opens
Microsoft's change-password page. Daybook itself never stores passwords, so there is nothing to
reset inside the app. For work accounts, self-service reset only works if your organization has
enabled it in Entra ID; otherwise the IT helpdesk resets it.

> A personal Outlook.com / Hotmail address needs Tenant set to `common` (step 2.3 must then allow
> personal accounts); the sign-in screen tells you if that's the case.

> **Note for IT / Betsson tenant.** The app asks only for *delegated* permissions
> (`User.Read`, `Calendars.ReadWrite`). It can act only on the signed-in user's own calendar,
> only while they use it. There is no backend and no client secret; tokens stay on the device.
> If user consent is turned off in the tenant, an admin must grant consent once.

## 3. Install on your devices

**iPhone / iPad (Safari)**
1. Open the app address in **Safari** (it must be Safari).
2. Tap **Share** → **Add to Home Screen** → **Add**.
3. Open Daybook from the home screen and sign in once.

**Windows (Microsoft Edge or Chrome)**
1. Open the app address in Edge.
2. Click the **App available / Install** icon in the address bar (or **⋯ → Apps → Install this site as an app**).
3. Daybook gets its own window, Start-menu entry and taskbar icon. Right-click the taskbar icon → **Pin**.

## How it works

| You do in Daybook | What happens in Outlook |
|---|---|
| Save a task with a start time | A timed event appears, with your reminder, Busy/Free and repeat settings |
| Save a task without a start time | An all-day event |
| Tick a task done | The event gets the category **Daybook: done** |
| Edit a task | The event is updated |
| Delete a task | The event is deleted (for a repeating task, the whole series) |
| Move or resize the event in Outlook | Daybook picks up the new time on the next refresh |

- **Your calendar inside the form.** Once a date is picked, the form shows that day's meetings,
  how much is booked vs. free, and where the new task would sit. If the time clashes, Daybook
  names the clashing meetings and offers the nearest free slots. If the day is full, it offers the
  next free day. Tap a slot, or tap an empty spot on the timeline, to use that time. Suggestions
  stay within your working hours (⚙ Settings → Working hours; weekends are skipped unless you turn them on).
  Events marked *Free* (such as lunch) don't count as busy; *Tentative* ones do.
- **Archive.** Completed tasks stay on the Tasks list for the rest of the day, then live under
  **Archive**: grouped by month, marked *On time* or *N days late*, searchable, with **Restore**,
  per-task **History** and **Export CSV** (opens in Excel).
- **Activity log.** Every create, edit (field by field: "Start: 16:00 → 17:00"), completion,
  restore and delete is recorded. Each task's history (last 30 entries) is saved on its Outlook event,
  so all devices see it. Deletions are logged on the device where they happened.
  See it all under **Activity**, or for one task under **History** in its edit form.
- **Analytics.** For the last 7, 30 or 90 days: tasks completed (vs. the previous period),
  on-time rate, overdue count, added-vs-completed trend, planned time by priority and area, and
  completed time by energy. **Workload · next 14 days** stacks meetings and Daybook tasks per day
  against your working hours, flags overbooked days and suggests the lightest day for deep work.
  Hover or tap any bar for exact numbers, or open **Show as table**.
- **Try it without signing in:** open the app address with `?demo` at the end
  (for example `http://localhost:8080/?demo`) to see these features with a made-up calendar.
- Priority, area, energy, "Done when…" and notes are stored on the event itself, so all your
  devices show the same list, with no extra database.
- Without a connection, tasks are saved on the device; the banner offers **Add to Outlook** once you're back online.
- Signing out in ⚙ Settings removes the app's session only; your Outlook stays signed in elsewhere.

## Run locally (for testing)

```bash
python -m http.server 8080 --directory daybook-app
```

Open <http://localhost:8080/>. To test sign-in locally, add `http://localhost:8080/` as an SPA redirect URI (step 2.4).

## Files

| File | Purpose |
|---|---|
| `index.html` | App layout |
| `styles.css` | Styling (light and dark, iPhone safe areas) |
| `app.js` | Tasks, Microsoft sign-in, Outlook sync |
| `config.js` | Client ID / tenant |
| `sw.js` | Offline support |
| `manifest.webmanifest`, `icons/` | Install metadata and icons |
| `vendor/msal-browser.min.js` | Microsoft Authentication Library v4.30.0 |
