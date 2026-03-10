# 📅 Appointment Booking System

A flexible, customizable appointment booking system built on Google Apps Script with a mobile-optimized UI.
The system features dual-language support (English/Hebrew), Google Calendar integration, automated email confirmations, and self-service appointment management (reschedule/cancel).

## ✨ Key Features

- **Dual-Language Interface** - Supports both English (LTR) and Hebrew (RTL) seamlessly
- **Dynamic Durations** - Base duration for one participant + additional time per extra participant (fully configurable)
- **Real-Time Availability** - Automatically checks multiple Google Calendars for conflicts
- **Automated Emails** - Confirmation & notification emails for both clients and business owners (includes ICS calendar files)
- **Self-Service Management** - Clients can securely reschedule or cancel via unique, tokenized links in their email
- **OTP Verification** - One-Time Password fallback via email in case the original link is lost
- **WhatsApp Integration** - Ready-to-use direct WhatsApp link generator for the client
- **Waze Navigation Integration** - Easily embed Waze links in emails and calendar events

---

## 🏗️ Technical Structure

| File | Description |
|-------|--------|
| `Code.gs` | The backend (Google Apps Script) - handles business logic, Calendar API, and email delivery |
| `index.html` | The main HTML template (served directly by the GAS Web App) |
| `script.html` | Client-side JavaScript - handles UI logic, validation, and date/time selection |
| `style.html` | Client-side CSS - responsive design supporting both RTL and LTR |
| `standalone.html` | An independent frontend file (HTML+CSS+JS bundled) - ideal for external hosting |
| `appsscript.json` | The GAS project manifest (timezones, required OAuth scopes) |

---

## 🚀 Installation Guide

### Step 1: Create the Google Apps Script Project

1. Go to [script.google.com](https://script.google.com) and click **"New project"**
2. Copy the contents of the local files into the new project:
   - `Code.gs` → Replace the default `Code.gs` content
   - `index.html` → File > New > HTML > Name: `index`
   - `script.html` → File > New > HTML > Name: `script`
   - `style.html` → File > New > HTML > Name: `style`

3. Update the `appsscript.json` manifest:
   - Click the ⚙️ (Project Settings) gear icon
   - Check the box for **"Show appsscript.json manifest file in editor"**
   - Go back to the editor, open `appsscript.json`, and replace its content with the local file's content

### Step 2: Configure `Code.gs` Variables

Open `Code.gs` and fill in your business details in the `CONFIG` object at the top of the file:

```javascript
const CONFIG = {
  // Operating hours - adjust to your business hours
  BUSINESS_HOURS: { ... },
  
  // Calendars - array of Calendar IDs to check for availability conflicts
  CALENDAR_IDS: ['primary'],
  
  // Buffer calendar - Calendar ID for adding mandatory buffer times around events (optional)
  SPECIAL_CALENDAR_ID: '',
  
  // Business Details
  BUSINESS_NAME: 'Hebrew Business Name',
  BUSINESS_NAME_EN: 'Business Name in English',
  BUSINESS_EMAIL: 'your-email@gmail.com',
  BUSINESS_EMAIL_FROM: 'your-email+meeting@gmail.com',
  BUSINESS_EMAIL_NAME: 'Your Business Name',
  BUSINESS_PHONE: '050-1234567',
  BUSINESS_PHONE_INTL: '972501234567',        // International format for WhatsApp links
  
  // Location
  BUSINESS_ADDRESS_HE: 'Street 1, City',
  BUSINESS_ADDRESS_EN: 'Street 1, City',
  WAZE_LINK: 'https://waze.com/ul/...',        // Waze sharing link
  
  // Event Titles
  EVENT_TITLE_HE: 'פגישה',
  EVENT_TITLE_EN: 'Appointment',
  
  // WhatsApp Pre-filled Messages
  WHATSAPP_MSG_EN: 'Hi, I booked an appointment for',
  WHATSAPP_MSG_HE: 'היי, קבעתי פגישה ל',
  
  // Domain for generated ICS files
  ICS_DOMAIN: 'yourdomain.com',
  
  // Frontend App URL (Where you will host standalone.html)
  MANAGE_BASE_URL: 'https://your-username.github.io/meeting.html'
};
```

> **💡 Tip:** Every customization variable you need to adjust is contained within the `CONFIG` block. You do not need to modify the core logic.

### Step 3: Deploy the Web App

1. In the Apps Script project, click **"Deploy" > "New deployment"**
2. Select type: **"Web app"**
3. Configure settings:
   - **Execute as:** "Me"
   - **Who has access:** "Anyone"
4. Click **"Deploy"** and copy the generated Web App URL.

> **⚠️ Important:** Every time you modify the backend code, you must create a **New deployment** for the changes to take effect.

### Step 4: Setup the Frontend (standalone.html) - Recommended

While the Web App works out of the box, it is highly recommended to host `standalone.html` externally (like GitHub Pages). This avoids Google iframe security restrictions, third-party cookie blocking, and provides a cleaner URL.

#### Link the Frontend to the Backend

Open `standalone.html` and locate the API connection constant near the top of the JavaScript block:

```javascript
const API_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
```

Replace the placeholder with the Web App URL you copied in Step 3.

```javascript
const API_URL = 'https://script.google.com/macros/s/xxxxx.../exec';
```

> **💡 Note:** Ensure you also update `YOUR_BUSINESS_NAME_HE` in `standalone.html` to act as a fallback title in case the backend server cannot be reached.

#### Free Hosting via GitHub Pages

If you do not have web hosting, you can host the frontend for free using GitHub Pages:

1. Create a **Public** repository named `your-username.github.io`
2. Upload `standalone.html` to this repo and rename it to `meeting.html` (or any preferred name).
3. In Repo Settings > **Pages**, set the Source branch to `main` and Save.
4. Your application will be live at: `https://your-username.github.io/meeting.html`
5. **CRITICAL:** Go back to `Code.gs` and update `MANAGE_BASE_URL` with this new URL so the management emails point to the correct place.

---

## 🔧 Two Deployment Options

### Option A: Complete Google Apps Script (Simple)
- Host everything as a single Google Web App script.
- **Pros:** Fast setup, no external hosting required.
- **Cons:** Runs inside a Google iframe, which can cause browser security warnings (Safari/iOS) or block third-party cookies needed for tracking.

### Option B: Decoupled Architecture (Recommended)
- Host `standalone.html` on your own domain (or GitHub Pages) while interacting with the Apps Script Web App strictly via API calls.
- **Pros:** Full control over the domain, no iframe warnings, optimal mobile performance.
- **Cons:** Requires a minor external hosting setup (Step 4).

---

## 🌐 Language Mechanics

- The client selects their preferred language (English/Hebrew) via a toggle on the first UI step.
- The UI translates dynamically without reloading.
- Confirmation emails to the client are localized in the chosen language.
- Calendar events are created in the chosen language.
- Notification emails to the business owner are consistently sent in Hebrew.

---

## ❓ Troubleshooting

| Issue | Solution |
|-------|----------|
| Code changes aren't updating | You must create a **New Deployment** in Apps Script after every change. |
| Emails aren't sending | Ensure `BUSINESS_EMAIL_FROM` matches the executing Google account (or is a verified alias). |
| Standalone frontend won't connect | Check that `API_URL` is correct, and verify the Web App deployment is set to **"Anyone"** access. |
| Page loads but buttons don't work (unclickable) | The client likely has a strict internet filter (like "Netspark" or "Rimon" in Israel) blocking background scripts. Have them whitelist `script.google.com` and `script.googleusercontent.com`. |
