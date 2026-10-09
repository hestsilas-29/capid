# CAP Cadet Check-In

A simple CAP attendance station with a keyboard barcode-reader mode, hosted as static files on Cloudflare Pages with a Google Apps Script backend.

## Features

- **Scanner page:** `index.html` for a USB/Bluetooth barcode scanner that types CAPIDs like a keyboard.
- **Google Sheets:** `Log` is always the first sheet and keeps the master history. Each new date automatically gets its own attendance sheet after `Log`.
- **No names:** attendance records contain only CAPID and time.
- **Offline queue:** scans are stored in the browser until the connection returns; the station code is not stored in the queue.
- **Email reports:** a scheduled Apps Script trigger emails staff once per day when there are check-ins.

## Google Sheets layout

The workbook is kept in this order:

1. **Log** — `Date | Time | CAPID`
2. **10-8-26** — the attendance for 10/8/26, with `CAPID | Time`
3. More dated attendance tabs as needed

Google Sheets does not accept `/` in a sheet-tab name, so the tab uses a dash format such as `10-8-26`. The dated sheet itself displays the exact date as `10/8/26` at the top.

Running `setup()` cleans the old Name columns from the app's Log/date tabs and removes the old `Roster` tab because names are no longer used by the app.

## Setup

### 1. Back end

1. Create or open the Google Sheet that will hold attendance.
2. **Extensions > Apps Script.** Replace the existing `Code.gs` with the provided version.
3. Check `TZ` near the top of `Code.gs` (default `America/New_York`).
4. Run **`setup`** once from the Apps Script editor and approve the permissions.
   - Creates/repairs the `Log` tab.
   - Moves `Log` to the first tab.
   - Removes the old Name column and old `Roster` tab used by previous versions.
   - Creates dated attendance tabs for any existing rows in `Log`.
   - Creates the scheduled email trigger.
5. **Deploy > New deployment > Web app.** Execute as **Me**. Who has access: **Anyone**.
6. Copy the Web App URL.

### 2. Front end / Cloudflare Pages

Put these files in the Cloudflare Pages project:

- `index.html`
- `config.html`
- `config.js`
- `app.css`
- `_headers`
Set the Apps Script Web App URL in `config.js` as `API_URL`.

### 3. First-time configuration

1. Open `.../config.html`.
2. Leave the current Admin PIN blank on first setup and click **Load settings**.
3. Enter staff email(s), the report time, a station code, and a new admin PIN.
4. Save settings.
5. Send a test email.

## Using the scanner page

Open the site root on the scanning computer. Enter the station code once for that page session, then scan CAP membership cards.

The page accepts the scanner's keyboard input and automatically submits once six digits are received. Pressing Enter also submits a scan. Duplicate scans of the same CAPID on the same day are reported as already checked in.

## Attendance behavior

- One check-in per CAPID per date.
- The master `Log` keeps every recorded check-in.
- A new dated attendance tab is created automatically for the date of the first scan.
- Existing `Log` rows can be backfilled into dated tabs by running `setup` again.
- Names are not collected, stored, displayed, or emailed by this version.
