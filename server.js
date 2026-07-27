/**
 * ============================================================
 *  Repair Slot Booking — standalone web app (no Google login)
 * ============================================================
 *  Talks to the SAME Google Sheet as the Apps Script version,
 *  but via a service account (Sheets API v4) instead of an
 *  Apps Script web app deployment. This means:
 *    - No "Anyone" execution restriction from Workspace admin.
 *    - Partner data (name/phone) never touches any personal
 *      Google account — it flows only through this server,
 *      using credentials your org controls (the service account).
 *
 *  Sheets used (same tabs/columns as Code.gs):
 *    - Auto            : col A = city list, col B = tool list
 *    - Main_Capacity   : city x date -> capacity/hour (helper)
 *    - Hourly_Capacity : city x date x time-slot -> availability
 *    - Main_bookings   : A Timestamp | B Name | C Phone | D City | E Tool | F Date | G Time Slot
 * ============================================================
 */

const express = require('express');
const path = require('path');
const { google } = require('googleapis');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEET_AUTO = 'Auto';
const SHEET_HOURLY = 'Hourly_Capacity';
const SHEET_BOOKINGS = 'Main_bookings';

const TIME_SLOTS = ['11 AM-12 PM', '12 PM-1 PM', '1 PM-2 PM', '2 PM-3 PM', '3 PM-4 PM', '4 PM-5 PM', '5 PM-6 PM'];
const SLOT_START_HOUR = {
  '11 AM-12 PM': 11, '12 PM-1 PM': 12, '1 PM-2 PM': 13,
  '2 PM-3 PM': 14, '3 PM-4 PM': 15, '4 PM-5 PM': 16, '5 PM-6 PM': 17
};

// ------------------------------------------------------------
// Google Sheets serial-date <-> yyyy-MM-dd helpers.
// Sheets epoch is Dec 30 1899 (a historical Lotus 1-2-3 quirk
// inherited by Excel/Sheets). Working in raw serials avoids any
// locale/date-format ambiguity from the Sheets API.
// ------------------------------------------------------------
const SHEETS_EPOCH_MS = Date.UTC(1899, 11, 30);

function serialToDateKey(serial) {
  const ms = SHEETS_EPOCH_MS + Math.round(Number(serial)) * 86400000;
  const d = new Date(ms);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function dateKeyToSerial(dateKey) {
  const [y, m, d] = dateKey.split('-').map(Number);
  const ms = Date.UTC(y, m - 1, d);
  return Math.round((ms - SHEETS_EPOCH_MS) / 86400000);
}

// ------------------------------------------------------------
// IST ("Asia/Kolkata", UTC+5:30) wall-clock helpers — used to
// decide "today" and "has this slot already started" regardless
// of the server's own system timezone.
// ------------------------------------------------------------
function nowIST() {
  return new Date(Date.now() + 5.5 * 60 * 60 * 1000);
}
function todayKeyIST() {
  const n = nowIST();
  const y = n.getUTCFullYear();
  const m = String(n.getUTCMonth() + 1).padStart(2, '0');
  const d = String(n.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
function currentHourIST() {
  return nowIST().getUTCHours();
}
function isSlotPast(dateKey, slot) {
  if (dateKey !== todayKeyIST()) return false;
  const startHour = SLOT_START_HOUR[slot];
  if (startHour === undefined) return false;
  return currentHourIST() >= startHour;
}

// ------------------------------------------------------------
// Google Sheets API client (service account JWT auth)
// ------------------------------------------------------------
let sheetsClientPromise = null;
let cachedSheetIds = null; // { 'Main_bookings': 12345, ... }

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

function getSheetsClient() {
  if (!sheetsClientPromise) {
    const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    if (keyJson) {
      // Explicit key (used for local testing / non-GCP hosts like Render).
      const credentials = JSON.parse(keyJson);
      const auth = new google.auth.JWT(credentials.client_email, undefined, credentials.private_key, SCOPES);
      sheetsClientPromise = auth.authorize().then(() => google.sheets({ version: 'v4', auth }));
    } else {
      // Application Default Credentials — used on Cloud Run when the service
      // is assigned the service account directly (no key file needed at all;
      // Google supplies credentials automatically at runtime).
      sheetsClientPromise = google.auth.getClient({ scopes: SCOPES })
        .then(auth => google.sheets({ version: 'v4', auth }));
    }
  }
  return sheetsClientPromise;
}

async function getSheetIdMap(sheets) {
  if (cachedSheetIds) return cachedSheetIds;
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties' });
  cachedSheetIds = {};
  meta.data.sheets.forEach(s => { cachedSheetIds[s.properties.title] = s.properties.sheetId; });
  return cachedSheetIds;
}

async function readRange(sheets, range) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE'
  });
  return res.data.values || [];
}

// ------------------------------------------------------------
// Simple in-process mutex — serializes booking writes so two
// simultaneous submissions can't both grab the last open slot.
// (Assumes a single server instance; fine for this scale.)
// ------------------------------------------------------------
let bookingChain = Promise.resolve();
function withBookingLock(fn) {
  const run = bookingChain.then(fn, fn);
  bookingChain = run.then(() => {}, () => {});
  return run;
}

// ------------------------------------------------------------
// GET /api/form-data -> { cities, tools }
// ------------------------------------------------------------
app.get('/api/form-data', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const rows = await readRange(sheets, `${SHEET_AUTO}!A2:B1000`);
    const cities = rows.map(r => r[0]).filter(v => v !== undefined && v !== null && String(v).trim() !== '');
    const tools = rows.map(r => r[1]).filter(v => v !== undefined && v !== null && String(v).trim() !== '');
    res.json({ cities, tools });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load form data' });
  }
});

// ------------------------------------------------------------
// GET /api/slots?city=CityName -> [{date, display, slots:[{slot,available}]}]
// ------------------------------------------------------------
app.get('/api/slots', async (req, res) => {
  try {
    const city = (req.query.city || '').toString().trim();
    if (!city) return res.json([]);

    const sheets = await getSheetsClient();
    const rows = await readRange(sheets, `${SHEET_HOURLY}!A2:G4000`); // City,Date,Slot,Total,Audits,Booked,Available

    const byDate = {};
    rows.forEach(row => {
      const [rowCity, rowDateSerial, slot, total, audits, booked, available] = row;
      if (rowCity !== city) return;
      if (rowDateSerial === undefined || rowDateSerial === null || rowDateSerial === '') return;
      if (!(Number(available) > 0)) return;

      const dateKey = serialToDateKey(rowDateSerial);
      if (isSlotPast(dateKey, slot)) return;

      if (!byDate[dateKey]) {
        const [y, m, d] = dateKey.split('-').map(Number);
        const display = new Date(Date.UTC(y, m - 1, d)).toLocaleDateString('en-GB', {
          day: '2-digit', month: 'short', weekday: 'short', timeZone: 'UTC'
        });
        byDate[dateKey] = { date: dateKey, display, slots: [] };
      }
      byDate[dateKey].slots.push({ slot, available: Math.floor(Number(available)) });
    });

    const result = Object.values(byDate);
    result.forEach(d => d.slots.sort((a, b) => TIME_SLOTS.indexOf(a.slot) - TIME_SLOTS.indexOf(b.slot)));
    result.sort((a, b) => a.date.localeCompare(b.date));
    res.json(result);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load slots' });
  }
});

// ------------------------------------------------------------
// Recompute live availability directly from source data
// (Hourly_Capacity's D/E plus a fresh count of Main_bookings),
// same approach as the Apps Script version — avoids trusting a
// formula's cached value at the exact moment of booking.
// ------------------------------------------------------------
async function computeLiveAvailability(sheets, city, dateKey, slot) {
  const hcRows = await readRange(sheets, `${SHEET_HOURLY}!A2:E4000`);
  let total = null, audits = 0;
  for (const row of hcRows) {
    const [rowCity, rowDateSerial, rowSlot, rowTotal, rowAudits] = row;
    if (rowCity === city && rowSlot === slot && rowDateSerial !== undefined &&
        serialToDateKey(rowDateSerial) === dateKey) {
      total = Number(rowTotal) || 0;
      audits = Number(rowAudits) || 0;
      break;
    }
  }
  if (total === null) return -1;

  const bookingRows = await readRange(sheets, `${SHEET_BOOKINGS}!D2:G20000`); // City,Tool,Date,Slot
  let bookedCount = 0;
  bookingRows.forEach(r => {
    const [bCity, bTool, bDateSerial, bSlot] = r;
    if (bCity === city && bSlot === slot && bDateSerial !== undefined &&
        serialToDateKey(bDateSerial) === dateKey) {
      bookedCount++;
    }
  });
  return total - audits - bookedCount;
}

// ------------------------------------------------------------
// POST /api/book { name, phone, city, tool, date, slot }
// ------------------------------------------------------------
app.post('/api/book', async (req, res) => {
  try {
    const result = await withBookingLock(() => handleBooking(req.body || {}));
    res.json(result);
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: 'Server error, please try again. / सर्वर में समस्या, कृपया फिर कोशिश करें।' });
  }
});

async function handleBooking(payload) {
  const name = (payload.name || '').toString().trim();
  const phone = (payload.phone || '').toString().trim();
  const partnerId = (payload.partnerId || '').toString().trim();
  const city = (payload.city || '').toString().trim();
  const tool = (payload.tool || '').toString().trim();
  const dateKey = (payload.date || '').toString().trim();
  const slot = (payload.slot || '').toString().trim();

  const phoneValid = /^\d{10}$/.test(phone);

  if (!name) return { success: false, message: 'Please enter name. / कृपया नाम भरें।' };
  if (!phoneValid && !partnerId) return { success: false, message: 'Please enter phone number or Partner ID. / कृपया फ़ोन नंबर या पार्टनर ID भरें।' };
  if (phone && !phoneValid) return { success: false, message: 'Please enter a valid 10-digit phone number. / कृपया सही 10 अंकों का फ़ोन नंबर भरें।' };
  if (!city || !tool || !dateKey || !slot) return { success: false, message: 'Please fill all fields. / कृपया सभी जानकारी भरें।' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return { success: false, message: 'Invalid date.' };

  if (isSlotPast(dateKey, slot)) {
    return { success: false, message: 'That slot has already passed for today. Please pick another. / यह स्लॉट आज के लिए निकल चुका है। कृपया दूसरा चुनें।', slotFilled: true };
  }

  const sheets = await getSheetsClient();
  const available = await computeLiveAvailability(sheets, city, dateKey, slot);
  if (available <= 0) {
    return { success: false, message: 'Sorry, that slot just got booked. Please pick another. / माफ़ करें, यह स्लॉट अभी भर गया। कृपया दूसरा चुनें।', slotFilled: true };
  }

  const dateSerial = dateKeyToSerial(dateKey);
  const timestamp = new Date(Date.now() + 5.5 * 60 * 60 * 1000).toISOString().replace('T', ' ').slice(0, 19) + ' IST';

  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_BOOKINGS}!A:H`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    // Partner ID is appended as column H (after Slot) rather than inserted
    // next to Phone, so the existing City/Tool/Date/Slot columns (D-G) that
    // Hourly_Capacity and Tomorrow_Summary formulas already point at don't shift.
    requestBody: { values: [[timestamp, name, phone, city, tool, dateSerial, slot, partnerId]] }
  });

  // cosmetic: format the newly-written Date cell as yyyy-mm-dd so it reads nicely in the sheet
  try {
    const updatedRange = appendRes.data.updates.updatedRange; // e.g. "Main_bookings!A15:G15"
    const rowMatch = updatedRange.match(/!A(\d+):/);
    if (rowMatch) {
      const rowNum = parseInt(rowMatch[1], 10);
      const sheetIds = await getSheetIdMap(sheets);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [{
            repeatCell: {
              range: { sheetId: sheetIds[SHEET_BOOKINGS], startRowIndex: rowNum - 1, endRowIndex: rowNum, startColumnIndex: 5, endColumnIndex: 6 },
              cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
              fields: 'userEnteredFormat.numberFormat'
            }
          }]
        }
      });
    }
  } catch (fmtErr) {
    console.warn('Could not format date cell (non-fatal):', fmtErr.message);
  }

  return {
    success: true,
    message: 'Booking confirmed! / बुकिंग पक्की हो गई!',
    details: { name, phone, partnerId, city, tool, date: dateKey, slot }
  };
}

app.get('/health', (req, res) => res.send('ok'));

app.listen(PORT, () => console.log(`Booking app listening on port ${PORT}`));
