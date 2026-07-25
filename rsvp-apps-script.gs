/**
 * WEDDING RSVP BACKEND — Google Apps Script
 * Receives RSVP submissions from the invitation page and serves
 * the blessings back for the Blessings Wall.
 *
 * SETUP (5 minutes):
 * 1. Go to sheets.google.com → create a new blank spreadsheet, name it "Wedding RSVPs"
 * 2. In the sheet: Extensions → Apps Script
 * 3. Delete any code there, paste THIS ENTIRE FILE, click Save (disk icon)
 * 4. Click "Deploy" → "New deployment" → gear icon → "Web app"
 *      - Description: rsvp
 *      - Execute as: Me
 *      - Who has access: Anyone          <-- important, guests are not logged in
 *    Click Deploy, approve the permissions (it only touches this sheet)
 * 5. Copy the "Web app URL" (ends with /exec) and paste it into the
 *    invitation HTML where it says RSVP_URL (or send it to Claude to wire in)
 *
 * MODERATION: every response is a row in your sheet. Delete a row and it
 * disappears from the Blessings Wall on the next page load.
 */

const SHEET_NAME = 'RSVPs';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Name', 'Guests', 'Attending', 'Side', 'Message']);
  }
  return sh;
}

/* Receives a new RSVP (called by the invitation page) */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || '').slice(0, 80).trim();
    if (!name) throw new Error('name required');
    getSheet_().appendRow([
      new Date(),
      name,
      Math.max(1, Math.min(12, parseInt(d.guests) || 1)),
      d.attending === 'Not attending' ? 'Not attending' : 'Attending',
      ['groom', 'bride', 'well'].includes(d.side) ? d.side : 'well',
      String(d.message || '').slice(0, 500).trim()
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* Returns all blessings (called by the page to fill the wall) */
function doGet() {
  const rows = getSheet_().getDataRange().getValues().slice(1); // skip header
  const out = rows
    .filter(r => r[1])
    .map(r => ({
      name: r[1],
      guests: r[2],
      attending: r[3],
      side: r[4],
      message: r[5]
    }))
    .reverse(); // newest first
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
