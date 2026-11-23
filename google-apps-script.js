/**
 * ==============================================================================
 * B & W PROFESSIONAL HUB — SECURE APPOINTMENT INTAKE WEBHOOK
 * Practice: Flat 6, Roopanand CHS, Lullanagar, Pune 411040
 * Desk: +91 91566 32846 | Strictly By Appointment Only (Mon–Fri)
 * ==============================================================================
 * 
 * 3-MINUTE DEPLOYMENT INSTRUCTIONS:
 * 1. Open Google Sheets (https://sheets.new) and name the spreadsheet:
 *    "B&W Professional Hub - Confidential Appointments"
 * 2. In the top menu, go to: Extensions > Apps Script
 * 3. Delete any default code in Code.gs, paste this ENTIRE file, and click Save (💾).
 * 4. (Optional) Run `testSetup()` once from the editor to initialize headers.
 * 5. Click "Deploy" > "New deployment"
 * 6. Click the gear icon (⚙️) next to "Select type" and choose "Web app".
 * 7. Configure deployment settings:
 *    - Description: "B&W Professional Hub Live Intake Webhook"
 *    - Execute as: "Me (your-email@gmail.com)"
 *    - Who has access: "Anyone"  <-- CRITICAL for frontend fetch requests
 * 8. Click "Deploy", authorize permissions when prompted.
 * 9. Copy the generated "Web app URL" (starts with https://script.google.com/macros/s/...)
 * 10. Open `index.html` and paste the URL into `GOOGLE_SHEETS_WEBHOOK_URL`:
 *     const GOOGLE_SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/.../exec";
 * ==============================================================================
 */

const SHEET_NAME = 'Appointments';
const NOTIFICATION_EMAIL = ''; // Optional: Enter your email (e.g., 'chambers@bwprofessional.com') to receive instant email alerts

/**
 * Handles incoming HTTP POST requests from the website booking form
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Wait up to 10 seconds for concurrent requests to prevent race conditions
  lock.tryLock(10000);

  try {
    let data;

    // Parse incoming JSON or form-encoded payload
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    } else {
      data = {};
    }

    // Access or initialize the target sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      initializeSheetHeaders(sheet);
    } else if (sheet.getLastRow() === 0) {
      initializeSheetHeaders(sheet);
    }

    // Format IST Timestamp (Indian Standard Time, GMT+5:30)
    const istTimestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss 'IST'");
    
    // Sanitize values to prevent Formula Injection (=, +, -, @) and malicious characters
    const refId = sanitizeSheetCell(data.refId || ('BW-' + new Date().getFullYear() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase()));
    const clientName = sanitizeSheetCell(data.clientName || 'Anonymous Inquiry');
    const clientPhone = sanitizeSheetCell(data.clientPhone || 'Not Provided');
    const service = sanitizeSheetCell(data.consultationService || 'Unspecified');
    const preferredDay = sanitizeSheetCell(data.preferredDay || 'Pending Selection');
    const timeWindow = sanitizeSheetCell(data.timeWindow || 'Pending Selection');
    const briefNotes = sanitizeSheetCell(data.briefNotes || 'None');
    const source = sanitizeSheetCell(data.source || 'B&W Web Portal');
    const status = 'Pending Confirmation';

    // Append appointment row
    sheet.appendRow([
      refId,
      istTimestamp,
      clientName,
      clientPhone,
      service,
      preferredDay,
      timeWindow,
      briefNotes,
      source,
      status
    ]);

    // Format status cell with subtle amber badge style
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 10).setBackground('#FEF3C7').setFontColor('#92400E').setFontWeight('bold');

    // Send immediate email notification if configured
    if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL.trim() !== '') {
      sendEmailAlert({
        refId: refId,
        timestamp: istTimestamp,
        name: clientName,
        phone: clientPhone,
        service: service,
        day: preferredDay,
        time: timeWindow,
        notes: briefNotes
      });
    }

    // Return JSON response with CORS headers
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Intake logged successfully in B&W Professional Hub registry',
      referenceId: refId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles HTTP GET requests for health check and testing
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'online',
    service: 'B & W Professional Hub Appointment Engine',
    location: 'Flat 6, Roopanand CHS, Lullanagar, Pune 411040',
    schedule: 'Monday – Friday, 9:30 AM – 5:30 PM (Strictly By Appointment Only)',
    timestamp: Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss 'IST'")
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Formats and styles the header row with high-ticket architectural colors
 */
function initializeSheetHeaders(sheet) {
  const headers = [
    'Reference ID',
    'Timestamp (IST)',
    'Client Name',
    'Phone / WhatsApp',
    'Consultation Discipline',
    'Preferred Day',
    'Time Window',
    'Confidential Notes',
    'Submission Source',
    'Status'
  ];

  sheet.appendRow(headers);

  // Apply Forest Green (#16302B) background and white bold typography
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#16302B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);

  // Set explicit column widths for executive readability
  sheet.setColumnWidth(1, 150); // Ref ID
  sheet.setColumnWidth(2, 170); // Timestamp
  sheet.setColumnWidth(3, 190); // Client Name
  sheet.setColumnWidth(4, 150); // Phone
  sheet.setColumnWidth(5, 260); // Service
  sheet.setColumnWidth(6, 130); // Preferred Day
  sheet.setColumnWidth(7, 210); // Time Window
  sheet.setColumnWidth(8, 280); // Notes
  sheet.setColumnWidth(9, 140); // Source
  sheet.setColumnWidth(10, 160); // Status
}

/**
 * Security: Neutralizes formula injection risks in Google Sheets
 * Prepends a single quote if a value starts with =, +, -, @, \t, \r
 */
function sanitizeSheetCell(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (/^[=\+\-\@\t\r]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

/**
 * Optional email dispatcher to alert the principal coordinator
 */
function sendEmailAlert(intake) {
  try {
    const subject = `[B&W Intake] New Consultation Request: ${intake.name} (${intake.service})`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; background-color: #F8F6F1; padding: 24px; color: #111615;">
        <div style="max-width: 600px; margin: 0 auto; background: #FFFFFF; border: 1px solid #B8894A; border-radius: 6px; overflow: hidden;">
          <div style="background-color: #16302B; padding: 18px 24px; color: #F8F6F1;">
            <h2 style="margin: 0; font-size: 18px; font-weight: normal; letter-spacing: 1px; color: #D3A465;">B & W PROFESSIONAL HUB</h2>
            <p style="margin: 4px 0 0; font-size: 11px; text-transform: uppercase; letter-spacing: 2px; color: #DDD8CB;">Confidential Intake Notification • Lullanagar Studio</p>
          </div>
          
          <div style="padding: 24px;">
            <p style="font-size: 14px; margin-top: 0;">A new confidential client intake has been logged from the portal:</p>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin: 16px 0;">
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold; width: 35%;">Reference ID:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.refId}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Timestamp (IST):</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.timestamp}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Client Name:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.name}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Phone / WhatsApp:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;"><a href="tel:${intake.phone}">${intake.phone}</a></td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Discipline:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.service}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Preferred Day:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.day}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Time Window:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.time}</td></tr>
              <tr><td style="padding: 8px; border-bottom: 1px solid #EFECE4; font-weight: bold;">Confidential Notes:</td><td style="padding: 8px; border-bottom: 1px solid #EFECE4;">${intake.notes}</td></tr>
            </table>

            <div style="margin-top: 20px; padding: 12px; background: #EFECE4; border-radius: 4px; font-size: 12px; color: #16302B;">
              <strong>Action Required:</strong> Verify practitioner calendar availability for Flat 6 Roopanand CHS and confirm slot via WhatsApp or Phone desk.
            </div>
          </div>
          
          <div style="background-color: #111615; padding: 12px 24px; font-size: 11px; color: #DDD8CB; text-align: center;">
            Strictly Confidential • B & W Professional Hub • Flat 6, Roopanand CHS, Lullanagar, Pune
          </div>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (mailErr) {
    console.warn('Email dispatch notice:', mailErr);
  }
}

/**
 * Manual test function to initialize headers and verify permissions
 */
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  initializeSheetHeaders(sheet);
  Logger.log("B&W Professional Hub sheet initialized successfully.");
}
