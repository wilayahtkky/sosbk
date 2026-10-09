const scriptProp = PropertiesService.getScriptProperties();

function initialSetup() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', doc.getId());
}

function sanitizeValue(value) {
  if (typeof value !== 'string') return value;
  const formulaTriggers = ['=', '+', '-', '@'];
  if (formulaTriggers.some(trigger => value.startsWith(trigger))) {
    return "'" + value;
  }
  return value;
}

function sendAdminNotification(nextRow) {
  const recipient = 'wilayahtkky@gmail.com';
  const senderName = 'wil TKKY - Buntoro';
  const subject = 'New Registration: Sosialisasi Bulan Keluarga 2026';
  const body = 'A new submission has been added to row ' + nextRow;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: body,
    name: senderName,
  });
}

function sendUserConfirmationEmail(userEmail, data) {
  const senderName = 'Panitia Bulan Keluarga Katedral Jakarta';
  const subject = 'Bukti Pendaftaran Sosialisasi Bulan Keluarga 2026';

  let peserta2Info = '';
  if (data['Nama Peserta 2']) {
    peserta2Info = `
      <li><strong>Nama Peserta 2:</strong> ${data['Nama Peserta 2']}</li>
      <li><strong>No. HP/WA Peserta 2:</strong> ${data['No. HP/WA Peserta 2']}</li>
    `;
  }

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px;">
      <h2 style="color: #2d6a4f;">Bukti Pendaftaran Sosialisasi Bulan Keluarga 2026</h2>
      <p>Syalom, terima kasih telah melakukan pendaftaran. Berikut adalah rincian data yang telah dikirimkan:</p>
      
      <ul style="background-color: #f8f9fa; padding: 15px 25px; border-radius: 6px; list-style-type: none;">
        <li><strong>Wilayah:</strong> ${data['Wilayah']}</li>
        <li><strong>Lingkungan:</strong> ${data['Lingkungan']}</li>
        <li><strong>Nama Peserta 1:</strong> ${data['Nama Peserta 1']}</li>
        <li><strong>No. HP/WA Peserta 1:</strong> ${data['No. HP/WA Peserta 1']}</li>
        ${peserta2Info}
      </ul>

      <p><strong>Pengingat Acara:</strong><br>
      📅 Jumat, 20 November 2026<br>
      🕐 Jam 17:00 - 21:00 WIB<br>
      📍 Grha Pemuda Katedral Jakarta – Lantai 4</p>

      <p><em>Mohon membawa alat tulis dan botol minum pribadi, serta uang pas untuk pengganti lilin corona Adven (bagi yang membeli).</em></p>
      
      <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 12px; color: #777;">Email ini dikirim secara otomatis. Mohon tidak membalas email ini.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: userEmail,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName,
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    // Bot honeypot check
    if (e.parameter.mobile_number && e.parameter.mobile_number !== '') {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'success', message: 'Bot detected' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    let sheet = doc.getSheetByName(e.parameter.sheet_name || 'BunSpreadsheet');
    
    // Fallback if sheet tab is not found
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;

    const newRow = headers.map(function (header) {
      if (header === 'id') return Utilities.getUuid();
      if (header === 'timestamp') return new Date();

      const rawValue = e.parameter[header] || '';
      return sanitizeValue(rawValue);
    });

    const newRange = sheet.getRange(nextRow, 1, 1, newRow.length);
    newRange.setNumberFormat('@');
    newRange.setValues([newRow]);

    // Send admin notification
    sendAdminNotification(nextRow);

    // Send confirmation email to user IF email field is present and non-empty
    const userEmail = e.parameter.email || e.parameter.Email;
    if (userEmail && userEmail.trim() !== '') {
      sendUserConfirmationEmail(userEmail.trim(), e.parameter);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'success', row: nextRow })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}