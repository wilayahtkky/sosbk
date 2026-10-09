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

function sendAdminNotification(rowNumber, statusMessage) {
  const recipient = 'wilayahtkky@gmail.com';
  const senderName = 'wil TKKY - Buntoro';
  const subject = 'Registration Update: Sosialisasi Bulan Keluarga 2026';
  const body = `Status: ${statusMessage} (Row ${rowNumber})`;

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
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const submittedLingk = (e.parameter.Lingkungan || '').trim();
    const newNama1 = (e.parameter['Nama Peserta 1'] || '').trim();
    const newHp1 = (e.parameter['No. HP/WA Peserta 1'] || '').trim();
    const newNama2 = (e.parameter['Nama Peserta 2'] || '').trim();
    const newHp2 = (e.parameter['No. HP/WA Peserta 2'] || '').trim();

    // Column Index Map (1-based for Sheets API)
    const colIndex = {
      lingk: headers.indexOf('Lingkungan') + 1,
      nama1: headers.indexOf('Nama Peserta 1') + 1,
      hp1: headers.indexOf('No. HP/WA Peserta 1') + 1,
      nama2: headers.indexOf('Nama Peserta 2') + 1,
      hp2: headers.indexOf('No. HP/WA Peserta 2') + 1,
      email: headers.indexOf('email') !== -1 ? headers.indexOf('email') + 1 : headers.indexOf('Email') + 1
    };

    // Find if Lingkungan already exists in the sheet
    let existingRowIndex = -1;
    if (colIndex.lingk > 0 && sheet.getLastRow() > 1) {
      const existingLingks = sheet.getRange(2, colIndex.lingk, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existingLingks.length; i++) {
        if (existingLingks[i][0].toString().trim().toLowerCase() === submittedLingk.toLowerCase()) {
          existingRowIndex = i + 2; // +2 for 1-based indexing and header row
          break;
        }
      }
    }

    // --- CASE 1: Lingkungan already exists ---
    if (existingRowIndex !== -1) {
      const curNama1 = sheet.getRange(existingRowIndex, colIndex.nama1).getValue().toString().trim();
      const curNama2 = sheet.getRange(existingRowIndex, colIndex.nama2).getValue().toString().trim();

      // Sub-case 1A: Both slots are already full
      if (curNama1 !== '' && curNama2 !== '') {
        return ContentService.createTextOutput(
          JSON.stringify({
            result: 'quota_full',
            message: `Lingkungan "${submittedLingk}" sudah memenuhi kuota maksimal 2 peserta.`
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Sub-case 1B: Slot 1 is taken, Slot 2 is open
      if (curNama1 !== '' && curNama2 === '') {
        // Reject if submitter entered 2 names at once
        if (newNama2 !== '') {
          return ContentService.createTextOutput(
            JSON.stringify({
              result: 'quota_exceeded',
              message: `Lingkungan "${submittedLingk}" sudah mendaftarkan 1 peserta sebelumnya. Sisa kuota hanya 1 peserta lagi. Hapus Data Peserta 2 lalu submit`
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }

        // Fill incoming Peserta 1 into Peserta 2 slots
        sheet.getRange(existingRowIndex, colIndex.nama2).setNumberFormat('@').setValue(sanitizeValue(newNama1));
        sheet.getRange(existingRowIndex, colIndex.hp2).setNumberFormat('@').setValue(sanitizeValue(newHp1));

        // Update email if provided
        const newEmail = (e.parameter.email || e.parameter.Email || '').trim();
        if (newEmail !== '' && colIndex.email > 0) {
          sheet.getRange(existingRowIndex, colIndex.email).setValue(sanitizeValue(newEmail));
        }

        sendAdminNotification(existingRowIndex, `Peserta 2 added to ${submittedLingk}`);

        if (newEmail !== '') {
          sendUserConfirmationEmail(newEmail, e.parameter);
        }

        return ContentService.createTextOutput(
          JSON.stringify({ result: 'success', row: existingRowIndex, isSecondary: true })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- CASE 2: New Lingkungan (Append as brand new row) ---
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

    sendAdminNotification(nextRow, `New Registration for ${submittedLingk}`);

    const userEmail = e.parameter.email || e.parameter.Email;
    if (userEmail && userEmail.trim() !== '') {
      sendUserConfirmationEmail(userEmail.trim(), e.parameter);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'success', row: nextRow, isSecondary: false })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
// Add this function to your Code.gs
function doGet(e) {
  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    let sheet = doc.getSheetByName('BunSpreadsheet');
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    // Find column positions dynamically
    const idxWilayah = headers.indexOf('wilayah');
    const idxLingk = headers.indexOf('lingkungan');
    const idxNama1 = headers.indexOf('nama peserta 1');
    const idxNama2 = headers.indexOf('nama peserta 2');

    const result = [];

    // Loop through rows (skip header row index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const wVal = idxWilayah !== -1 && row[idxWilayah] ? row[idxWilayah].toString().trim() : '';
      const lVal = idxLingk !== -1 && row[idxLingk] ? row[idxLingk].toString().trim() : '';
      const p1Val = idxNama1 !== -1 && row[idxNama1] ? row[idxNama1].toString().trim() : '';
      const p2Val = idxNama2 !== -1 && row[idxNama2] ? row[idxNama2].toString().trim() : '';

      if (wVal && lVal) {
        // ONLY expose non-sensitive fields
        result.push({
          wilayah: wVal,
          lingkungan: lVal,
          nama1: p1Val,
          nama2: p2Val
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}