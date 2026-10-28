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
      <p>Syalom, terima kasih telah melakukan pendaftaran Sosialisasi Bulan Keluarga. Berikut adalah rincian data yang telah dikirimkan:</p>
      
      <ul style="background-color: #f8f9fa; padding: 15px 25px; border-radius: 6px; list-style-type: none;">
        <li><strong>Wilayah:</strong> ${data['Wilayah']}</li>
        <li><strong>Lingkungan:</strong> ${data['Lingkungan']}</li>
        <li><strong>Beli Lilin Adven:</strong> ${data['Beli Lilin']}</li>
        <li><strong>Nama Peserta 1:</strong> ${data['Nama Peserta 1']}</li>
        <li><strong>No. HP/WA Peserta 1:</strong> ${data['No. HP/WA Peserta 1']}</li>
        ${peserta2Info}
      </ul>

      <p><strong>Pengingat Acara:</strong><br>
      📅 Jumat, 20 November 2026<br>
      🕐 Jam 18:30 - 21:00 WIB<br>
      📍 Grha Pemuda Katedral Jakarta – Lantai 4</p>

      <p><em>Mohon membawa alat tulis dan botol minum pribadi, serta uang pas untuk pengganti lilin corona Adven.</em></p>
      <p><b>Contact person:</b> Santi 08-5959-8877-12 / Buntoro 0823-1151-4658 </p>
      
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
  const success = lock.tryLock(10000);

  if (!success) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', message: 'Server sedang sibuk, silakan coba beberapa saat lagi.' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
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
    const newBeliLilin = (e.parameter['Beli Lilin'] || '').trim();

    const colIndex = {
      lingk: headers.indexOf('Lingkungan') + 1,
      nama1: headers.indexOf('Nama Peserta 1') + 1,
      hp1: headers.indexOf('No. HP/WA Peserta 1') + 1,
      nama2: headers.indexOf('Nama Peserta 2') + 1,
      hp2: headers.indexOf('No. HP/WA Peserta 2') + 1,
      beliLilin: headers.indexOf('Beli Lilin') + 1,
      email: headers.indexOf('email') !== -1 ? headers.indexOf('email') + 1 : headers.indexOf('Email') + 1
    };

    let existingRowIndex = -1;
    if (colIndex.lingk > 0 && sheet.getLastRow() > 1) {
      const existingLingks = sheet.getRange(2, colIndex.lingk, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existingLingks.length; i++) {
        if (existingLingks[i][0].toString().trim().toLowerCase() === submittedLingk.toLowerCase()) {
          existingRowIndex = i + 2; 
          break;
        }
      }
    }

    if (existingRowIndex !== -1) {
      const curNama1 = sheet.getRange(existingRowIndex, colIndex.nama1).getValue().toString().trim();
      const curNama2 = sheet.getRange(existingRowIndex, colIndex.nama2).getValue().toString().trim();

      if (curNama1 !== '' && curNama2 !== '') {
        return ContentService.createTextOutput(
          JSON.stringify({
            result: 'quota_full',
            message: `Lingkungan "${submittedLingk}" sudah memenuhi kuota maksimal 2 peserta.`
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      if (curNama1 !== '' && curNama2 === '') {
        if (newNama2 !== '') {
          return ContentService.createTextOutput(
            JSON.stringify({
              result: 'quota_exceeded',
              message: `Lingkungan "${submittedLingk}" sudah mendaftarkan 1 peserta sebelumnya. Sisa kuota hanya 1 peserta lagi. Hapus Data Peserta 2 lalu submit.`
            })
          ).setMimeType(ContentService.MimeType.JSON);
        }

        sheet.getRange(existingRowIndex, colIndex.nama2).setNumberFormat('@').setValue(sanitizeValue(newNama1));
        sheet.getRange(existingRowIndex, colIndex.hp2).setNumberFormat('@').setValue(sanitizeValue(newHp1));
        
        let finalBeliLilin = newBeliLilin;
        let warningMsg = "";

        if (colIndex.beliLilin > 0) {
          const curBeliLilin = sheet.getRange(existingRowIndex, colIndex.beliLilin).getValue().toString().trim();
          
          if (curBeliLilin === 'Ya') {
            finalBeliLilin = 'Ya';
            if (newBeliLilin === 'Tidak') {
              warningMsg = "Orang lingkungan anda yang sudah mendaftar membeli lilin, jika lingkungan mau membatalkan, hubungi panitia.";
            }
          } else if (newBeliLilin !== '') {
            sheet.getRange(existingRowIndex, colIndex.beliLilin).setValue(sanitizeValue(newBeliLilin));
            finalBeliLilin = newBeliLilin;
          }
          
          e.parameter['Beli Lilin'] = finalBeliLilin;
        }

        const newEmail = (e.parameter.email || e.parameter.Email || '').trim();
        if (newEmail !== '' && colIndex.email > 0) {
          sheet.getRange(existingRowIndex, colIndex.email).setValue(sanitizeValue(newEmail));
        }

        sendAdminNotification(existingRowIndex, `Peserta 2 added to ${submittedLingk}`);

        if (newEmail !== '') {
          sendUserConfirmationEmail(newEmail, e.parameter);
        }

        return ContentService.createTextOutput(
          JSON.stringify({ 
            result: 'success', 
            row: existingRowIndex, 
            isSecondary: true,
            finalBeliLilin: finalBeliLilin,
            warningMessage: warningMsg
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

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
      JSON.stringify({ 
        result: 'success', 
        row: nextRow, 
        isSecondary: false,
        finalBeliLilin: newBeliLilin
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    let sheet = doc.getSheetByName('BunSpreadsheet');
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    const idxWilayah = headers.indexOf('wilayah');
    const idxLingk = headers.indexOf('lingkungan');
    const idxNama1 = headers.indexOf('nama peserta 1');
    const idxNama2 = headers.indexOf('nama peserta 2');
    const idxBeliLilin = headers.indexOf('beli lilin'); // Tambahan kolom lilin
    
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const wVal = idxWilayah !== -1 && row[idxWilayah] ? row[idxWilayah].toString().trim() : '';
      const lVal = idxLingk !== -1 && row[idxLingk] ? row[idxLingk].toString().trim() : '';
      const p1Val = idxNama1 !== -1 && row[idxNama1] ? row[idxNama1].toString().trim() : '';
      const p2Val = idxNama2 !== -1 && row[idxNama2] ? row[idxNama2].toString().trim() : '';
      const lilinVal = idxBeliLilin !== -1 && row[idxBeliLilin] ? row[idxBeliLilin].toString().trim() : ''; // Ekstrak data lilin
      
      if (wVal && lVal) {
        result.push({ 
          wilayah: wVal, 
          lingkungan: lVal, 
          nama1: p1Val, 
          nama2: p2Val,
          beliLilin: lilinVal // Kirim data lilin ke frontend
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
