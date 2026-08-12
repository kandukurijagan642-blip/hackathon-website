var SPREADSHEET_ID     = '1fPRAqTe3eGhsW46vpQ0LgjjOwK6JEhpOevG85zAU0JQ';
var TELEGRAM_BOT_TOKEN = '8766828763:AAGi68e9f5_tXEcvi3UQv8pitRVTxncYlhs';
var TELEGRAM_CHAT_ID   = '6877857251';
var ORGANIZER_EMAIL    = 'kandukurijagan99@gmail.com';

/**
 * GET handler: Live API check
 */
function doGet(e) {
  return ContentService.createTextOutput('✅ Synora Registration API is LIVE!')
                       .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST handler: Registers new team, decodes uploaded photo, saves to Google Drive, forwards to Email & Telegram, and logs to sheet.
 */
function doPost(e) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet) {
      sheet = ss.insertSheet('Registrations');
    }

    // Force headers on row 1 - dynamically sized to prevent range errors
    var headers = [
      'Timestamp',
      'Team Name',
      'Team Leader Name', 'Team Leader Mail', 'Team Leader Mobile',
      'Member 1 Name',    'Member 1 Mail',    'Member 1 Phone',
      'Member 2 Name',    'Member 2 Mail',    'Member 2 Phone',
      'Member 3 Name',    'Member 3 Mail',    'Member 3 Phone',
      'Registration Type', 'Reg Number', 'Transaction ID', 'Attachment Link'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var p              = e.parameter;
    var teamName       = p.teamName         || '';
    var teamLeaderName = p.teamLeaderName   || '';
    var teamLeaderMail = p.teamLeaderMail   || '';
    var teamLeaderMob  = p.teamLeaderMobile || '';
    var m1Name         = p.member1Name      || '';
    var m1Mail         = p.member1Mail      || '';
    var m1Phone        = p.member1Phone     || '';
    var m2Name         = p.member2Name      || '';
    var m2Mail         = p.member2Mail      || '';
    var m2Phone        = p.member2Phone     || '';
    var m3Name         = p.member3Name      || '';
    var m3Mail         = p.member3Mail      || '';
    var m3Phone        = p.member3Phone     || '';
    var regType        = p.regType          || 'internal';
    var internalRegNo  = p.internalRegNo    || '';
    var externalTxnId  = p.externalTxnId    || '';
    var ts             = new Date();

    // ─── DUPLICATE REGISTRATION CHECKS ──────────────────────────────────
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    // Normalise inputs for comparison
    var compTeam = teamName.toLowerCase().replace(/\s+/g, '');
    var compMail = teamLeaderMail.toLowerCase().trim();
    var compMobile = teamLeaderMob.replace(/\D/g, '');

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row.length > 4) {
        var existingTeam = row[1].toString().toLowerCase().replace(/\s+/g, '');
        var existingMail = row[3].toString().toLowerCase().trim();
        var existingMobile = row[4].toString().replace(/\D/g, '');

        if (compTeam && existingTeam === compTeam) {
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error', 
            message: 'Team Name "' + teamName + '" is already registered! Please choose a different team name.'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        if (compMail && existingMail === compMail) {
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error', 
            message: 'Leader Email "' + teamLeaderMail + '" is already registered with another team!'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        if (compMobile && existingMobile === compMobile) {
          return ContentService.createTextOutput(JSON.stringify({
            status: 'error', 
            message: 'Leader Mobile number is already registered with another team!'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // ─── PROCESS PHOTO BLOB & SAVE TO GOOGLE DRIVE ─────────────────────
    var photoBlob = null;
    var attachmentUrl = '';
    
    if (p.fileData && p.fileName && p.fileMime) {
      try {
        var decodedData = Utilities.base64Decode(p.fileData);
        photoBlob = Utilities.newBlob(decodedData, p.fileMime, teamName.replace(/\s+/g, '_') + '_' + p.fileName);
        
        // Save to Google Drive (directly to root to prevent folder permission issues)
        var file = DriveApp.createFile(photoBlob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        attachmentUrl = file.getUrl();
      } catch(fileErr) {
        Logger.log('Drive save error: ' + fileErr.toString());
        attachmentUrl = 'Error saving to Drive: ' + fileErr.toString();
      }
    }

    // Append row to sheet
    sheet.appendRow([
      ts,
      teamName,
      teamLeaderName, teamLeaderMail, teamLeaderMob,
      m1Name,    m1Mail,    m1Phone,
      m2Name,    m2Mail,    m2Phone,
      m3Name,    m3Mail,    m3Phone,
      regType.toUpperCase(),
      internalRegNo,
      externalTxnId,
      attachmentUrl
    ]);

    // ─── SEND EMAIL NOTIFICATION WITH ATTACHMENT ────────────────────
    try {
      var emailBody = 'TEAM Name: ' + teamName + '\n' +
                      'Status   : ' + regType.toUpperCase() + '\n';
      
      if (regType === 'internal') {
        emailBody += 'College Reg Number: ' + internalRegNo + '\n';
      } else {
        emailBody += 'Transaction ID / UPI Reference: ' + externalTxnId + '\n';
      }
      
      emailBody += '\nATTACHMENT LINK:\n' + (attachmentUrl || 'None') + '\n\n' +
                   'TEAM LEADER\n' +
                   'Name   : ' + teamLeaderName   + '\n' +
                   'Email  : ' + teamLeaderMail   + '\n' +
                   'Mobile : ' + teamLeaderMob + '\n\n' +
                   'MEMBER 1\n' +
                   'Name  : ' + m1Name  + '\n' +
                   'Email : ' + m1Mail  + '\n' +
                   'Phone : ' + m1Phone + '\n\n' +
                   (m2Name ? 'MEMBER 2\nName: ' + m2Name + '\nEmail: ' + m2Mail + '\n\n' : '') +
                   (m3Name ? 'MEMBER 3\nName: ' + m3Name + '\nEmail: ' + m3Mail + '\n\n' : '') +
                   'Time: ' + ts.toLocaleString();

      var emailOptions = {
        to: ORGANIZER_EMAIL,
        subject: 'New Hackathon Registration [' + regType.toUpperCase() + ']: ' + teamName,
        body: emailBody
      };

      MailApp.sendEmail(emailOptions);
    } catch(mailErr) { 
      Logger.log('Email error: ' + mailErr.toString()); 
    }

    // ─── SEND TELEGRAM MESSAGE & PHOTO ──────────────────────────────
    try {
      var telegramText = 'NEW HACKATHON REGISTRATION\n\n' +
                         'Team   : ' + teamName         + '\n' +
                         'Status : ' + regType.toUpperCase() + '\n';

      if (regType === 'internal') {
        telegramText += 'Reg No : ' + internalRegNo + '\n';
      } else {
        telegramText += 'Txn ID : ' + externalTxnId + '\n';
      }

      telegramText += 'Leader : ' + teamLeaderName   + '\n' +
                      'Email  : ' + teamLeaderMail   + '\n' +
                      'Mobile : ' + teamLeaderMob + '\n\n' +
                      'Member1: ' + m1Name  + '\n' +
                      'Email  : ' + m1Mail  + '\n' +
                      'Phone  : ' + m1Phone +
                      (m2Name ? '\nMember2: ' + m2Name : '') +
                      (m3Name ? '\nMember3: ' + m3Name : '') +
                      '\n\nAttachment URL: ' + (attachmentUrl || 'None') +
                      '\n\nshared to telegram and mail' +
                      '\n\nTime: ' + ts.toLocaleString();

      UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: TELEGRAM_CHAT_ID,
          text: telegramText
        }),
        muteHttpExceptions: true
      });
    } catch(tgErr) { 
      Logger.log('Telegram error: ' + tgErr.toString()); 
    }

    return ContentService
      .createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('doPost error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Run this function manually in the Google Apps Script editor (click 'Run')
 * to trigger the authorization prompt and grant Google Drive & Mail permissions.
 */
function authorizeScript() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  Logger.log("✅ Script successfully authorized!");
}

