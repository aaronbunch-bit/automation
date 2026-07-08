/**
 * Permanent repair utility for existing Reflex-AI Training Scheduler offers.
 *
 * Adds/refreshes an "Escalate Link" column in TS Offers using the current
 * TS_WEB_APP_URL and each offer row's Escalate Token. This lets operators send
 * a corrected "None of these times work" link without resending the full offer.
 *
 * Requires the main scheduler file to be present because it uses:
 * - TS
 * - tsGetSpreadsheet_
 * - tsGetWebAppUrl_
 * - tsEnsureOffersSheet_
 * - tsEnsureOfferEscalateToken_
 * - tsAudit_
 */
function repairEscalationLinks() {
  var ss = tsGetSpreadsheet_();
  var sheet = tsEnsureOffersSheet_(ss);
  var webAppUrl = tsGetWebAppUrl_();

  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert(
      'Repair Escalation Links',
      'TS_WEB_APP_URL is blank. Add the deployed /exec web app URL to TS Config or Script Properties first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var normalizedWebAppUrl = String(webAppUrl).replace(/\/+$/, '');
  var linkCol = tsEnsureEscalateLinkColumn_(sheet);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert(
      'Repair Escalation Links',
      'No TS Offers rows found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), linkCol)).getValues();
  var updated = 0;
  var skippedBooked = 0;
  var tokenCreated = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 1;
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (status === 'BOOKED') {
      skippedBooked++;
      continue;
    }

    var tokenBefore = String(row[TS.OFFER_COLS.ESCALATE_TOKEN - 1] || '').trim();
    var token = tsEnsureOfferEscalateToken_(sheet, rowNum, row);

    if (!tokenBefore && token) {
      tokenCreated++;
    }

    if (!token) {
      continue;
    }

    sheet.getRange(rowNum, linkCol).setValue(tsBuildEscalateUrl_(normalizedWebAppUrl, token));
    updated++;
  }

  sheet.autoResizeColumn(linkCol);
  SpreadsheetApp.flush();
  tsAudit_(
    'REPAIR_ESCALATE_LINKS',
    '',
    'Updated ' + updated + ' link(s); created ' + tokenCreated + ' token(s); skipped ' + skippedBooked + ' booked offer(s).',
    'OK'
  );

  SpreadsheetApp.getUi().alert(
    'Repair Escalation Links',
    'Escalation links refreshed.\n\n' +
      'Updated links: ' + updated + '\n' +
      'New tokens created: ' + tokenCreated + '\n' +
      'Booked offers skipped: ' + skippedBooked + '\n\n' +
      'Use the Escalate Link column in TS Offers for any rep whose old Slack link is broken.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function tsEnsureEscalateLinkColumn_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), TS.OFFER_COLS.ESCALATE_TOKEN);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'Escalate Link') {
      return i + 1;
    }
  }

  var newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue('Escalate Link');
  sheet.getRange(1, newColumn).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  return newColumn;
}

function tsBuildEscalateUrl_(webAppUrl, escalateToken) {
  return webAppUrl + (webAppUrl.indexOf('?') === -1 ? '?' : '&') +
    'escalate=' + encodeURIComponent(escalateToken);
}
