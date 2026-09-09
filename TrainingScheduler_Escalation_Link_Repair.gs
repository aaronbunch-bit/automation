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

/**
 * Backward-compatible alias for menu entries accidentally configured with a
 * lowercase function name. Apps Script menu function names are case-sensitive.
 */
function repairescalationlinks() {
  repairEscalationLinks();
}

/**
 * Resends active offer Slack messages using the current TS_WEB_APP_URL.
 *
 * This is for repairing stale Book It / escalation URLs already sent in Slack.
 * It does not create new offers or new booking tokens. It reuses each active
 * offer row's Token IDs and Escalate Token, rebuilds the Slack message, and
 * sends it again.
 */
function resendActiveOffersWithCurrentLinks() {
  var ss = tsGetSpreadsheet_();
  var sheet = tsEnsureOffersSheet_(ss);
  var webAppUrl = tsGetWebAppUrl_();

  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert(
      'Resend Active Offers',
      'TS_WEB_APP_URL is blank. Add the deployed /exec web app URL to TS Config or Script Properties first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  if (!tsValidateSlackReady_()) {
    SpreadsheetApp.getUi().alert(
      'Resend Active Offers',
      'SLACK_BOT_TOKEN is missing. Add it to TS Config or Script Properties, then try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert(
      'Resend Active Offers',
      'No TS Offers rows found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  repairEscalationLinks();

  var values = sheet.getDataRange().getValues();
  var sent = 0;
  var skippedBooked = 0;
  var skippedEscalated = 0;
  var skippedNoTokens = 0;
  var failed = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 1;
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (status === 'BOOKED') {
      skippedBooked++;
      continue;
    }
    if (status === 'ESCALATED') {
      skippedEscalated++;
      continue;
    }

    var tokenIds = String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '')
      .split(',')
      .map(function(value) { return value.trim(); })
      .filter(Boolean);

    if (!tokenIds.length) {
      skippedNoTokens++;
      continue;
    }

    var need = tsBuildNeedFromOfferRow_(sheet, rowNum, row, tokenIds);
    var windows = tsBuildWindowsFromTokenIds_(tokenIds);

    if (!windows.length) {
      skippedNoTokens++;
      continue;
    }

    var message = tsBuildOfferSlackMessage_(need, windows, tokenIds);
    var ok = tsSendOfferSlackMessage_(need, message);

    if (ok) {
      sent++;
      sheet.getRange(rowNum, TS.OFFER_COLS.OFFER_SENT_AT).setValue(new Date());
      tsAudit_('RESEND_ACTIVE_OFFERS', need.email, 'Resent active offer with repaired links', 'OK');
    } else {
      failed++;
      tsAudit_('RESEND_ACTIVE_OFFERS', need.email, 'Failed to resend active offer with repaired links', 'WARN');
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Resend Active Offers',
    'Active offer resend complete.\n\n' +
      'Sent: ' + sent + '\n' +
      'Failed: ' + failed + '\n' +
      'Skipped booked: ' + skippedBooked + '\n' +
      'Skipped escalated: ' + skippedEscalated + '\n' +
      'Skipped missing tokens/windows: ' + skippedNoTokens + '\n\n' +
      'New Slack messages should now contain repaired Book It and escalation links.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Backward-compatible alias for lowercase menu naming.
 */
function resendactiveofferswithcurrentlinks() {
  resendActiveOffersWithCurrentLinks();
}

function tsBuildNeedFromOfferRow_(offerSheet, rowNum, row, tokenIds) {
  var escalateToken = tsEnsureOfferEscalateToken_(offerSheet, rowNum, row);

  return {
    name: String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim(),
    email: String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase(),
    salesGroup: String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim(),
    queue: String(row[TS.OFFER_COLS.QUEUE - 1] || '').trim(),
    managerName: String(row[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim(),
    incompleteCount: Number(row[TS.OFFER_COLS.INCOMPLETE_COUNT - 1] || 0),
    sims: String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '')
      .split(',')
      .map(function(value) { return value.trim(); })
      .filter(Boolean),
    durationMin: Number(row[TS.OFFER_COLS.DURATION_MIN - 1] || 30),
    tokenIds: tokenIds,
    escalateToken: escalateToken
  };
}

function tsBuildWindowsFromTokenIds_(tokenIds) {
  var windows = [];

  tokenIds.forEach(function(tokenId) {
    var hit = tsLookupBookingTokenRow_(tokenId);
    if (!hit) return;

    var status = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status !== 'PENDING') return;

    windows.push({
      dateStr: tsParseDateStr_(hit.row[TS.BOOK_COLS.DATE_STR - 1]),
      startStr: tsParseTimeInt_(hit.row[TS.BOOK_COLS.START_STR - 1]),
      endStr: tsParseTimeInt_(hit.row[TS.BOOK_COLS.END_STR - 1])
    });
  });

  return windows;
}

function tsSendOfferSlackMessage_(need, message) {
  if (tsIsTestMode_()) {
    return !!tsSlackDmTestRecipient_(
      '*[TEST MODE — repaired active offer would have gone to ' + (need.name || need.email) + ']*\n\n' + message
    );
  }

  var userId = tsSlackLookupUserId_(need.name) || tsSlackLookupUserIdByEmail_(need.email);
  if (!userId) {
    tsAudit_('SLACK', need.name || need.email, 'Could not resolve consultant for repaired offer resend', 'WARN');
    return false;
  }

  return !!tsSlackSendDm_(userId, message);
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
