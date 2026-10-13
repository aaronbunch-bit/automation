/**
 * Manager + Sales Coach nudges for active Reflex-AI training offers.
 *
 * Adds support for extending the manager alias section on TS Config from:
 *   Manager Name | Slack Alias | Notes
 *
 * to:
 *   Manager Name | Slack Alias | Coach Name | Coach Slack Alias | Notes
 *
 * The nudge action groups active, unbooked offers by manager and sends one
 * shared Slack DM to the manager and their sales coach listing reps who have
 * not actioned their offer.
 */

function setupManagerCoachColumns() {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('TS Config was not found. Run Setup workbook + triggers first.');
    return;
  }

  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  if (sectionRow < 0) {
    SpreadsheetApp.getUi().alert('Manager alias section was not found. Run Refresh manager list from Looker first.');
    return;
  }

  tsEnsureManagerCoachConfigColumns_(sheet, sectionRow);
  SpreadsheetApp.getUi().alert(
    'Manager coach columns are ready.\n\n' +
    'Fill Coach Name and Coach Slack Alias next to each manager in TS Config.'
  );
}

function testNudgeManagersAndCoaches() {
  nudgeManagersAndCoaches_(true, 'Aaron Bunch');
}

function nudgeManagersAndCoaches() {
  nudgeManagersAndCoaches_(false, '');
}

function nudgeManagersAndCoaches_(testMode, managerFilter) {
  var ss = tsGetSpreadsheet_();
  var offersSheet = ss.getSheetByName(TS.SHEETS.OFFERS);
  if (!offersSheet || offersSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No TS Offers rows found.');
    return;
  }

  if (!tsValidateSlackReady_()) {
    SpreadsheetApp.getUi().alert(
      'SLACK_BOT_TOKEN is missing.\n\nAdd it to TS Config or Script Properties, then try again.'
    );
    return;
  }

  var configSheet = ss.getSheetByName(TS.SHEETS.CONFIG);
  var sectionRow = configSheet ? tsFindManagerAliasSectionRow_(configSheet) : -1;
  if (sectionRow < 0) {
    SpreadsheetApp.getUi().alert('Manager alias section was not found in TS Config.');
    return;
  }

  tsEnsureManagerCoachConfigColumns_(configSheet, sectionRow);

  var coachConfig = tsLoadManagerCoachConfig_(configSheet, sectionRow);
  var groups = tsBuildManagerNudgeGroups_(offersSheet);
  var normalizedManagerFilter = tsNormalizePersonName_(managerFilter || '');
  var managerNames = Object.keys(groups).filter(function(managerName) {
    return !normalizedManagerFilter || tsNormalizePersonName_(managerName) === normalizedManagerFilter;
  }).sort();
  var sent = 0;
  var skippedNoCoach = 0;
  var skippedNoSlack = 0;

  managerNames.forEach(function(managerName) {
    var offers = groups[managerName];
    if (!offers.length) return;

    var route = coachConfig[tsNormalizePersonName_(managerName)] || {};
    var managerAlias = route.managerAlias || tsManagerNameToSlackAlias_(managerName);
    var coachAlias = route.coachAlias || '';

    if (!coachAlias) {
      skippedNoCoach++;
      tsAudit_('MANAGER_COACH_NUDGE', managerName, 'Missing coach alias in TS Config', 'WARN');
      return;
    }

    var message = tsBuildManagerCoachNudgeMessage_(managerName, route.coachName || '', managerAlias, coachAlias, offers);
    if (testMode) {
      message = '*[TEST MODE — Aaron Bunch manager/coach nudge]*\n\n' + message;
    }
    var ok = tsSendManagerCoachSeparateDms_(managerAlias, coachAlias, message);

    if (ok) {
      sent++;
      tsAudit_('MANAGER_COACH_NUDGE', managerName, (testMode ? 'Test sent' : 'Sent') + ' nudge for ' + offers.length + ' active offer(s)', 'OK');
    } else {
      skippedNoSlack++;
      tsAudit_('MANAGER_COACH_NUDGE', managerName, 'Failed Slack group DM', 'WARN');
    }
  });

  SpreadsheetApp.getUi().alert(
    (testMode ? 'Test Aaron manager + coach nudge complete.\n\n' : 'Manager + Coach nudges complete.\n\n') +
    'Messages sent: ' + sent + '\n' +
    'Skipped missing coach alias: ' + skippedNoCoach + '\n' +
    'Skipped Slack failures: ' + skippedNoSlack
  );
}

function sendManagerCoachNudgeForOfferRowOnce_(offerSheet, rowNum, row) {
  var tracking = tsEnsureManagerCoachNudgeTrackingColumns_(offerSheet);
  var sentAt = row[tracking.sentAtCol - 1];

  if (sentAt instanceof Date && !isNaN(sentAt.getTime())) {
    return false;
  }

  if (!tsOfferRowIsLatestActionableForNudge_(offerSheet, row)) {
    offerSheet.getRange(rowNum, tracking.resultCol).setValue('Skipped - latest offer status is not OFFERED/PENDING');
    return false;
  }

  var configSheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  var sectionRow = configSheet ? tsFindManagerAliasSectionRow_(configSheet) : -1;
  if (sectionRow < 0) {
    offerSheet.getRange(rowNum, tracking.resultCol).setValue('Missing manager alias section');
    return false;
  }

  tsEnsureManagerCoachConfigColumns_(configSheet, sectionRow);

  var coachConfig = tsLoadManagerCoachConfig_(configSheet, sectionRow);
  var managerName = String(row[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim();
  var route = coachConfig[tsNormalizePersonName_(managerName)] || {};
  var managerAlias = route.managerAlias || tsManagerNameToSlackAlias_(managerName);
  var coachAlias = route.coachAlias || '';

  if (!managerName || !coachAlias) {
    offerSheet.getRange(rowNum, tracking.resultCol).setValue('Missing manager or coach alias');
    return false;
  }

  var offer = {
    consultantName: String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim(),
    consultantEmail: String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim(),
    salesGroup: String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim(),
    simsCsv: String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '').trim(),
    durationMin: Number(row[TS.OFFER_COLS.DURATION_MIN - 1] || 0),
    offerSentAt: row[TS.OFFER_COLS.OFFER_SENT_AT - 1],
    createdAt: row[TS.OFFER_COLS.CREATED_AT - 1]
  };

  var message = tsBuildSingleOfferManagerCoachNudgeMessage_(
    managerName,
    route.coachName || '',
    managerAlias,
    coachAlias,
    offer
  );
  var ok = tsSendManagerCoachSeparateDms_(managerAlias, coachAlias, message);

  if (ok) {
    offerSheet.getRange(rowNum, tracking.sentAtCol).setValue(new Date());
    offerSheet.getRange(rowNum, tracking.resultCol).setValue('Sent');
    tsAudit_('MANAGER_COACH_NUDGE', offer.consultantEmail, '48h nudge sent for ' + managerName, 'OK');
    return true;
  }

  offerSheet.getRange(rowNum, tracking.resultCol).setValue('Slack send failed');
  tsAudit_('MANAGER_COACH_NUDGE', offer.consultantEmail, '48h nudge failed for ' + managerName, 'WARN');
  return false;
}

function tsEnsureManagerCoachNudgeTrackingColumns_(offerSheet) {
  var sentAtCol = tsFindOrCreateOfferColumn_(offerSheet, 'Manager Coach Nudge Sent At');
  var resultCol = tsFindOrCreateOfferColumn_(offerSheet, 'Manager Coach Nudge Result');

  return {
    sentAtCol: sentAtCol,
    resultCol: resultCol
  };
}

function tsFindOrCreateOfferColumn_(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === headerName) {
      return i + 1;
    }
  }

  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(headerName);
  sheet.getRange(1, col).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  return col;
}

function tsBuildSingleOfferManagerCoachNudgeMessage_(managerName, coachName, managerAlias, coachAlias, offer) {
  var lines = [
    '*Reflex-AI training follow-up needed*',
    '',
    'The following rep has not actioned their training offer after 48 hours:',
    '',
    '• *' + (offer.consultantName || offer.consultantEmail) + '* — ' +
      (offer.salesGroup || 'Unknown group') +
      (offer.durationMin ? ', ' + offer.durationMin + ' min' : '') +
      (offer.simsCsv ? ' — Sims: ' + offer.simsCsv : ''),
    '',
    'Both parties have been notified: manager `' + managerAlias + '` and coach `' + coachAlias + '`.',
    'Please coordinate to make sure this rep actions the most recent Ops Bot offer sent via DM.'
  ];

  return lines.join('\n');
}

/**
 * Call this inside processTrainingReminders after the rep reminder sends.
 *
 * Example placement:
 *   sheet.getRange(i + 1, TS.OFFER_COLS.REMINDER_SENT_AT).setValue(new Date());
 *   sendManagerCoachNudgeForReminderRow_(sheet, i + 1, row);
 */
function sendManagerCoachNudgeForReminderRow_(offerSheet, rowNum, row) {
  return sendManagerCoachNudgeForOfferRowOnce_(offerSheet, rowNum, row);
}

function tsEnsureManagerCoachConfigColumns_(sheet, sectionRow) {
  sheet.getRange(sectionRow, 1, 1, 5).setValues([[
    'Manager Name',
    'Slack Alias',
    'Coach Name',
    'Coach Slack Alias',
    'Notes'
  ]]);
  sheet.getRange(sectionRow, 1, 1, 5)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#ffffff');
}

function tsLoadManagerCoachConfig_(sheet, sectionRow) {
  var config = {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= sectionRow) return config;

  var values = sheet.getRange(sectionRow + 1, 1, lastRow - sectionRow, 5).getValues();
  values.forEach(function(row) {
    var managerName = String(row[0] || '').trim();
    if (!managerName || managerName.indexOf('(no managers found') === 0) return;

    config[tsNormalizePersonName_(managerName)] = {
      managerName: managerName,
      managerAlias: String(row[1] || '').trim(),
      coachName: String(row[2] || '').trim(),
      coachAlias: String(row[3] || '').trim()
    };
  });

  return config;
}

function tsBuildManagerNudgeGroups_(offersSheet) {
  var values = offersSheet.getDataRange().getValues();
  var groups = {};
  var latestRows = tsLatestOfferRowsByRepGroup_(values);

  Object.keys(latestRows).sort().forEach(function(key) {
    var row = latestRows[key];
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (!tsOfferStatusNeedsManagerCoachNudge_(status)) return;
    if (tsOfferRowInactiveByHeader_(values[0], row)) return;

    var managerName = String(row[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim();
    if (!managerName) return;

    if (!groups[managerName]) groups[managerName] = [];
    groups[managerName].push({
      consultantName: String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim(),
      consultantEmail: String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim(),
      salesGroup: String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim(),
      simsCsv: String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '').trim(),
      durationMin: Number(row[TS.OFFER_COLS.DURATION_MIN - 1] || 0),
      offerSentAt: row[TS.OFFER_COLS.OFFER_SENT_AT - 1],
      createdAt: row[TS.OFFER_COLS.CREATED_AT - 1]
    });
  });

  return groups;
}

function tsLatestOfferRowsByRepGroup_(values) {
  var latest = {};
  if (!values || values.length < 2) return latest;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var key = tsOfferRepGroupKey_(row);
    if (!key) continue;

    if (!latest[key] || tsOfferRowSortTime_(row) >= tsOfferRowSortTime_(latest[key])) {
      latest[key] = row;
    }
  }

  return latest;
}

function tsOfferRowIsLatestActionableForNudge_(offersSheet, row) {
  var values = offersSheet.getDataRange().getValues();
  var latestRows = tsLatestOfferRowsByRepGroup_(values);
  var key = tsOfferRepGroupKey_(row);
  if (!key || !latestRows[key]) return false;

  var latestRow = latestRows[key];
  var status = String(latestRow[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

  return latestRow === row &&
    tsOfferStatusNeedsManagerCoachNudge_(status) &&
    !tsOfferRowInactiveByHeader_(values[0], latestRow);
}

function tsOfferRepGroupKey_(row) {
  var email = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
  var salesGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
  if (!email || !salesGroup) return '';
  return email + '|' + salesGroup;
}

function tsOfferRowSortTime_(row) {
  var bookedAt = tsNudgeParseDate_(row[TS.OFFER_COLS.BOOKED_AT - 1]);
  var createdAt = tsNudgeParseDate_(row[TS.OFFER_COLS.CREATED_AT - 1]);
  var offerSentAt = tsNudgeParseDate_(row[TS.OFFER_COLS.OFFER_SENT_AT - 1]);
  var best = bookedAt || createdAt || offerSentAt;
  return best ? best.getTime() : 0;
}

function tsNudgeParseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var raw = String(value || '').trim();
  if (!raw) return null;
  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function tsOfferStatusNeedsManagerCoachNudge_(status) {
  return ['OFFERED', 'PENDING'].indexOf(String(status || '').trim().toUpperCase()) !== -1;
}

function tsOfferRowInactiveByHeader_(headers, row) {
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i] || '').trim().toLowerCase();
    if (header === 'active or inactive' || header === 'active/inactive' || header === 'active') {
      return String(row[i] || '').trim().toLowerCase() === 'inactive';
    }
  }
  return false;
}

function tsBuildManagerCoachNudgeMessage_(managerName, coachName, managerAlias, coachAlias, offers) {
  var lines = [
    '*Reflex-AI training follow-up needed*',
    '',
    'The following reps under *' + managerName + '* have not actioned their training offer yet' +
      (coachName ? ' (coach: *' + coachName + '*)' : '') + ':',
    ''
  ];

  offers.forEach(function(offer) {
    lines.push(
      '• *' + (offer.consultantName || offer.consultantEmail) + '* — ' +
      (offer.salesGroup || 'Unknown group') +
      (offer.durationMin ? ', ' + offer.durationMin + ' min' : '') +
      (offer.simsCsv ? ' — Sims: ' + offer.simsCsv : '')
    );
  });

  lines.push('');
  lines.push('Both parties have been notified: manager `' + managerAlias + '` and coach `' + coachAlias + '`.');
  lines.push('Please coordinate to make sure these reps action the most recent Ops Bot offer sent via DM.');

  return lines.join('\n');
}

function tsSendManagerCoachSeparateDms_(managerAlias, coachAlias, message) {
  var managerUserId = tsSlackLookupUserId_(managerAlias);
  var coachUserId = tsSlackLookupUserId_(coachAlias);

  if (!managerUserId || !coachUserId) {
    return false;
  }

  if (managerUserId === coachUserId) {
    return tsSlackSendDm_(managerUserId, message);
  }

  var managerOk = tsSlackSendDm_(managerUserId, message);
  var coachOk = tsSlackSendDm_(coachUserId, message);
  return managerOk && coachOk;
}
