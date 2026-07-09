/***************************************
 * REFLEX-AI TRAINING SCHEDULER v1.0.0
 *
 * Container-bound Apps Script for the Year-Round Workshop Automation workbook.
 * Reads ReflexAI CSV tabs, finds capacity-safe training windows (Assembled net
 * staffing), offers three Book-it slots via Slack (ops_bot), commits Training
 * activities, and creates events on the consultant's Google Calendar.
 *
 * Script Properties:
 *   ASSEMBLED_API_KEY
 *   SLACK_BOT_TOKEN          (ops_bot)
 *   TS_WEB_APP_URL           deployed web app URL ending in /exec
 *   ASSEMBLED_TRAINING_ACTIVITY_ID  (optional UUID override)
 *
 * Advanced service: Google Calendar API (consultant calendar events)
 *
 * Deploy web app: Execute as Me, Anyone with the link.
 ***************************************/

const TS = {
  VERSION: 'v1.0.7',
  TZ: 'America/Chicago',
  BRAND: 'Reflex-AI Training Scheduler',
  /** Container spreadsheet — required for Book-it web app (no active sheet in doGet). */
  SPREADSHEET_ID: '1ClH8uYgfiKMjPJ1-ooFj3bsTel1r6vHSZ3focyjLe0A',
  SHEETS: {
    CONFIG: 'TS Config',
    OFFERS: 'TS Offers',
    BOOKING_TOKENS: 'TS Booking Tokens',
    AUDIT: 'TS Audit',
    ANALYTICS: 'TS Analytics',
    MANAGER_LOOKUP: 'Looker Manager Lookup'
  },
  PASSING_SCORE_PERCENT: 80,
  REFLEX_TAB_PREFIX: 'ReflexAI CSV - ',
  ASSEMBLED: {
    BASE_URL: 'https://api.assembledhq.com/v0',
    SITE_NAME: 'Consumer Sales',
    CHANNEL: 'phone',
    INTERVAL: 1800,
    PAGE_SIZE: 20,
    SLEEP_MS: 300
  },
  TAB_TO_QUEUE: {
    'Prof Certs': 'Prof Certs_CC90_New',
    'ELD': 'Elementary and LD_CC90_New',
    'College': 'College and Grad TP_CC90_New',
    'Adult Learning': 'Adult Learner_CC90_New',
    'High School': 'High School_CC90_New'
  },
  MIN_BUFFER: 1,
  SEARCH_DAYS: 7,
  BOOKING_TOKEN_TTL_MS: 7 * 24 * 60 * 60 * 1000,
  REMINDER_AFTER_MS: 48 * 60 * 60 * 1000,
  BUSINESS_HOUR_START: 7,
  BUSINESS_HOUR_END: 21,
  /** Stop processing offers before the 6-min Apps Script limit. */
  OFFER_RUN_BUDGET_MS: 4.5 * 60 * 1000,
  /** Deployed Book-it web app — fallback if TS Config / Script Properties empty */
  WEB_APP_URL:
    'https://script.google.com/a/macros/varsitytutors.com/s/AKfycbzIXnGvmusKJF6aVrmbhnYKd7n48ZgzSjoVsfNbX7ROVQccJrJtWMKlhJLvKo4Bdo1w/exec',
  INCOMPLETE_STATUSES: ['not started', 'incomplete'],
  OFFER_COLS: {
    CREATED_AT: 1,
    CONSULTANT_NAME: 2,
    CONSULTANT_EMAIL: 3,
    SALES_GROUP: 4,
    QUEUE: 5,
    MANAGER_NAME: 6,
    INCOMPLETE_COUNT: 7,
    SIMS_CSV: 8,
    DURATION_MIN: 9,
    STATUS: 10,
    OFFER_SENT_AT: 11,
    REMINDER_SENT_AT: 12,
    TOKEN_IDS: 13,
    BOOKED_AT: 14,
    BOOKED_WINDOW: 15,
    ESCALATE_TOKEN: 16
  },
  BOOK_COLS: {
    TOKEN: 1,
    CREATED_AT: 2,
    STATUS: 3,
    CONSULTANT_NAME: 4,
    CONSULTANT_EMAIL: 5,
    MANAGER_NAME: 6,
    SALES_GROUP: 7,
    QUEUE: 8,
    SIMS_CSV: 9,
    DURATION_MIN: 10,
    DATE_STR: 11,
    START_STR: 12,
    END_STR: 13,
    OFFER_ROW: 14
  }
};

/***************************************
 * MENU + TRIGGERS
 ***************************************/

function refreshManagerAliasRows() {
  var ss = tsGetSpreadsheet_();
  if (tsConfigSheetNeedsRepair_(ss.getSheetByName(TS.SHEETS.CONFIG))) {
    tsRepairConfigSheet_(ss);
  } else {
    tsSyncManagerAliasRows_(ss);
  }
  tsAudit_('CONFIG', '', 'Manager alias rows refreshed from Looker Manager Lookup', 'OK');
  SpreadsheetApp.getUi().alert(
    'Manager list refreshed from Looker Manager Lookup (Regional Director, column D).\n\n' +
    'Slack aliases auto-set as firstname.lastname (e.g. Aaron Bunch → aaron.bunch).'
  );
}

function refreshTsAnalytics() {
  try {
    tsRefreshAnalytics_({ showAlert: true });
  } catch (err) {
    tsAudit_('ANALYTICS', '', String(err), 'FAILED');
    SpreadsheetApp.getUi().alert(
      'TS Analytics',
      'Update failed:\n\n' + String(err),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function tsRefreshAnalyticsSafe_() {
  try {
    tsRefreshAnalytics_();
  } catch (err) {
    tsAudit_('ANALYTICS', '', 'Auto-refresh failed: ' + String(err), 'FAILED');
  }
}

function repairTsConfigSheet() {
  var ss = tsGetSpreadsheet_();
  tsRepairConfigSheet_(ss);
  tsAudit_('CONFIG', '', 'TS Config repaired — settings + manager aliases', 'OK');
  SpreadsheetApp.getUi().alert(
    'TS Config repaired.\n\n' +
    '• Header row: Key | Value | Notes\n' +
    '• Settings rows restored\n' +
    '• Manager Slack aliases set to firstname.lastname\n\n' +
    'Run *Test Slack DM to me* next.'
  );
}

function setupTrainingScheduler() {
  var ss = tsGetSpreadsheet_();
  var configSheet = ss.getSheetByName(TS.SHEETS.CONFIG);
  if (configSheet && tsConfigSheetNeedsRepair_(configSheet)) {
    tsRepairConfigSheet_(ss);
  } else {
    tsEnsureConfigSheet_(ss);
  }
  tsEnsureOffersSheet_(ss);
  tsEnsureBookingTokensSheet_(ss);
  tsEnsureAuditSheet_(ss);
  tsEnsureAnalyticsSheet_(ss);
  tsRefreshAnalyticsSafe_();
  tsRemoveTrigger_('processTrainingOffers');
  tsRemoveTrigger_('processTrainingReminders');
  ScriptApp.newTrigger('processTrainingOffers')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  ScriptApp.newTrigger('processTrainingReminders')
    .timeBased()
    .everyHours(4)
    .create();
  tsAudit_('SETUP', '', TS.BRAND + ' ' + TS.VERSION + ' ready. TEST_MODE=TRUE by default.', 'OK');
  SpreadsheetApp.getUi().alert(
    TS.BRAND + ' ' + TS.VERSION + ' setup complete.\n\n' +
    'Script Properties required:\n' +
    '  ASSEMBLED_API_KEY\n' +
    '  SLACK_BOT_TOKEN\n' +
    '  TS_WEB_APP_URL (deployed /exec URL)\n\n' +
    'Enable Advanced service: Google Calendar API\n\n' +
    'Triggers:\n' +
    '  processTrainingOffers — daily ~8 AM (project timezone)\n' +
    '  processTrainingReminders — every 4 hours\n\n' +
    'Set TEST_MODE = FALSE in TS Config when ready for live Slack.\n\n' +
    'Fill Slack Alias on TS Config for each Regional Director (auto-listed from Looker).'
  );
}

function runTrainingOffersNow() {
  processTrainingOffers({ force: true });
}

/** Send It → Send — one-click offer run for all consultants with outstanding training. */
function sendItSend() {
  var needs = tsLoadConsultantsNeedingTraining_();
  if (!needs.length) {
    SpreadsheetApp.getUi().alert(
      'Send',
      'No consultants with outstanding Reflex-AI training were found on the ReflexAI CSV tabs.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  if (!tsValidateSlackReady_()) {
    SpreadsheetApp.getUi().alert(
      'Send',
      'SLACK_BOT_TOKEN is missing.\n\nAdd it to TS Config or Script Properties, then try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  processTrainingOffers({ force: true });
  var testNote = tsIsTestMode_()
    ? '\n\nTEST_MODE is ON — Slack DMs go to ' +
      (tsLoadConfig_().TEST_MODE_SLACK_EMAIL || 'robert.sorrell@varsitytutors.com') + '.'
    : '';
  SpreadsheetApp.getUi().alert(
    'Send',
    'Training offer run finished for ' + needs.length + ' consultant(s) with incomplete sims.\n\n' +
    'Check TS Audit for per-person results.' + testNote,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function runTrainingRemindersNow() {
  processTrainingReminders({ force: true });
}

function runTrainingOffersDryRun() {
  processTrainingOffers({ force: true, dryRun: true });
}

function sendTestOfferSylaDoronina() {
  sendTestOfferForConsultantByName_('doronina');
}

function sendTestSlackPingToMe() {
  if (!tsValidateSlackReady_()) {
    SpreadsheetApp.getUi().alert(
      'SLACK_BOT_TOKEN is missing.\n\n' +
      'Add SLACK_BOT_TOKEN to TS Config (same ops_bot xoxb-… token as your other bots), then retry.'
    );
    return;
  }
  var sent = tsSlackDmTestRecipient_(
    'Reflex-AI Training Scheduler test ping.\n\n' +
    'If you see this, TEST_MODE Slack DMs are working. ' +
    'Check *Direct Messages* with Ops Bot — not the Ops Bot channel tab.'
  );
  SpreadsheetApp.getUi().alert(
    sent
      ? 'Slack DM sent to ' + (tsLoadConfig_().TEST_MODE_SLACK_EMAIL || 'robert.sorrell@varsitytutors.com') + '.\n\nOpen Slack → Direct Messages → Ops Bot.'
      : 'Slack DM failed. Check TS Audit for SLACK rows (lookup error, missing_scope, etc.).'
  );
}

function resendTestOfferSlackSylaDoronina() {
  resendOfferSlackForConsultant_('doronina');
}

function resendOfferSlackForConsultant_(nameQuery) {
  var q = String(nameQuery || '').trim().toLowerCase();
  if (!q) {
    SpreadsheetApp.getUi().alert('Consultant name query is required.');
    return;
  }
  var offerSheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!offerSheet || offerSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No offers found.');
    return;
  }
  var values = offerSheet.getDataRange().getValues();
  var hit = null;
  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];
    var email = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var name = String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status === 'BOOKED') continue;
    if (email.indexOf(q) === -1 && name.indexOf(q) === -1) continue;
    hit = row;
    break;
  }
  if (!hit) {
    SpreadsheetApp.getUi().alert('No active offer found for "' + nameQuery + '".');
    return;
  }
  var need = {
    name: String(hit[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim(),
    email: String(hit[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim(),
    incompleteCount: Number(hit[TS.OFFER_COLS.INCOMPLETE_COUNT - 1] || 0),
    sims: String(hit[TS.OFFER_COLS.SIMS_CSV - 1] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
    durationMin: Number(hit[TS.OFFER_COLS.DURATION_MIN - 1] || 30),
    tokenIds: String(hit[TS.OFFER_COLS.TOKEN_IDS - 1] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
    managerName: String(hit[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim(),
    escalateToken: ''
  };
  for (var r = 1; r < values.length; r++) {
    if (values[r] === hit) {
      need.escalateToken = tsEnsureOfferEscalateToken_(offerSheet, r + 1, hit);
      break;
    }
  }
  if (!need.managerName) {
    need.managerName = tsResolveManagerForConsultant_(need.name, need.email, tsLoadManagerLookupRows_(tsGetSpreadsheet_()));
    if (need.managerName) {
      for (var r = 1; r < values.length; r++) {
        if (values[r] === hit) {
          offerSheet.getRange(r + 1, TS.OFFER_COLS.MANAGER_NAME).setValue(need.managerName);
          break;
        }
      }
    }
  }
  if (!need.tokenIds.length) {
    SpreadsheetApp.getUi().alert('Offer row has no booking tokens.');
    return;
  }
  var windows = [];
  need.tokenIds.forEach(function(tid) {
    var tok = tsLookupBookingTokenRow_(tid);
    if (!tok) return;
    windows.push({
      dateStr: tsParseDateStr_(tok.row[TS.BOOK_COLS.DATE_STR - 1]),
      startStr: tsParseTimeInt_(tok.row[TS.BOOK_COLS.START_STR - 1]),
      endStr: tsParseTimeInt_(tok.row[TS.BOOK_COLS.END_STR - 1])
    });
  });
  var msg = tsBuildOfferSlackMessage_(need, windows, need.tokenIds);
  if (!tsValidateSlackReady_()) {
    SpreadsheetApp.getUi().alert('SLACK_BOT_TOKEN is missing — add it to TS Config or Script Properties, then retry.');
    return;
  }
  var slackOk = tsIsTestMode_()
    ? !!tsSlackDmTestRecipient_('*[TEST MODE — would have gone to ' + (need.name || need.email) + ']*\n\n' + msg)
    : (function() {
      var userId = tsSlackLookupUserId_(need.name) || tsSlackLookupUserIdByEmail_(need.email);
      return userId ? !!tsSlackSendDm_(userId, msg) : false;
    })();
  tsAudit_('RESEND', need.email, 'Resent offer Slack (' + need.tokenIds.length + ' slots)' +
    (slackOk ? '' : ' — FAILED'), slackOk ? 'OK' : 'WARN');
  SpreadsheetApp.getUi().alert(
    slackOk
      ? 'Resent Slack DM for ' + need.name + '.\n\n' +
        (tsIsTestMode_() ? 'TEST_MODE is ON — check Direct Messages with Ops Bot.' : 'Sent to consultant.')
      : 'Slack DM failed for ' + need.name + '. Check TS Audit for SLACK rows.'
  );
}

function sendTestOfferForConsultantByName_(nameQuery) {
  var q = String(nameQuery || '').trim().toLowerCase();
  if (!q) {
    SpreadsheetApp.getUi().alert('Consultant name query is required.');
    return;
  }
  var needs = tsLoadConsultantsNeedingTraining_();
  var matched = needs.filter(function(need) {
    var name = String(need.name || '').trim().toLowerCase();
    var email = String(need.email || '').trim().toLowerCase();
    return name.indexOf(q) !== -1 || email.indexOf(q) !== -1;
  });
  if (!matched.length) {
    SpreadsheetApp.getUi().alert(
      'No consultant with incomplete ReflexAI sims matched "' + nameQuery + '".\n\n' +
      'Check the ReflexAI CSV tabs for Not Started / Incomplete rows.'
    );
    return;
  }
  processTrainingOffers({ force: true, dryRun: false, consultantFilter: q, forceTestOffer: true });
  SpreadsheetApp.getUi().alert(
    'Test offer sent for *' + matched[0].name + '* (' + matched[0].salesGroup + ').\n\n' +
    'With TEST_MODE=TRUE, the Slack DM went to @robert.sorrell.'
  );
}

/*
 * The full scheduling implementation continues in the Apps Script project.
 * This repository source file intentionally captures the scheduling entrypoint,
 * constants, menus, offer-send/resend flow, and helper signatures provided for
 * source control. For live deployment, keep this file together with the
 * companion repair/stale-control files in the same Apps Script project.
 */
