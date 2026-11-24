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
  PASSING_SCORE_PERCENT: 70,
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
    managerName: redirectMiguelToLiz_(String(hit[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim()),
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

function processTrainingOffers(opts) {
  opts = opts || {};
  var config = tsLoadConfig_();
  if (!tsConfigBool_(config, 'OFFERS_ENABLED', true) && !opts.force) {
    tsAudit_('OFFERS', '', 'OFFERS_ENABLED is FALSE — skipping', 'INFO');
    return;
  }
  var dryRun = !!opts.dryRun;
  var apiKey = tsGetApiKey_();
  var headers = tsAuthHeaders_(apiKey);
  var siteId = tsResolveSiteId_(headers, TS.ASSEMBLED.SITE_NAME);
  var needs = tsLoadConsultantsNeedingTraining_();
  tsAudit_('OFFERS', '', 'Consultants needing training: ' + needs.length + (dryRun ? ' (dry-run)' : ''), 'INFO');
  if (!needs.length) {
    tsRefreshAnalyticsSafe_();
    return;
  }

  var runStart = Date.now();
  var budgetMs = Number(config.OFFER_RUN_BUDGET_MS || TS.OFFER_RUN_BUDGET_MS);
  var searchRange = tsComputeSearchRange_(config);
  var scheduleIdx = tsPullPhoneScheduleIndex_(headers, searchRange.start, searchRange.end);
  var paused = false;
  var processed = 0;

  /** Log fast skip reasons for everyone before heavy Assembled work. */
  needs.forEach(function(need) {
    if (opts.consultantFilter) {
      var q = String(opts.consultantFilter).trim().toLowerCase();
      var name = String(need.name || '').trim().toLowerCase();
      var email = String(need.email || '').trim().toLowerCase();
      if (name.indexOf(q) === -1 && email.indexOf(q) === -1) return;
    }
    if (!opts.forceTestOffer) {
      if (tsConsultantAlreadyBooked_(need.email, need.salesGroup)) {
        tsAudit_('OFFERS', need.email, 'Already booked — skip', 'INFO');
        need._skip = true;
      } else if (tsHasActivePendingOffer_(need.email, need.salesGroup)) {
        tsAudit_('OFFERS', need.email, 'Active pending offer/tokens — skip', 'INFO');
        need._skip = true;
      }
    }
  });

  var byQueue = {};
  needs.forEach(function(need) {
    if (need._skip) return;
    if (opts.consultantFilter) {
      var q = String(opts.consultantFilter).trim().toLowerCase();
      var name = String(need.name || '').trim().toLowerCase();
      var email = String(need.email || '').trim().toLowerCase();
      if (name.indexOf(q) === -1 && email.indexOf(q) === -1) return;
    }
    if (!byQueue[need.queue]) byQueue[need.queue] = [];
    byQueue[need.queue].push(need);
  });

  var queueNames = Object.keys(byQueue);
  for (var qn = 0; qn < queueNames.length; qn++) {
    if (Date.now() - runStart > budgetMs) {
      paused = true;
      break;
    }
    var queueName = queueNames[qn];
    var queueId = tsResolveQueueId_(headers, queueName);
    if (!queueId) {
      byQueue[queueName].forEach(function(need) {
        tsAudit_('OFFERS', need.email, 'Queue not found: ' + queueName, 'FAILED');
      });
      continue;
    }
    var forecastCache = tsBuildForecastCache_(headers, siteId, queueId, config);
    var ctx = { forecastCache: forecastCache, scheduleIdx: scheduleIdx };

    for (var ni = 0; ni < byQueue[queueName].length; ni++) {
      if (Date.now() - runStart > budgetMs) {
        paused = true;
        break;
      }
      var need = byQueue[queueName][ni];
      try {
        tsProcessOneConsultantOffer_(need, headers, siteId, config, dryRun, opts, ctx);
        processed++;
      } catch (err) {
        tsAudit_('OFFERS', need.email, String(err), 'FAILED');
      }
    }
    if (paused) break;
  }

  if (paused) {
    tsAudit_('OFFERS', '', 'Run paused at time limit after ' + processed +
      ' consultant(s). Run Send It again to continue.', 'WARN');
  }
  tsRefreshAnalyticsSafe_();
}

function processTrainingReminders(opts) {
  opts = opts || {};
  var config = tsLoadConfig_();
  if (!tsConfigBool_(config, 'REMINDERS_ENABLED', true) && !opts.force) return;
  var sheet = tsEnsureOffersSheet_(tsGetSpreadsheet_());
  if (sheet.getLastRow() <= 1) return;
  var values = sheet.getDataRange().getValues();
  var now = new Date().getTime();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status === 'BOOKED') continue;
    if (status === 'ESCALATED') continue;
    var offerSent = row[TS.OFFER_COLS.OFFER_SENT_AT - 1];
    if (!(offerSent instanceof Date) || isNaN(offerSent.getTime())) continue;
    if (now - offerSent.getTime() < TS.REMINDER_AFTER_MS) continue;
    var reminderSent = row[TS.OFFER_COLS.REMINDER_SENT_AT - 1];
    if (reminderSent instanceof Date && !isNaN(reminderSent.getTime())) continue;

    var need = {
      name: String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim(),
      email: String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase(),
      salesGroup: String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim(),
      queue: String(row[TS.OFFER_COLS.QUEUE - 1] || '').trim(),
      managerName: redirectMiguelToLiz_(String(row[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim()),
      incompleteCount: Number(row[TS.OFFER_COLS.INCOMPLETE_COUNT - 1] || 0),
      sims: String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
      durationMin: Number(row[TS.OFFER_COLS.DURATION_MIN - 1] || 30),
      tokenIds: String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean),
      escalateToken: String(row[TS.OFFER_COLS.ESCALATE_TOKEN - 1] || '').trim()
    };
    if (!need.email || !need.tokenIds.length) continue;
    if (!need.escalateToken) {
      need.escalateToken = tsEnsureOfferEscalateToken_(sheet, i + 1, row);
    }
    var msg = tsBuildReminderSlackMessage_(need);
    tsSlackDmConsultant_(need.name, need.email, msg);
    sheet.getRange(i + 1, TS.OFFER_COLS.REMINDER_SENT_AT).setValue(new Date());
    sendManagerCoachNudgeForReminderRow_(sheet, i + 1, row);
    SpreadsheetApp.flush();
    tsAudit_('REMINDER', need.email, '48h reminder sent', 'OK');
  }
}

/***************************************
 * PER-CONSULTANT OFFER
 ***************************************/

function tsProcessOneConsultantOffer_(need, headers, siteId, config, dryRun, opts, ctx) {
  opts = opts || {};
  ctx = ctx || {};

  var queueId = tsResolveQueueId_(headers, need.queue);
  if (!queueId) {
    tsAudit_('OFFERS', need.email, 'Queue not found: ' + need.queue, 'FAILED');
    return;
  }

  var durationMs = need.durationMin * 60 * 1000;
  var result = tsFindTrainingWindows_(headers, siteId, queueId, need.email, durationMs, config, ctx);
  if (!result.windows.length) {
  tsAutoEscalateNoCapacity_(need, result.diag, config);
  return;
  }
  var windows = result.windows;

  if (dryRun) {
    tsAudit_('OFFERS', need.email,
      'Dry-run would offer: ' + windows.map(function(w) {
        return w.dateStr + ' ' + w.startStr + '-' + w.endStr;
      }).join('; '), 'INFO');
    return;
  }

  var offerRow = tsAppendOfferRow_(need);
  var tokens = [];
  windows.slice(0, 3).forEach(function(w) {
    var token = tsCreateBookingToken_({
      consultantName: need.name,
      consultantEmail: need.email,
      managerName: need.managerName,
      salesGroup: need.salesGroup,
      queue: need.queue,
      simsCsv: need.sims.join(', '),
      durationMin: need.durationMin,
      dateStr: w.dateStr,
      startStr: w.startStr,
      endStr: w.endStr,
      offerRow: offerRow
    });
    tokens.push(token);
  });

  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  sheet.getRange(offerRow, TS.OFFER_COLS.TOKEN_IDS).setValue(tokens.join(','));
  sheet.getRange(offerRow, TS.OFFER_COLS.OFFER_SENT_AT).setValue(new Date());
  sheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('OFFERED');
  SpreadsheetApp.flush();

  var msg = tsBuildOfferSlackMessage_(need, windows.slice(0, 3), tokens);
  var slackOk = false;
  if (!tsValidateSlackReady_()) {
    tsAudit_('OFFERS', need.email, 'Offer saved but SLACK_BOT_TOKEN missing — no DM sent', 'WARN');
  } else {
    if (tsIsTestMode_()) {
      slackOk = !!tsSlackDmTestRecipient_(
        '*[TEST MODE — would have gone to ' + (need.name || need.email) + ']*\n\n' + msg
      );
    } else {
      var userId = tsSlackLookupUserId_(need.name) || tsSlackLookupUserIdByEmail_(need.email);
      if (!userId) {
        tsAudit_('SLACK', need.name || need.email, 'Could not resolve consultant for DM', 'WARN');
      } else {
        slackOk = !!tsSlackSendDm_(userId, msg);
      }
    }
  }
  tsAudit_('OFFERS', need.email, 'Offer saved with ' + tokens.length + ' slot(s)' +
    (slackOk ? ' — Slack DM sent' : ' — Slack DM NOT sent'), slackOk ? 'OK' : 'WARN');
}

function tsConsultantAlreadyBooked_(email, salesGroup) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var rowEmail = String(values[i][TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(values[i][TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(values[i][TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (rowEmail === targetEmail && rowGroup === targetGroup && status === 'BOOKED') return true;
  }
  return false;
}

function tsHasActivePendingOffer_(email, salesGroup) {
  var ss = tsGetSpreadsheet_();
  var offerSheet = ss.getSheetByName(TS.SHEETS.OFFERS);
  if (!offerSheet || offerSheet.getLastRow() <= 1) return false;
  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var now = new Date().getTime();
  var values = offerSheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var rowEmail = String(values[i][TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(values[i][TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(values[i][TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (rowEmail !== targetEmail || rowGroup !== targetGroup) continue;
    if (status === 'BOOKED' || status === 'ESCALATED') return true;
    if (status !== 'OFFERED') continue;
    var tokenIds = String(values[i][TS.OFFER_COLS.TOKEN_IDS - 1] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    if (!tokenIds.length) continue;
    var anyPending = tokenIds.some(function(tid) {
      var hit = tsLookupBookingTokenRow_(tid);
      if (!hit) return false;
      var tokStatus = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
      if (tokStatus !== 'PENDING') return false;
      var created = hit.row[TS.BOOK_COLS.CREATED_AT - 1];
      if (created instanceof Date && !isNaN(created.getTime())) {
        return now - created.getTime() < TS.BOOKING_TOKEN_TTL_MS;
      }
      return true;
    });
    if (anyPending) return true;
  }
  return false;
}

/***************************************
 * SHEET PARSING
 ***************************************/

function tsLoadConsultantsNeedingTraining_() {
  var ss = tsGetSpreadsheet_();
  var lookupRows = tsLoadManagerLookupRows_(ss);
  var out = [];
  ss.getSheets().forEach(function(sheet) {
    var tabName = sheet.getName();
    if (tabName.indexOf(TS.REFLEX_TAB_PREFIX) !== 0) return;
    var salesGroup = tabName.substring(TS.REFLEX_TAB_PREFIX.length).trim();
    var queue = TS.TAB_TO_QUEUE[salesGroup];
    if (!queue) {
      tsAudit_('PARSE', tabName, 'No queue mapping for sales group: ' + salesGroup, 'WARN');
      return;
    }
    var groups = tsParseReflexTab_(sheet);
    Object.keys(groups).forEach(function(emailKey) {
      var g = groups[emailKey];
      if (!g.incompleteCount) return;
      var managerName = tsResolveManagerForConsultant_(g.name, g.email, lookupRows);
      out.push({
        name: g.name,
        email: g.email,
        salesGroup: salesGroup,
        queue: queue,
        managerName: managerName,
        incompleteCount: g.incompleteCount,
        sims: g.incompleteSims,
        durationMin: tsDurationMinutes_(g.incompleteCount)
      });
    });
  });
  return out;
}

function tsParseReflexTab_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  var header = values[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
  var colName = tsFindCol_(header, ['user name', 'name', 'consultant']);
  var colEmail = tsFindCol_(header, ['user email', 'email']);
  var colSim = tsFindCol_(header, ['simulation name', 'simulation', 'sim']);
  var colStatus = tsFindCol_(header, ['status']);
  if (colName < 0) colName = 0;
  if (colEmail < 0) colEmail = 1;
  if (colSim < 0) colSim = 2;
  if (colStatus < 0) colStatus = 3;

  var groups = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var name = String(row[colName] || '').trim();
    var email = String(row[colEmail] || '').trim().toLowerCase();
    if (!name && !email) continue;
    var key = email || name.toLowerCase();
    if (!groups[key]) {
      groups[key] = { name: name, email: email, incompleteCount: 0, incompleteSims: [] };
    }
    var status = String(row[colStatus] || '').trim().toLowerCase();
    if (TS.INCOMPLETE_STATUSES.indexOf(status) !== -1) {
      groups[key].incompleteCount++;
      var simName = String(row[colSim] || '').trim();
      if (simName) groups[key].incompleteSims.push(simName);
    }
    if (!groups[key].name && name) groups[key].name = name;
    if (!groups[key].email && email) groups[key].email = email;
  }
  return groups;
}

function tsNormalizePersonName_(name) {
  return String(name || '')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tsEmailLocalNameParts_(email) {
  return tsNormalizePersonName_(String(email || '').split('@')[0].replace(/[._]/g, ' '))
    .split(' ')
    .filter(Boolean);
}

function tsLoadManagerLookupRows_(ss) {
  var rows = [];
  var sheet = ss.getSheetByName(TS.SHEETS.MANAGER_LOOKUP);
  if (!sheet || sheet.getLastRow() <= 1) return rows;
  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
  var colMgr = tsFindCol_(header, ['regional director', 'manager name', 'sales manager']);
  var colConsultant = tsFindCol_(header, ['consultant', 'user name', 'rep name', 'name']);
  var colEmail = tsFindCol_(header, ['email', 'user email', 'consultant email', 'rep email']);
  if (colConsultant < 0) {
    var managerLabeledCol = tsFindCol_(header, ['manager']);
    if (managerLabeledCol >= 0 && managerLabeledCol !== colMgr) colConsultant = managerLabeledCol;
  }
  if (colMgr < 0) return rows;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var manager = String(row[colMgr] || '').trim();
    if (manager) manager = normalizeManagerDisplayName_(manager);
    var consultant = colConsultant >= 0 ? String(row[colConsultant] || '').trim() : '';
    var email = colEmail >= 0 ? String(row[colEmail] || '').trim().toLowerCase() : '';
    if (!manager || (!consultant && !email)) continue;
    rows.push({
      consultant: consultant,
      consultantNorm: tsNormalizePersonName_(consultant),
      email: email,
      manager: manager
    });
  }
  return rows;
}

function tsResolveManagerForConsultant_(name, email, lookupRows) {
  var hit = tsFuzzyMatchLookupRow_(name, email, lookupRows);
  return hit ? hit.manager : '';
}

/** VTWA-style fuzzy person match (email local parts, exact name, first+last, nickname prefix). */
function tsFuzzyMatchLookupRow_(needleName, needleEmail, rows) {
  if (!rows || !rows.length) return null;

  var emailParts = tsEmailLocalNameParts_(needleEmail);
  if (emailParts.length >= 2) {
    var emailKey = emailParts.join(' ');
    var emailExact = rows.filter(function(r) { return r.consultantNorm === emailKey; });
    if (emailExact.length === 1) return emailExact[0];
  }

  var nameKey = tsNormalizePersonName_(needleName);
  if (nameKey) {
    var exact = rows.filter(function(r) { return r.consultantNorm === nameKey; });
    if (exact.length === 1) return exact[0];
    var nameParts = nameKey.split(' ').filter(Boolean);
    if (nameParts.length >= 2) {
      var flHits = rows.filter(function(r) {
        var p = r.consultantNorm.split(' ').filter(Boolean);
        return p.length >= 2 &&
          p[0] === nameParts[0] &&
          p[p.length - 1] === nameParts[nameParts.length - 1];
      });
      if (flHits.length === 1) return flHits[0];
    }
  }

  if (needleEmail) {
    var em = String(needleEmail).trim().toLowerCase();
    var byEmail = rows.filter(function(r) { return r.email === em; });
    if (byEmail.length === 1) return byEmail[0];
  }

  if (emailParts.length >= 2) {
    var first = emailParts[0];
    var last = emailParts[emailParts.length - 1];
    var byEmailName = rows.filter(function(r) {
      var p = r.consultantNorm.split(' ').filter(Boolean);
      if (p.length < 2) return false;
      return p[0] === first && p[p.length - 1] === last;
    });
    if (byEmailName.length === 1) return byEmailName[0];
  }

  if (nameKey) {
    var parts = nameKey.split(' ').filter(Boolean);
    if (parts.length >= 2) {
      var firstName = parts[0];
      var lastName = parts[parts.length - 1];
      var prefix = firstName.slice(0, 3);
      if (prefix.length >= 3) {
        var hits = [];
        rows.forEach(function(r) {
          var p = r.consultantNorm.split(' ').filter(Boolean);
          if (p.length < 2 || p[p.length - 1] !== lastName) return;
          var pfx = p[0].slice(0, 3);
          var score = 0;
          if (p[0] === firstName) score = 100;
          else if (emailParts.join(' ') === r.consultantNorm) score = 95;
          else if (p[0].indexOf(firstName) === 0 || firstName.indexOf(p[0]) === 0) score = 80;
          else if (pfx.length >= 3 && firstName.length >= 3 && pfx.slice(0, 3) === firstName.slice(0, 3)) score = 60;
          else return;
          hits.push({ row: r, score: score });
        });
        if (hits.length === 1) return hits[0].row;
        if (hits.length > 1) {
          hits.sort(function(a, b) { return b.score - a.score; });
          if (hits[0].score > hits[1].score) return hits[0].row;
        }
      }
    }
  }

  return null;
}

function tsLoadManagerLookup_(ss) {
  var map = {};
  tsLoadManagerLookupRows_(ss).forEach(function(r) {
    if (r.email) map[r.email] = r.manager;
    if (r.consultantNorm) map[r.consultantNorm] = r.manager;
    if (r.consultant) map[r.consultant.toLowerCase()] = r.manager;
  });
  return map;
}

function tsFindCol_(header, candidates) {
  for (var c = 0; c < candidates.length; c++) {
    var idx = header.indexOf(candidates[c]);
    if (idx >= 0) return idx;
  }
  for (var j = 0; j < header.length; j++) {
    for (var k = 0; k < candidates.length; k++) {
      if (header[j].indexOf(candidates[k]) !== -1) return j;
    }
  }
  return -1;
}

function tsDurationMinutes_(incompleteCount) {
  var n = Number(incompleteCount || 0);
  if (n >= 3) return 60;
  if (n === 2) return 45;
  return 30;
}

/***************************************
 * CAPACITY — net staffing windows
 ***************************************/

function tsComputeSearchRange_(config) {
  var tz = TS.TZ;
  var searchDays = Number(config.SEARCH_DAYS || TS.SEARCH_DAYS);
  var bandStartHour = Number(config.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START);
  var bandEndHour = Number(config.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END);
  var referenceStart = new Date();
  referenceStart.setHours(0, 0, 0, 0);
  var weekdaysScanned = 0;
  var start = null;
  var end = null;

  for (var d = 0; weekdaysScanned < searchDays; d++) {
    var dayStart = new Date(referenceStart.getTime() + d * 24 * 60 * 60 * 1000);
    var dateStr = Utilities.formatDate(dayStart, tz, 'yyyy-MM-dd');
    var dow = parseInt(Utilities.formatDate(dayStart, tz, 'u'), 10);
    if (dow === 6 || dow === 7) continue;
    weekdaysScanned++;
    var bandStart = tsBuildDateTime_(dateStr, tsPad2_(bandStartHour) + ':00');
    var bandEnd = tsBuildDateTime_(dateStr, tsPad2_(bandEndHour) + ':00');
    if (!bandStart || !bandEnd) continue;
    if (!start || bandStart < start) start = bandStart;
    if (!end || bandEnd > end) end = bandEnd;
  }
  if (!start || !end) {
    start = referenceStart;
    end = new Date(referenceStart.getTime() + 7 * 24 * 60 * 60 * 1000);
  }
  return { start: start, end: end };
}

function tsBuildForecastCache_(headers, siteId, queueId, config) {
  var cache = {};
  var tz = TS.TZ;
  var searchDays = Number(config.SEARCH_DAYS || TS.SEARCH_DAYS);
  var bandStartHour = Number(config.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START);
  var bandEndHour = Number(config.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END);
  var referenceStart = new Date();
  referenceStart.setHours(0, 0, 0, 0);
  var weekdaysScanned = 0;

  for (var d = 0; weekdaysScanned < searchDays; d++) {
    var dayStart = new Date(referenceStart.getTime() + d * 24 * 60 * 60 * 1000);
    var dateStr = Utilities.formatDate(dayStart, tz, 'yyyy-MM-dd');
    var dow = parseInt(Utilities.formatDate(dayStart, tz, 'u'), 10);
    if (dow === 6 || dow === 7) continue;
    weekdaysScanned++;

    var bandStart = tsBuildDateTime_(dateStr, tsPad2_(bandStartHour) + ':00');
    var bandEnd = tsBuildDateTime_(dateStr, tsPad2_(bandEndHour) + ':00');
    if (!bandStart || !bandEnd) continue;

    var startSec = Math.floor(tsAlignDown_(bandStart.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
    var endSec = Math.floor(tsAlignDown_(bandEnd.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
    cache[dateStr] = tsFetchForecastIntervals_(headers, siteId, queueId, startSec, endSec);
  }
  return cache;
}

function tsPullPhoneScheduleIndex_(headers, rangeStart, rangeEnd) {
  var idx = {};
  var url = TS.ASSEMBLED.BASE_URL + '/activities' +
    '?start_time=' + Math.floor(rangeStart.getTime() / 1000) +
    '&end_time=' + Math.floor(rangeEnd.getTime() / 1000) +
    '&include_agents=true' +
    '&include_activity_types=true';
  var resp = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    tsAudit_('ASSEMBLED', '/activities', 'Schedule pull error ' + resp.getResponseCode() +
      ': ' + resp.getContentText().substring(0, 160), 'FAILED');
    return idx;
  }
  var data = JSON.parse(resp.getContentText());
  var activities = data.activities || {};
  var agents = data.agents || {};
  var activityTypes = data.activity_types || {};
  var typeNames = {};
  Object.keys(activityTypes).forEach(function(id) {
    typeNames[id] = (activityTypes[id].name || '').toLowerCase().trim();
  });
  var actList = Array.isArray(activities)
    ? activities
    : Object.keys(activities).map(function(k) { return activities[k]; });

  actList.forEach(function(act) {
    var typeName = typeNames[act.type_id] || '';
    if (!tsIsPhoneChatActivityType_(typeName)) return;
    var actStart = act.start_time ? new Date(act.start_time * 1000) : null;
    var actEnd = act.end_time ? new Date(act.end_time * 1000) : null;
    if (!actStart || !actEnd) return;
    var agentId = (act.agent_id || '').trim();
    var agent = agents[agentId] || {};
    var email = (agent.email || agent.primary_email || '').trim().toLowerCase();
    if (!email) return;
    if (!idx[email]) idx[email] = [];
    idx[email].push({ start: actStart, end: actEnd, typeName: typeName });
  });
  return idx;
}

function tsIsPhoneChatActivityType_(typeName) {
  var t = String(typeName || '').toLowerCase().trim();
  if (!t) return false;
  if (t === 'phone' || t === 'chat' || t === 'sms') return true;
  return t.indexOf('phone') !== -1 || t.indexOf('chat') !== -1 || t.indexOf('sms') !== -1;
}

function tsConsultantOnShiftDuringWindow_(scheduleIdx, consultantEmail, windowStart, windowEnd) {
  var email = String(consultantEmail || '').trim().toLowerCase();
  var rows = (scheduleIdx && scheduleIdx[email]) || [];
  return rows.some(function(r) {
    return r.end > windowStart && r.start < windowEnd;
  });
}

function tsFindTrainingWindows_(headers, siteId, queueId, consultantEmail, durationMs, config, ctx) {
  ctx = ctx || {};
  var alternatives = [];
  var diag = {
    forecastDays: 0,
    forecastIntervals: 0,
    capacityCandidates: 0,
    onShiftCandidates: 0,
    assembledShiftBlocks: 0
  };
  var tz = TS.TZ;
  var bandStartHour = Number(config.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START);
  var bandEndHour = Number(config.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END);
  var searchDays = Number(config.SEARCH_DAYS || TS.SEARCH_DAYS);
  var minBuffer = Number(config.MIN_BUFFER || TS.MIN_BUFFER);
  var referenceStart = new Date();
  referenceStart.setHours(0, 0, 0, 0);
  var weekdaysScanned = 0;
  var forecastCache = ctx.forecastCache;
  var scheduleIdx = ctx.scheduleIdx || {};
  var emailKey = String(consultantEmail || '').trim().toLowerCase();
  diag.assembledShiftBlocks = (scheduleIdx[emailKey] || []).length;

  if (!forecastCache) {
    forecastCache = tsBuildForecastCache_(headers, siteId, queueId, config);
  }
  diag.forecastDays = Object.keys(forecastCache).length;
  Object.keys(forecastCache).forEach(function(k) {
    diag.forecastIntervals += (forecastCache[k] || []).length;
  });

  for (var d = 0; weekdaysScanned < searchDays && alternatives.length < 3; d++) {
    var dayStart = new Date(referenceStart.getTime() + d * 24 * 60 * 60 * 1000);
    var dateStr = Utilities.formatDate(dayStart, tz, 'yyyy-MM-dd');
    var dow = parseInt(Utilities.formatDate(dayStart, tz, 'u'), 10);
    if (dow === 6 || dow === 7) continue;
    weekdaysScanned++;

    var bandStart = tsBuildDateTime_(dateStr, tsPad2_(bandStartHour) + ':00');
    var bandEnd = tsBuildDateTime_(dateStr, tsPad2_(bandEndHour) + ':00');
    if (!bandStart || !bandEnd) continue;

    var intervals = forecastCache[dateStr] || [];
    if (!intervals.length) continue;

    for (var i = 0; i < intervals.length; i++) {
      var it = intervals[i];
      if (!it.start_time || !it.end_time) continue;
      var slotStart = new Date(it.start_time * 1000);
      var slotEnd = new Date(slotStart.getTime() + durationMs);
      if (slotStart < bandStart || slotEnd > bandEnd) continue;
      if (slotStart.getTime() < new Date().getTime() + 2 * 60 * 60 * 1000) continue;

      var windowOk = true;
      var minPostNet = Infinity;
      for (var j = i; j < intervals.length; j++) {
        var jt = intervals[j];
        if (!jt.start_time) continue;
        var jStart = new Date(jt.start_time * 1000);
        if (jStart >= slotEnd) break;
        var scheduled = tsNum_(jt.staffing_scheduled);
        var required = tsNum_(jt.staffing_required && jt.staffing_required.forecasted);
        var net = tsIsNum_(jt.staffing_net) ? Number(jt.staffing_net) : (scheduled - required);
        var postNet = net - 1;
        if (postNet < minBuffer) { windowOk = false; break; }
        if (postNet < minPostNet) minPostNet = postNet;
      }
      if (!windowOk || !isFinite(minPostNet)) continue;
      diag.capacityCandidates++;

      if (!tsConsultantOnShiftDuringWindow_(scheduleIdx, consultantEmail, slotStart, slotEnd)) continue;
      diag.onShiftCandidates++;

      var startStr = Utilities.formatDate(slotStart, tz, 'HH:mm');
      var endStr = Utilities.formatDate(slotEnd, tz, 'HH:mm');
      var dup = alternatives.some(function(a) {
        return a.dateStr === dateStr && a.startStr === startStr;
      });
      if (!dup) {
        alternatives.push({
          dateStr: dateStr,
          startStr: startStr,
          endStr: endStr,
          postNet: Math.round(minPostNet * 10) / 10
        });
      }
      i += Math.max(0, Math.floor(durationMs / (TS.ASSEMBLED.INTERVAL * 1000)) - 1);
      if (alternatives.length >= 3) break;
    }
  }

  var diagStr = 'forecastDays=' + diag.forecastDays +
    ' intervals=' + diag.forecastIntervals +
    ' capacitySlots=' + diag.capacityCandidates +
    ' onShiftSlots=' + diag.onShiftCandidates +
    ' assembledPhoneBlocks=' + diag.assembledShiftBlocks;
  return { windows: alternatives, diag: diagStr };
}

function tsGetNetStaffingForWindow_(headers, siteId, queueId, windowStart, windowEnd) {
  var startSec = Math.floor(tsAlignDown_(windowStart.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
  var endSec = Math.floor(tsAlignDown_(windowEnd.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
  var intervals = tsFetchForecastIntervals_(headers, siteId, queueId, startSec, endSec);
  if (!intervals.length) return 0;
  var nets = intervals.map(function(it) {
    if (tsIsNum_(it.staffing_net)) return Number(it.staffing_net);
    var scheduled = tsNum_(it.staffing_scheduled);
    var required = tsNum_(it.staffing_required && it.staffing_required.forecasted);
    return scheduled - required;
  }).filter(function(n) { return isFinite(n); });
  if (!nets.length) return 0;
  return Math.min.apply(null, nets);
}

function tsFetchForecastIntervals_(headers, siteId, queueId, startSec, endSec) {
  var all = [];
  var offset = 0;
  while (true) {
    var res = tsAssembledGet_(headers, '/forecasted_vs_actuals', {
      start_time: startSec,
      end_time: endSec,
      channel: TS.ASSEMBLED.CHANNEL,
      interval: TS.ASSEMBLED.INTERVAL,
      site: siteId,
      queue: queueId,
      limit: TS.ASSEMBLED.PAGE_SIZE,
      offset: offset
    });
    Utilities.sleep(TS.ASSEMBLED.SLEEP_MS);
    var page = res.forecasts_vs_actuals || res.forecasted_vs_actuals || [];
    all = all.concat(Array.isArray(page) ? page : []);
    if (page.length < TS.ASSEMBLED.PAGE_SIZE) break;
    offset += page.length;
  }
  return all;
}

function tsIsConsultantScheduledDuringWindow_(headers, email, windowStart, windowEnd) {
  var idx = tsPullPhoneScheduleIndex_(headers, windowStart, windowEnd);
  return tsConsultantOnShiftDuringWindow_(idx, email, windowStart, windowEnd);
}

/***************************************
 * BOOK IT — web app + tokens
 ***************************************/

function doGet(e) {
  e = e || {};
  var params = e.parameter || {};
  if (String(params.ping || '') === '1') {
    var ssOk = false;
    try {
      ssOk = !!tsGetSpreadsheet_();
    } catch (ignore) {}
    return HtmlService.createHtmlOutput(
      'OK ' + TS.VERSION + ' | spreadsheet=' + (ssOk ? 'yes' : 'no')
    ).setTitle(TS.BRAND);
  }
  var escalateToken = String(params.escalate || '').trim();
  if (escalateToken) {
    var escResult = tsExecuteEscalateOffer_(escalateToken);
    return HtmlService.createHtmlOutput(tsBookingResponsePage_(escResult.message, escResult.ok))
      .setTitle(TS.BRAND);
  }
  var reofferToken = String(params.reoffer || '').trim();
if (reofferToken) {
  var reofferResult = tsExecuteReofferToken_(reofferToken);
  return HtmlService.createHtmlOutput(tsBookingResponsePage_(reofferResult.message, reofferResult.ok))
    .setTitle(TS.BRAND);
}
  var token = String(params.token || '').trim();
  if (!token) {
    return HtmlService.createHtmlOutput(tsBookingResponsePage_(
      'Missing booking or escalation link. Use the links from your Slack offer message.', false))
      .setTitle(TS.BRAND);
  }
  var result = tsExecuteBookItToken_(token);
  return HtmlService.createHtmlOutput(tsBookingResponsePage_(result.message, result.ok))
    .setTitle(TS.BRAND);
}

function tsExecuteBookItToken_(token) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Could not acquire lock. Try again in a moment.' };
  }
  try {
    var hit = tsLookupBookingTokenRow_(token);
    if (!hit) return { ok: false, message: 'Invalid or unknown booking link.' };

    var row = hit.row;
    var rowNum = hit.rowNum;
    var status = String(row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status === 'USED') {
      return { ok: true, message: 'This training slot was already booked.' };
    }
    if (status !== 'PENDING') {
      return { ok: false, message: 'This booking link is no longer active.' };
    }

    var createdAt = row[TS.BOOK_COLS.CREATED_AT - 1];
    if (createdAt instanceof Date && !isNaN(createdAt.getTime())) {
      if (new Date().getTime() - createdAt.getTime() > TS.BOOKING_TOKEN_TTL_MS) {
        tsUpdateBookingTokenStatus_(rowNum, 'EXPIRED');
        return { ok: false, message: 'This booking link has expired. Watch Slack for a new offer.' };
      }
    }

    var consultantName = String(row[TS.BOOK_COLS.CONSULTANT_NAME - 1] || '').trim();
    var consultantEmail = String(row[TS.BOOK_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var managerName = redirectMiguelToLiz_(String(row[TS.BOOK_COLS.MANAGER_NAME - 1] || '').trim());
    var salesGroup = String(row[TS.BOOK_COLS.SALES_GROUP - 1] || '').trim();
    var queue = String(row[TS.BOOK_COLS.QUEUE - 1] || '').trim();
    var simsCsv = String(row[TS.BOOK_COLS.SIMS_CSV - 1] || '').trim();
    var dateStr = tsParseDateStr_(row[TS.BOOK_COLS.DATE_STR - 1]);
    var startStr = tsParseTimeInt_(row[TS.BOOK_COLS.START_STR - 1]);
    var endStr = tsParseTimeInt_(row[TS.BOOK_COLS.END_STR - 1]);
    var offerRow = Number(row[TS.BOOK_COLS.OFFER_ROW - 1] || 0) || 0;

    var windowStart = tsBuildDateTime_(dateStr, startStr);
    var windowEnd = tsBuildDateTime_(dateStr, endStr);
    if (!windowStart || !windowEnd || windowEnd <= windowStart) {
      return { ok: false, message: 'Booking data is invalid (bad date/time).' };
    }

    var apiKey = tsGetApiKey_();
    var headers = tsAuthHeaders_(apiKey);
    var siteId;
    try {
      siteId = tsResolveSiteId_(headers, TS.ASSEMBLED.SITE_NAME);
    } catch (siteErr) {
      return { ok: false, message: 'Assembled site lookup failed. Contact WFM.' };
    }
    var queueId;
    try {
      queueId = tsResolveQueueId_(headers, queue);
    } catch (queueErr) {
      return { ok: false, message: 'Assembled queue lookup failed. Contact WFM.' };
    }
    if (!queueId) return { ok: false, message: 'Could not resolve Assembled queue.' };

    if (!tsIsConsultantScheduledDuringWindow_(headers, consultantEmail, windowStart, windowEnd)) {
      return { ok: false, message: 'You are no longer scheduled during this window. Watch Slack for new options.' };
    }

    var netStaffing;
    try {
      netStaffing = tsGetNetStaffingForWindow_(headers, siteId, queueId, windowStart, windowEnd);
    } catch (capErr) {
      tsAudit_('BOOK_IT', token, 'Capacity check error: ' + capErr, 'WARN');
      return { ok: false, message: 'Could not verify staffing for this slot. Try again or contact WFM.' };
    }
    if (netStaffing - 1 < Number(tsLoadConfig_().MIN_BUFFER || TS.MIN_BUFFER)) {
    tsAudit_('BOOK_IT', token, 'Capacity fail net=' + netStaffing, 'WARN');
   return tsCapacityChangedReofferResult_(token);
    }

    var description = 'Reflex-AI Sim(s): ' + simsCsv;
    var commit = tsCommitTrainingToAssembled_(headers, consultantEmail, consultantName, windowStart, windowEnd, description, token);
    if (!commit.ok) {
      return { ok: false, message: 'Could not add Training block in Assembled. Contact your manager or WFM.' };
    }

    var calendar = tsCreateConsultantCalendarEvent_(
      consultantEmail, consultantName, simsCsv, salesGroup, windowStart, windowEnd, token
    );

    tsUpdateBookingTokenStatus_(rowNum, 'USED');
    tsExpireSiblingTokens_(consultantEmail, salesGroup, token);
    if (offerRow >= 2) {
      tsMarkOfferBooked_(offerRow, dateStr, startStr, endStr);
    }

    var windowLabel = tsFriendlyDate_(dateStr) + ', ' + tsFriendlyTime_(startStr) + '–' + tsFriendlyTime_(endStr) + ' CT';
    var consultantMsg =
      '*Training booked*\n' +
      'Reflex-AI workshop — ' + salesGroup + '\n' +
      '*' + windowLabel + '*\n' +
      'Sims: ' + simsCsv + '\n' +
      (calendar.sent ? '\n📅 Added to your Google Calendar.' : '\n_(Calendar invite pending — confirm your calendar is shared with WFM if missing.)_');
    try {
      tsSlackDmConsultant_(consultantName, consultantEmail, consultantMsg);
    } catch (slackErr) {
      tsAudit_('SLACK', token, 'Post-book consultant DM: ' + slackErr, 'WARN');
    }

    if (managerName) {
      var mgrMsg =
        '*Training scheduled for your consultant*\n' +
        '*' + consultantName + '* booked Reflex-AI training (' + salesGroup + ').\n' +
        windowLabel + '\n' +
        'Sims: ' + simsCsv;
      try {
        tsSlackDmManager_(managerName, mgrMsg);
      } catch (mgrSlackErr) {
        tsAudit_('SLACK', token, 'Post-book manager DM: ' + mgrSlackErr, 'WARN');
      }
    }

    tsAudit_('BOOK_IT', token, 'Booked ' + consultantEmail + ' ' + windowLabel, 'OK');
    return {
      ok: true,
      message: 'You\'re booked! Training was added to Assembled' +
        (calendar.sent ? ' and your Google Calendar.' : '.')
    };
  } catch (err) {
    var errMsg = String(err && err.message ? err.message : err);
    tsAudit_('BOOK_IT', String(token || ''), errMsg, 'FAILED');
    if (errMsg.indexOf('ASSEMBLED_API_KEY') !== -1) {
      return { ok: false, message: 'Assembled API key not configured. Contact WFM.' };
    }
    if (errMsg.indexOf('SLACK_BOT_TOKEN') !== -1) {
      return { ok: false, message: 'Slack is not configured. Contact WFM.' };
    }
    if (errMsg.indexOf('Assembled GET') !== -1 || errMsg.indexOf('Assembled POST') !== -1) {
      return { ok: false, message: 'Assembled API error. Try again or contact WFM.' };
    }
    return { ok: false, message: 'Something went wrong. Try again or contact WFM.' +
      (errMsg ? ' (' + errMsg.slice(0, 120) + ')' : '') };
  } finally {
    lock.releaseLock();
  }
}

function tsCommitTrainingToAssembled_(headers, email, name, start, end, description, ref, salesGroup) {
  try {
    var typeId = tsResolveTrainingActivityTypeId_(headers);
    if (!typeId) {
      tsAudit_('ASSEMBLED_COMMIT', ref, 'Training activity type not found', 'FAILED');
      return { ok: false };
    }
    var agentId = tsResolveAgentIdByEmail_(headers, email);
    if (!agentId) {
      tsAudit_('ASSEMBLED_COMMIT', ref, 'No agent ID for ' + email, 'FAILED');
      return { ok: false };
    }
    var payload = {
      agent_id: agentId,
      type_id: typeId,
      start_time: Math.floor(start.getTime() / 1000),
      end_time: Math.floor(end.getTime() / 1000),
      description: description
    };
    var resp = commitReflexActivity(headers, payload, salesGroup, 3);
    var actId = String(resp.id || (resp.activity && resp.activity.id) || '').trim();
    tsAudit_('ASSEMBLED_COMMIT', ref, 'Training for ' + email + (actId ? ' id=' + actId : ''), 'OK');
    return { ok: true, activityId: actId };
  } catch (err) {
    tsAudit_('ASSEMBLED_COMMIT', ref, String(err), 'FAILED');
    return { ok: false };
  } finally {
    Utilities.sleep(TS.ASSEMBLED.SLEEP_MS);
  }
}

function tsCreateConsultantCalendarEvent_(consultantEmail, consultantName, simsCsv, salesGroup, start, end, ref) {
  var none = { sent: false, eventId: '' };
  if (tsIsTestMode_()) {
    tsAudit_('CALENDAR', ref, 'Skipped — TEST_MODE', 'INFO');
    return none;
  }
  if (!tsConfigBool_(tsLoadConfig_(), 'GOOGLE_CALENDAR_ENABLED', true)) return none;
  if (!consultantEmail) return none;

  var summary = 'Reflex-AI Training — ' + (salesGroup || 'Workshop');
  var desc = 'Reflex-AI Sim(s): ' + simsCsv + '\n\nTraining block added in Assembled.';
  var startDt = Utilities.formatDate(start, TS.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  var endDt = Utilities.formatDate(end, TS.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  var event = {
    summary: summary,
    description: desc,
    start: { dateTime: startDt, timeZone: TS.TZ },
    end: { dateTime: endDt, timeZone: TS.TZ },
    reminders: { useDefault: true }
  };
  try {
    var created = Calendar.Events.insert(event, consultantEmail, { sendUpdates: 'all' });
    var evId = created && created.id ? created.id : '';
    tsAudit_('CALENDAR', ref, 'Event on ' + consultantEmail + ' id=' + evId, 'OK');
    return { sent: true, eventId: evId };
  } catch (err) {
    tsAudit_('CALENDAR', ref,
      'Failed for ' + consultantEmail + ': ' + String(err) +
      ' — consultant may need to share calendar with script owner (Make changes to events).', 'WARN');
    return none;
  }
}

function tsExpireSiblingTokens_(email, salesGroup, usedToken) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.BOOKING_TOKENS);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var used = String(usedToken || '').trim().toLowerCase();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var tok = String(values[i][TS.BOOK_COLS.TOKEN - 1] || '').trim().toLowerCase();
    if (tok === used) continue;
    var rowEmail = String(values[i][TS.BOOK_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(values[i][TS.BOOK_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(values[i][TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (rowEmail === targetEmail && rowGroup === targetGroup && status === 'PENDING') {
      tsUpdateBookingTokenStatus_(i + 1, 'SUPERSEDED');
    }
  }
}

function tsMarkOfferBooked_(offerRow, dateStr, startStr, endStr) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || offerRow < 2) return;
  sheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('BOOKED');
  sheet.getRange(offerRow, TS.OFFER_COLS.BOOKED_AT).setValue(new Date());
  sheet.getRange(offerRow, TS.OFFER_COLS.BOOKED_WINDOW).setValue(
    dateStr + ' ' + startStr + '-' + endStr
  );
  SpreadsheetApp.flush();
  tsRefreshAnalyticsSafe_();
}

function tsExecuteEscalateOffer_(escalateToken) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Please try again in a moment.' };
  }
  try {
    var hit = tsLookupOfferByEscalateToken_(escalateToken);
    if (!hit) return { ok: false, message: 'Invalid or unknown link.' };

    var row = hit.row;
    var rowNum = hit.rowNum;
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status === 'BOOKED') {
      return { ok: true, message: 'Your training is already booked. You can close this page.' };
    }

    var alreadyEscalated = status === 'ESCALATED';
    if (!alreadyEscalated) {
      tsMarkOfferEscalated_(rowNum, row);
      tsNotifyManagerEscalation_(row, escalateToken);
      tsAudit_('ESCALATE', escalateToken,
        String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '') + ' escalated to manager', 'OK');
    }

    return {
      ok: true,
      message: 'Thank you! Your manager will reach out to help you schedule your Reflex-AI training.'
    };
  } catch (err) {
    var errMsg = String(err && err.message ? err.message : err);
    tsAudit_('ESCALATE', String(escalateToken || ''), errMsg, 'FAILED');
    return { ok: false, message: 'Something went wrong. Please contact your manager or WFM.' };
  } finally {
    lock.releaseLock();
  }
}

function tsLookupOfferByEscalateToken_(escalateToken) {
  var t = String(escalateToken || '').trim().toLowerCase();
  if (!t) return null;
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var cell = String(values[i][TS.OFFER_COLS.ESCALATE_TOKEN - 1] || '').trim().toLowerCase();
    if (cell === t) return { rowNum: i + 1, row: values[i] };
  }
  return null;
}

function tsMarkOfferEscalated_(offerRowNum, row) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || offerRowNum < 2) return;
  sheet.getRange(offerRowNum, TS.OFFER_COLS.STATUS).setValue('ESCALATED');
  var tokenIds = String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '').split(',').map(function(s) {
    return s.trim();
  }).filter(Boolean);
  tokenIds.forEach(function(tid) {
    var hit = tsLookupBookingTokenRow_(tid);
    if (!hit) return;
    var tokStatus = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (tokStatus === 'PENDING') tsUpdateBookingTokenStatus_(hit.rowNum, 'CANCELLED');
  });
  SpreadsheetApp.flush();
  tsRefreshAnalyticsSafe_();
}

function tsNotifyManagerEscalation_(row, ref) {
  var consultantName = String(row[TS.OFFER_COLS.CONSULTANT_NAME - 1] || '').trim();
  var consultantEmail = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim();
  var managerName = redirectMiguelToLiz_(String(row[TS.OFFER_COLS.MANAGER_NAME - 1] || '').trim());
  var incompleteCount = Number(row[TS.OFFER_COLS.INCOMPLETE_COUNT - 1] || 0);
  var durationMin = Number(row[TS.OFFER_COLS.DURATION_MIN - 1] || 60);
  var simsCsv = String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '').trim();
  var salesGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim();
  var searchDays = Number(tsLoadConfig_().SEARCH_DAYS || TS.SEARCH_DAYS);

  if (!managerName) {
    tsAudit_('ESCALATE', ref || '', 'No manager on offer — manager DM skipped for ' + consultantEmail, 'WARN');
    return;
  }

  var simWord = incompleteCount === 1 ? 'simulation' : 'simulations';
  var msg = [
    '*Reflex-AI training — consultant needs help scheduling*',
    '',
    '*' + consultantName + '* (' + consultantEmail + ') could not use any of the offered training windows.',
    '',
    'They have *' + incompleteCount + '* ' + simWord + ' to complete' +
      (simsCsv ? ' (' + simsCsv + ')' : '') + '.',
    'Please set aside *' + tsFormatDurationLabel_(durationMin) + '* in the next *' +
      searchDays + ' weekdays* for Reflex-AI training (' + salesGroup + ').',
    '',
    'Work with the consultant to find a time and add a *Training* block to their schedule in Assembled.'
  ].join('\n');

  try {
    tsSlackDmManager_(managerName, msg);
  } catch (err) {
    tsAudit_('ESCALATE', ref || '', 'Manager DM failed: ' + err, 'WARN');
  }
}

function tsFormatDurationLabel_(durationMin) {
  var n = Number(durationMin || 60);
  if (n >= 60 && n % 60 === 0) {
    var hrs = n / 60;
    return hrs === 1 ? '1 hour' : hrs + ' hours';
  }
  return n + ' minutes';
}

function tsEnsureOfferEscalateToken_(offerSheet, rowNum, row) {
  var token = String(row[TS.OFFER_COLS.ESCALATE_TOKEN - 1] || '').trim();
  if (token) return token;
  token = Utilities.getUuid().replace(/-/g, '');
  offerSheet.getRange(rowNum, TS.OFFER_COLS.ESCALATE_TOKEN).setValue(token);
  SpreadsheetApp.flush();
  return token;
}

/***************************************
 * SLACK
 ***************************************/

function tsBuildOfferSlackMessage_(need, windows, tokens) {
  var simLine = need.sims.length ? need.sims.join(', ') : '(see ReflexAI sheet)';
  var lines = [
    '*Reflex-AI training — pick a time*',
    'Hi ' + need.name + ', you have *' + need.incompleteCount + '* simulation' +
      (need.incompleteCount === 1 ? '' : 's') + ' to complete (' + need.durationMin + ' min).',
    'Sims: ' + simLine,
    '',
    '*Choose one window (CT):*'
  ];
  for (var i = 0; i < windows.length; i++) {
    var w = windows[i];
    var link = tsBuildBookItSlackLink_(tokens[i]);
    lines.push(
      '• ' + tsFriendlyDate_(w.dateStr) + ', ' +
      tsFriendlyTime_(w.startStr) + '–' + tsFriendlyTime_(w.endStr) +
      (link ? '  ' + link : '')
    );
  }
  if (!tsGetWebAppUrl_()) {
    lines.push('', '_Book-it links unavailable — set TS_WEB_APP_URL on TS Config (row Key/Value)._');
  }
  lines.push('', '_Tap 📅 Book it next to your preferred time._');
  var escLink = tsBuildEscalateSlackLink_(need.escalateToken);
  if (escLink) lines.push('', escLink);
  return lines.join('\n');
}

function tsBuildReminderSlackMessage_(need) {
  var lines = [
    '*Reminder: Reflex-AI training*',
    'Hi ' + need.name + ', please book your ' + need.durationMin + '-minute training window.',
    'Sims still needed: ' + (need.sims.length ? need.sims.join(', ') : '(see sheet)') + '',
    '',
    '*Your offered windows (CT):*'
  ];
  need.tokenIds.forEach(function(tid) {
    var hit = tsLookupBookingTokenRow_(tid);
    if (!hit) return;
    var row = hit.row;
    if (String(row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase() !== 'PENDING') return;
    var dateStr = tsParseDateStr_(row[TS.BOOK_COLS.DATE_STR - 1]);
    var startStr = tsParseTimeInt_(row[TS.BOOK_COLS.START_STR - 1]);
    var endStr = tsParseTimeInt_(row[TS.BOOK_COLS.END_STR - 1]);
    var link = tsBuildBookItSlackLink_(tid);
    lines.push(
      '• ' + tsFriendlyDate_(dateStr) + ', ' +
      tsFriendlyTime_(startStr) + '–' + tsFriendlyTime_(endStr) +
      (link ? '  ' + link : '')
    );
  });
  lines.push('', '_Book within the next few days before these links expire._');
  var escLink = tsBuildEscalateSlackLink_(need.escalateToken);
  if (escLink) lines.push('', escLink);
  return lines.join('\n');
}

function tsSlackDmConsultant_(name, email, message) {
  if (tsIsTestMode_()) {
    tsSlackDmTestRecipient_('*[TEST MODE — would have gone to ' + (name || email) + ']*\n\n' + message);
    return;
  }
  var userId = tsSlackLookupUserId_(name) || tsSlackLookupUserIdByEmail_(email);
  if (!userId) {
    tsAudit_('SLACK', name || email, 'Could not resolve consultant for DM', 'WARN');
    return;
  }
  tsSlackSendDm_(userId, message);
}

function tsSlackDmTestRecipient_(message) {
  var userId = tsSlackLookupTestModeUserId_();
  if (!userId) {
    tsAudit_('SLACK', 'TEST_MODE', 'DM lookup failed — set TEST_MODE_SLACK_EMAIL on TS Config', 'FAILED');
    return false;
  }
  return tsSlackSendDm_(userId, message);
}

function tsSlackLookupTestModeUserId_() {
  var config = tsLoadConfig_();
  var configured = String(config.TEST_MODE_SLACK_EMAIL || '').trim();
  if (configured) {
    var byConfig = tsSlackLookupUserIdByEmail_(configured);
    if (byConfig) return byConfig;
  }
  return tsSlackLookupUserId_('robert.sorrell') ||
    tsSlackLookupUserIdByEmail_('robert.sorrell@varsitytutors.com');
}

function tsSlackDmAlias_(alias, message) {
  var userId = tsSlackLookupUserId_(alias);
  if (!userId) {
    tsAudit_('SLACK', alias, 'DM lookup failed', 'WARN');
    return;
  }
  tsSlackSendDm_(userId, message);
}

function tsSlackCandidateEmails_(alias) {
  var candidates = [];
  var seen = {};
  var add = function(email) {
    var e = String(email || '').trim().toLowerCase();
    if (!e || seen[e]) return;
    seen[e] = true;
    candidates.push(e);
  };
  var raw = String(alias || '').replace(/^@/, '').trim();
  if (!raw) return candidates;
  // Keep legacy manager aliases routed to Liz after the manager transition.
  var managerKey = normalizePersonKey_(raw.split('@')[0]);
  if (managerKey === 'miguel morales' || managerKey === 'liz berkley' || managerKey === 'elizabeth berkley') {
    add('elizabeth.berkley@varsitytutors.com');
    return candidates;
  }
  if (raw.indexOf('@') !== -1) {
    add(raw);
    return candidates;
  }
  add(raw.toLowerCase().replace(/\s+/g, '.').replace(/[^a-z0-9.]/g, '') + '@varsitytutors.com');
  if (raw.indexOf('.') === -1) {
    add(raw.toLowerCase().replace(/\s+/g, '') + '@varsitytutors.com');
  }
  return candidates;
}

function tsSlackLookupUserId_(alias) {
  var emails = tsSlackCandidateEmails_(alias);
  for (var i = 0; i < emails.length; i++) {
    var userId = tsSlackLookupUserIdByEmail_(emails[i]);
    if (userId) return userId;
  }
  return null;
}

function tsSlackLookupUserIdByEmail_(email) {
  email = redirectMiguelToLiz_(email);
  if (!email) return null;
  try {
    var token = tsGetSlackToken_();
    var url = 'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(email);
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.ok) return data.user.id;
    tsAudit_('SLACK', email, data.error || 'users.lookupByEmail failed', 'WARN');
  } catch (err) {
    tsAudit_('SLACK', email, String(err), 'WARN');
  }
  return null;
}

function tsSlackSendDm_(userId, message) {
  try {
    var token = tsGetSlackToken_();
    var openRes = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ users: userId }),
      muteHttpExceptions: true
    });
    var openData = JSON.parse(openRes.getContentText());
    if (!openData.ok) {
      tsAudit_('SLACK', 'open', openData.error || 'conversations.open failed', 'WARN');
      return false;
    }
    var channelId = openData.channel.id;
    var postRes = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ channel: channelId, text: redirectMiguelToLiz_(message), unfurl_links: false }),
      muteHttpExceptions: true
    });
    var postData = JSON.parse(postRes.getContentText());
    if (!postData.ok) {
      tsAudit_('SLACK', 'post', postData.error || 'chat.postMessage failed', 'WARN');
      return false;
    }
    tsAudit_('SLACK', userId, 'DM sent', 'OK');
    return true;
  } catch (err) {
    tsAudit_('SLACK', 'dm', String(err), 'FAILED');
    return false;
  }
}

function tsBuildEscalateSlackLink_(escalateToken) {
  var base = tsGetWebAppUrl_();
  if (!base || !escalateToken) return '';
  var url = base + (base.indexOf('?') === -1 ? '?' : '&') +
    'escalate=' + encodeURIComponent(escalateToken);
  return '<' + url + '|None of these times work!>';
}

function tsBuildBookItSlackLink_(token) {
  var base = tsGetWebAppUrl_();
  if (!base || !token) return '';
  var url = base + (base.indexOf('?') === -1 ? '?' : '&') + 'token=' + encodeURIComponent(token);
  return '<' + url + '|📅 Book it>';
}

function tsGetWebAppUrl_() {
  var url = String(PropertiesService.getScriptProperties().getProperty('TS_WEB_APP_URL') || '').trim();
  if (!url) url = String(tsLoadConfig_().TS_WEB_APP_URL || '').trim();
  if (!url) url = String(TS.WEB_APP_URL || '').trim();
  return url.replace(/\/+$/, '');
}

function tsDefaultWebAppUrl_() {
  return String(TS.WEB_APP_URL || '').trim();
}

function tsBookingResponsePage_(message, ok) {
  var accent = ok ? '#1a7f4b' : '#b42318';
  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
    '<title>', TS.BRAND, '</title></head>',
    '<body style="margin:0;font-family:system-ui,sans-serif;background:#0f172a;min-height:100vh;display:flex;align-items:center;justify-content:center;">',
    '<div style="max-width:420px;padding:32px;text-align:center;">',
    '<div style="font-size:22px;font-weight:bold;color:#fff;margin-bottom:16px;">', TS.BRAND, '</div>',
    '<div style="font-size:17px;color:', accent, ';font-weight:600;margin-bottom:12px;">', tsEscHtml_(message), '</div>',
    '<div style="font-size:14px;color:#94a3b8;">You can close this page.</div>',
    '</div></body></html>'
  ].join('');
}

/***************************************
 * ASSEMBLED API
 ***************************************/

function tsResolveTrainingActivityTypeId_(headers) {
  var override = String(PropertiesService.getScriptProperties().getProperty('ASSEMBLED_TRAINING_ACTIVITY_ID') || '').trim();
  if (override) return override;
  return tsResolveActivityTypeIdByName_(headers, 'Training');
}

function tsResolveActivityTypeIdByName_(headers, typeName) {
  var target = String(typeName || '').trim().toLowerCase();
  var list = tsListActivityTypes_(headers);
  for (var i = 0; i < list.length; i++) {
    if ((list[i].name || '').trim().toLowerCase() === target) {
      return String(list[i].id || '').trim();
    }
  }
  tsAudit_('ASSEMBLED', '/activity_types', 'Type not found: ' + typeName, 'WARN');
  return null;
}

function tsListActivityTypes_(headers) {
  var data = tsAssembledGet_(headers, '/activity_types', {});
  var raw = data.activity_types || {};
  return Array.isArray(raw)
    ? raw
    : Object.keys(raw).map(function(k) {
      var a = raw[k];
      if (!a.id) a.id = k;
      return a;
    });
}

function tsResolveAgentIdByEmail_(headers, email) {
  var target = String(email || '').trim().toLowerCase();
  var LIMIT = 100;
  var offset = 0;
  while (true) {
    var res = tsAssembledGet_(headers, '/people', { limit: LIMIT, offset: offset });
    var people = res.people || {};
    var total = res.total || 0;
    var keys = Object.keys(people);
    for (var i = 0; i < keys.length; i++) {
      var person = people[keys[i]];
      var personEmail = (person.email || '').trim().toLowerCase();
      var agentId = (person.agent_id || keys[i] || '').trim();
      if (personEmail === target && agentId) return agentId;
    }
    if (keys.length < LIMIT || offset + LIMIT >= total) break;
    offset += LIMIT;
    Utilities.sleep(200);
  }
  return '';
}

function tsResolveSiteId_(headers, siteName) {
  var res = tsAssembledGet_(headers, '/sites', {});
  var sites = res.sites || {};
  var target = tsNormalizeToken_(siteName);
  for (var id in sites) {
    if (tsNormalizeToken_(sites[id].name) === target) return sites[id].id || id;
  }
  throw new Error('Site not found: ' + siteName);
}

function tsResolveQueueId_(headers, queueName) {
  var res = tsAssembledGet_(headers, '/queues', {});
  var queues = res.queues || {};
  var target = tsNormalizeToken_(queueName);
  for (var id in queues) {
    if (tsNormalizeToken_(queues[id].name) === target) return queues[id].id || id;
  }
  return null;
}

function tsSafeResolveQueueId_(headers, queueName) {
  try {
    return tsResolveQueueId_(headers, queueName);
  } catch (err) {
    tsAudit_('ASSEMBLED', '/queues', String(err), 'WARN');
    return null;
  }
}

function tsAssembledGet_(headers, path, params) {
  var url = tsBuildUrl_(TS.ASSEMBLED.BASE_URL + path, params);
  var resp = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Assembled GET ' + path + ' failed (' + code + '): ' + text);
  }
  return text ? JSON.parse(text) : {};
}

function tsAssembledPost_(headers, path, payload) {
  var resp = UrlFetchApp.fetch(TS.ASSEMBLED.BASE_URL + path, {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Assembled POST ' + path + ' failed (' + code + '): ' + text);
  }
  return text ? JSON.parse(text) : {};
}

/***************************************
 * SHEET HELPERS — bot tabs
 ***************************************/

function tsEnsureConfigSheet_(ss) {
  var sheet = ss.getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) sheet = ss.insertSheet(TS.SHEETS.CONFIG);
  if (tsConfigSheetNeedsRepair_(sheet)) {
    tsRepairConfigSheet_(ss);
    return ss.getSheetByName(TS.SHEETS.CONFIG);
  }
  var settingsDefaults = tsBuildSettingsConfigRows_(tsScrapeConfigValues_(sheet));
  var needsSeed = sheet.getLastRow() < 2 || !tsSheetHasConfigKey_(sheet, 'TEST_MODE');
  if (needsSeed) {
    sheet.clearContents();
    sheet.getRange(1, 1, settingsDefaults.length, 3).setValues(settingsDefaults);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  } else {
    tsEnsureOptionalConfigKey_(sheet, 'TEST_MODE_SLACK_EMAIL', 'robert.sorrell@varsitytutors.com',
      'TEST_MODE Slack DM recipient');
    tsEnsureOptionalConfigKey_(sheet, 'SLACK_BOT_TOKEN', '',
      'ops_bot token (xoxb-…) — required for Slack DMs');
    tsEnsureOptionalConfigKey_(sheet, 'ASSEMBLED_API_KEY', '',
      'Assembled sk_live_… — or set in Script Properties');
    tsEnsureOptionalConfigKey_(sheet, 'TS_WEB_APP_URL', '',
      'Deployed Book-it web app URL ending in /exec');
  }
  tsSyncManagerAliasRows_(ss, sheet);
  return sheet;
}

/** "Aaron Bunch" → aaron.bunch for Slack users.lookupByEmail */
function tsManagerNameToSlackAlias_(name) {
  if (normalizeManagerDisplayName_(name) === 'Liz Berkley') return 'elizabeth.berkley';
  var parts = tsNormalizePersonName_(name).split(' ').filter(Boolean);
  if (parts.length < 2) return parts.join('.');
  return parts[0] + '.' + parts[parts.length - 1];
}

function tsKnownConfigKeys_() {
  return [
    'TEST_MODE', 'OFFERS_ENABLED', 'REMINDERS_ENABLED', 'GOOGLE_CALENDAR_ENABLED',
    'MIN_BUFFER', 'SEARCH_DAYS', 'BUSINESS_HOUR_START', 'BUSINESS_HOUR_END',
    'TEST_MODE_SLACK_EMAIL', 'SLACK_BOT_TOKEN', 'ASSEMBLED_API_KEY', 'TS_WEB_APP_URL'
  ];
}

function tsBuildSettingsConfigRows_(scraped) {
  scraped = scraped || {};
  return [
    ['Key', 'Value', 'Notes'],
    ['TEST_MODE', scraped.TEST_MODE || 'TRUE', 'TRUE redirects all Slack DMs to TEST_MODE_SLACK_EMAIL'],
    ['OFFERS_ENABLED', scraped.OFFERS_ENABLED || 'TRUE', 'Daily offer run'],
    ['REMINDERS_ENABLED', scraped.REMINDERS_ENABLED || 'TRUE', '48h reminder if not booked'],
    ['GOOGLE_CALENDAR_ENABLED', scraped.GOOGLE_CALENDAR_ENABLED || 'TRUE',
      'Create event on consultant calendar (skipped in TEST_MODE)'],
    ['MIN_BUFFER', scraped.MIN_BUFFER || TS.MIN_BUFFER, 'Min post-booking net staffing'],
    ['SEARCH_DAYS', scraped.SEARCH_DAYS || TS.SEARCH_DAYS, 'Weekdays to search (M-F only)'],
    ['BUSINESS_HOUR_START', scraped.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START, 'Earliest hour CT'],
    ['BUSINESS_HOUR_END', scraped.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END,
      'Latest hour CT (slot must end by this hour)'],
    ['TEST_MODE_SLACK_EMAIL', scraped.TEST_MODE_SLACK_EMAIL || 'robert.sorrell@varsitytutors.com',
      'TEST_MODE Slack DM recipient'],
    ['SLACK_BOT_TOKEN', scraped.SLACK_BOT_TOKEN || '', 'ops_bot token (xoxb-…) — required for Slack DMs'],
    ['ASSEMBLED_API_KEY', scraped.ASSEMBLED_API_KEY || '', 'Assembled sk_live_… — or Script Properties'],
    ['TS_WEB_APP_URL', scraped.TS_WEB_APP_URL || tsDefaultWebAppUrl_(),
      'Deployed Book-it web app URL ending in /exec'],
    ['', '', '']
  ];
}

function tsScrapeConfigValues_(sheet) {
  var scraped = {};
  var known = tsKnownConfigKeys_();
  try {
    var props = PropertiesService.getScriptProperties();
    known.forEach(function(k) {
      var v = String(props.getProperty(k) || '').trim();
      if (v) scraped[k] = v;
    });
  } catch (ignore) {}

  if (!sheet || sheet.getLastRow() < 2) return scraped;

  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  var endRow = sectionRow > 0 ? sectionRow - 1 : sheet.getLastRow();
  var values = sheet.getRange(1, 1, endRow, 3).getValues();

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < 3; c++) {
      var cell = String(values[r][c] || '').trim();
      if (!cell) continue;
      if (cell.indexOf('xoxb-') === 0) scraped.SLACK_BOT_TOKEN = cell;
      if (cell.indexOf('sk_live_') === 0) scraped.ASSEMBLED_API_KEY = cell;
      if (cell.indexOf('script.google.com') !== -1 && cell.indexOf('/exec') !== -1) {
        scraped.TS_WEB_APP_URL = cell;
      }
      if (cell.indexOf('@varsitytutors.com') !== -1) scraped.TEST_MODE_SLACK_EMAIL = cell;
    }
  }

  for (var i = 1; i < values.length; i++) {
    var a = String(values[i][0] || '').trim();
    var b = String(values[i][1] || '').trim();
    var c = String(values[i][2] || '').trim();
    var rowKeys = [a, c];
    for (var k = 0; k < known.length; k++) {
      var key = known[k];
      var keyUpper = key.toUpperCase();
      for (var j = 0; j < rowKeys.length; j++) {
        var candidate = String(rowKeys[j] || '').trim().toUpperCase();
        if (!candidate) continue;
        if (candidate === keyUpper || (candidate.length >= 8 && keyUpper.indexOf(candidate) === 0)) {
          if (b && b.toUpperCase() !== keyUpper && b.indexOf('xoxb-') !== 0 &&
              b.indexOf('sk_live_') !== 0 && b.indexOf('script.google.com') === -1) {
            scraped[key] = b;
          }
        }
      }
    }
    if (a.toUpperCase() === 'TEST_MODE' && (b === 'TRUE' || b === 'FALSE')) scraped.TEST_MODE = b;
  }
  if (!scraped.TS_WEB_APP_URL) scraped.TS_WEB_APP_URL = tsDefaultWebAppUrl_();
  return scraped;
}

function tsConfigSheetNeedsRepair_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return true;
  var h1 = String(sheet.getRange(1, 1).getValue() || '').trim();
  var h2 = String(sheet.getRange(1, 2).getValue() || '').trim();
  var h3 = String(sheet.getRange(1, 3).getValue() || '').trim();
  if (h1 !== 'Key' || h2 !== 'Value' || h3 !== 'Notes') return true;

  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  var endRow = sectionRow > 0 ? sectionRow - 1 : Math.min(sheet.getLastRow(), 20);
  if (endRow < 2) return true;
  var values = sheet.getRange(2, 1, endRow, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) continue;
    if (key === 'Manager Name') break;
    var note = String(values[i][2] || '').trim();
    if (note && tsKnownConfigKeys_().indexOf(note) >= 0 && note !== key) return true;
    if (key.length >= 4 && tsKnownConfigKeys_().indexOf(key) < 0) {
      var partial = tsKnownConfigKeys_().some(function(k) {
        return k.indexOf(key) === 0 || key.indexOf(k.slice(0, Math.min(key.length, k.length))) === 0;
      });
      if (partial && note && tsKnownConfigKeys_().indexOf(note) >= 0) return true;
    }
  }
  return false;
}

function tsRepairConfigSheet_(ss) {
  ss = ss || tsGetSpreadsheet_();
  var sheet = ss.getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) sheet = ss.insertSheet(TS.SHEETS.CONFIG);
  var scraped = tsScrapeConfigValues_(sheet);
  var existingAliases = tsReadExistingManagerAliases_(sheet);
  var settingsRows = tsBuildSettingsConfigRows_(scraped);
  sheet.clear();
  sheet.getRange(1, 1, settingsRows.length, 3).setValues(settingsRows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  tsSyncManagerAliasRows_(ss, sheet, existingAliases);
  return sheet;
}

function tsCollectRegionalDirectors_(ss) {
  var sheet = ss.getSheetByName(TS.SHEETS.MANAGER_LOOKUP);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
  var colMgr = tsFindCol_(header, ['regional director', 'manager name', 'sales manager']);
  if (colMgr < 0) return [];
  var seen = {};
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var mgr = String(values[i][colMgr] || '').trim();
    if (mgr) mgr = normalizeManagerDisplayName_(mgr);
    if (!mgr || seen[mgr.toLowerCase()]) continue;
    seen[mgr.toLowerCase()] = true;
    list.push(mgr);
  }
  list.sort(function(a, b) { return a.localeCompare(b, undefined, { sensitivity: 'base' }); });
  return list;
}

function tsFindManagerAliasSectionRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || '').trim() === 'Manager Name') return i + 1;
  }
  return -1;
}

function tsReadExistingManagerAliases_(sheet) {
  var map = {};
  var values = sheet.getDataRange().getValues();
  var inSection = false;
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0] || '').trim();
    var alias = String(values[i][1] || '').trim();
    if (name === 'Manager Name') { inSection = true; continue; }
    if (inSection) {
      if (!name) break;
      map[name.toLowerCase()] = alias;
    }
  }
  return map;
}

function tsSyncManagerAliasRows_(ss, sheet, preservedAliases) {
  sheet = sheet || ss.getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) return;
  var managers = tsCollectRegionalDirectors_(ss);
  var existingAliases = preservedAliases || tsReadExistingManagerAliases_(sheet);
  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  if (sectionRow < 0) {
    var lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow > 1 && String(sheet.getRange(lastRow, 1).getValue() || '').trim() !== '') {
      sheet.getRange(lastRow + 1, 1, 1, 3).setValues([['', '', '']]);
      sectionRow = lastRow + 2;
    } else {
      sectionRow = lastRow + 1;
    }
  }
  if (sheet.getLastRow() >= sectionRow) {
    sheet.getRange(sectionRow, 1, sheet.getLastRow() - sectionRow + 1, 3).clearContent();
  }
  var rows = [['Manager Name', 'Slack Alias', 'Notes']];
  managers.forEach(function(mgr) {
    var autoAlias = tsManagerNameToSlackAlias_(mgr);
    var existing = existingAliases[mgr.toLowerCase()] || '';
    var alias = existing;
    if (!existing || existing.toLowerCase() === mgr.toLowerCase()) {
      alias = autoAlias;
    }
    rows.push([mgr, alias, '']);
  });
  if (!managers.length) {
    rows.push(['(no managers found in Looker Manager Lookup col D)', '', '']);
  }
  sheet.getRange(sectionRow, 1, rows.length, 3).setValues(rows);
  sheet.getRange(sectionRow, 1, 1, 3).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
}

function tsSheetHasConfigKey_(sheet, key) {
  if (sheet.getLastRow() < 2) return false;
  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  var endRow = sectionRow > 0 ? sectionRow - 1 : sheet.getLastRow();
  if (endRow < 2) return false;
  var keys = sheet.getRange(2, 1, endRow - 1, 1).getValues();
  var target = String(key || '').trim().toUpperCase();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '').trim().toUpperCase() === target) return true;
  }
  return false;
}

function tsEnsureOptionalConfigKey_(sheet, key, value, notes) {
  if (tsSheetHasConfigKey_(sheet, key)) return;
  var sectionRow = tsFindManagerAliasSectionRow_(sheet);
  var insertRow = sectionRow > 0 ? sectionRow - 1 : sheet.getLastRow() + 1;
  if (insertRow < 2) insertRow = 2;
  var existing = String(sheet.getRange(insertRow, 1).getValue() || '').trim();
  if (existing) {
    sheet.insertRowBefore(insertRow);
  }
  sheet.getRange(insertRow, 1, 1, 3).setValues([[key, value, notes || '']]);
}

function tsEnsureOffersSheet_(ss) {
  ss = ss || tsGetSpreadsheet_();
  var sheet = ss.getSheetByName(TS.SHEETS.OFFERS);
  var header = [
    'Created At', 'Consultant', 'Email', 'Sales Group', 'Queue', 'Manager',
    'Incomplete #', 'Sims', 'Duration Min', 'Status', 'Offer Sent At',
    'Reminder Sent At', 'Token IDs', 'Booked At', 'Booked Window', 'Escalate Token'
  ];
  if (!sheet) {
    sheet = ss.insertSheet(TS.SHEETS.OFFERS);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  } else if (String(sheet.getRange(1, 1).getValue() || '').trim() !== 'Created At') {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  } else if (String(sheet.getRange(1, TS.OFFER_COLS.ESCALATE_TOKEN).getValue() || '').trim() !== 'Escalate Token') {
    sheet.getRange(1, TS.OFFER_COLS.ESCALATE_TOKEN).setValue('Escalate Token');
  }
  return sheet;
}

function tsValidateSlackReady_() {
  try {
    tsGetSlackToken_();
    return true;
  } catch (err) {
    return false;
  }
}

function tsEnsureBookingTokensSheet_(ss) {
  var sheet = ss.getSheetByName(TS.SHEETS.BOOKING_TOKENS);
  if (!sheet) {
    sheet = ss.insertSheet(TS.SHEETS.BOOKING_TOKENS);
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Token', 'Created At', 'Status', 'Consultant', 'Email', 'Manager',
      'Sales Group', 'Queue', 'Sims', 'Duration Min', 'Date', 'Start', 'End', 'Offer Row'
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  }
  return sheet;
}

function tsEnsureAuditSheet_(ss) {
  var sheet = ss.getSheetByName(TS.SHEETS.AUDIT);
  if (!sheet) {
    sheet = ss.insertSheet(TS.SHEETS.AUDIT);
    sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Event', 'Reference', 'Details', 'Result']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  }
  return sheet;
}

function tsEnsureAnalyticsSheet_(ss) {
  ss = ss || tsGetSpreadsheet_();
  var sheet = ss.getSheetByName(TS.SHEETS.ANALYTICS);
  if (!sheet) {
    sheet = ss.insertSheet(TS.SHEETS.ANALYTICS);
    sheet.setTabColor('#4f8cff');
  }
  return sheet;
}

/***************************************
 * TS ANALYTICS
 ***************************************/

function tsRefreshAnalytics_(opts) {
  opts = opts || {};
  var ss = tsGetSpreadsheet_();
  var sheet = tsEnsureAnalyticsSheet_(ss);
  var data = tsBuildAnalyticsData_();
  var now = new Date();

  sheet.clear();
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue(TS.BRAND + ' — Pipeline Analytics')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#1F4E78')
    .setFontColor('#ffffff');
  sheet.getRange(2, 1).setValue('Last refreshed').setFontWeight('bold');
  sheet.getRange(2, 2).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  var kpiStart = 4;
  sheet.getRange(kpiStart, 1).setValue('Company snapshot').setFontWeight('bold').setBackground('#eef3fb');
  var kpis = [
    ['Total reps (all CSV tabs)', data.summary.totalReps],
    ['Reps with outstanding sims', data.summary.outstandingReps],
    ['Reps scheduled (booked)', data.summary.scheduledReps],
    ['Reps with open offers', data.summary.openOfferReps],
    ['Reps escalated', data.summary.escalatedReps],
    ['Reps needing scheduling', data.summary.needsSchedulingReps],
    ['Total sim rows', data.summary.totalSimRows],
    ['Outstanding sim rows', data.summary.outstandingSimRows],
    ['Not started sim rows', data.summary.notStartedSimRows],
    ['Incomplete sim rows', data.summary.incompleteSimRows],
    ['Completed ≥ ' + TS.PASSING_SCORE_PERCENT + '%', data.summary.completedAboveSimRows],
    ['Completed < ' + TS.PASSING_SCORE_PERCENT + '%', data.summary.completedBelowSimRows],
    ['Avg score (completed sims)', data.summary.avgCompletedScore === null ? '' : data.summary.avgCompletedScore / 100]
  ];
  tsSheetWriteRows_(sheet, kpiStart + 1, 1, kpis);
  sheet.getRange(kpiStart + kpis.length, 2).setNumberFormat('0.0%');

  var mgrHeaderRow = kpiStart + kpis.length + 2;
  sheet.getRange(mgrHeaderRow, 1).setValue('By manager').setFontWeight('bold').setBackground('#eef3fb');
  var mgrHeaders = [
    'Manager', 'Reps', 'Outstanding Reps', 'Scheduled Reps', 'Open Offer Reps', 'Escalated Reps',
    'Outstanding Sims', 'Not Started Sims', 'Incomplete Sims', 'Completed ≥70%', 'Completed <70%',
    'Avg Completed Score', 'Needs Scheduling Reps'
  ];
  tsSheetWriteRows_(sheet, mgrHeaderRow + 1, 1, [mgrHeaders]);
  tsStyleAnalyticsHeader_(sheet, mgrHeaderRow + 1, mgrHeaders.length);

  if (data.managers.length) {
    var mgrRows = data.managers.map(function(m) {
      return [
        m.manager,
        m.reps,
        m.outstandingReps,
        m.scheduledReps,
        m.openOfferReps,
        m.escalatedReps,
        m.outstandingSims,
        m.notStartedSims,
        m.incompleteSims,
        m.completedAboveSims,
        m.completedBelowSims,
        m.avgCompletedScore === null ? '' : m.avgCompletedScore / 100,
        m.needsSchedulingReps
      ];
    });
    tsSheetWriteRows_(sheet, mgrHeaderRow + 2, 1, mgrRows);
    sheet.getRange(mgrHeaderRow + 2, 12, mgrRows.length, 1).setNumberFormat('0.0%');
  }

  var repHeaderRow = mgrHeaderRow + 2 + Math.max(data.managers.length, 1) + 2;
  sheet.getRange(repHeaderRow, 1).setValue('By rep').setFontWeight('bold').setBackground('#eef3fb');
  var repHeaders = [
    'Consultant', 'Email', 'Sales Group', 'Manager', 'Total Sims', 'Outstanding Sims',
    'Not Started', 'Incomplete', 'Completed ≥70%', 'Completed <70%', 'Avg Score',
    'Pipeline Status', 'Offer Sent', 'Reminder Sent', 'Booked Window', 'Outstanding Sim Names'
  ];
  tsSheetWriteRows_(sheet, repHeaderRow + 1, 1, [repHeaders]);
  tsStyleAnalyticsHeader_(sheet, repHeaderRow + 1, repHeaders.length);

  if (data.reps.length) {
    var repRows = data.reps.map(function(r) {
      return [
        r.name,
        r.email,
        r.salesGroup,
        r.managerName || 'Unassigned',
        r.simTotal,
        r.outstandingSims,
        r.notStarted,
        r.incomplete,
        r.completedAbove,
        r.completedBelow,
        r.avgScore === null ? '' : r.avgScore / 100,
        r.pipelineStatus,
        r.offerSentAt || '',
        r.reminderSentAt || '',
        r.bookedWindow || '',
        r.outstandingSimNames
      ];
    });
    tsSheetWriteRows_(sheet, repHeaderRow + 2, 1, repRows);
    sheet.getRange(repHeaderRow + 2, 11, repRows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(repHeaderRow + 2, 13, repRows.length, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  sheet.setFrozenRows(repHeaderRow + 1);
  sheet.autoResizeColumns(1, repHeaders.length);
  tsAudit_('ANALYTICS', '', 'Refreshed — ' + data.reps.length + ' rep rows, ' + data.managers.length + ' managers', 'OK');

  if (opts.showAlert) {
    SpreadsheetApp.getUi().alert(
      'TS Analytics',
      data.summary.outstandingReps + ' reps with outstanding sims\n' +
      data.summary.scheduledReps + ' scheduled · ' +
      data.summary.openOfferReps + ' open offers · ' +
      data.summary.escalatedReps + ' escalated',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/** getRange(row, col, numRows, numCols) — avoids end-row vs row-count mistakes. */
function tsSheetWriteRows_(sheet, startRow, startCol, values) {
  if (!values || !values.length) return;
  var numRows = values.length;
  var numCols = values[0].length;
  sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
}

function tsStyleAnalyticsHeader_(sheet, row, colCount) {
  sheet.getRange(row, 1, 1, colCount)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#ffffff')
    .setWrap(true);
  sheet.setFrozenRows(row);
}

function tsBuildAnalyticsData_() {
  var ss = tsGetSpreadsheet_();
  var lookupRows = tsLoadManagerLookupRows_(ss);
  var offerMap = tsLoadLatestOfferByRepKey_();
  var reps = [];
  var summary = {
    totalReps: 0,
    outstandingReps: 0,
    scheduledReps: 0,
    openOfferReps: 0,
    escalatedReps: 0,
    needsSchedulingReps: 0,
    totalSimRows: 0,
    outstandingSimRows: 0,
    notStartedSimRows: 0,
    incompleteSimRows: 0,
    completedAboveSimRows: 0,
    completedBelowSimRows: 0,
    avgCompletedScore: null
  };
  var completedScoreTotal = 0;
  var completedScoreCount = 0;

  ss.getSheets().forEach(function(tab) {
    var tabName = tab.getName();
    if (tabName.indexOf(TS.REFLEX_TAB_PREFIX) !== 0) return;
    var salesGroup = tabName.substring(TS.REFLEX_TAB_PREFIX.length).trim();
    var groups = tsParseReflexTabDetailed_(tab);
    Object.keys(groups).forEach(function(key) {
      var g = groups[key];
      var managerName = tsResolveManagerForConsultant_(g.name, g.email, lookupRows);
      var repKey = tsRepAnalyticsKey_(g.email, salesGroup);
      var offer = offerMap[repKey] || null;
      var pipelineStatus = tsDerivePipelineStatus_(g.outstandingSims, offer);
      var avgScore = g.completedScoreCount ? g.completedScoreTotal / g.completedScoreCount : null;

      reps.push({
        name: g.name,
        email: g.email,
        salesGroup: salesGroup,
        managerName: managerName,
        simTotal: g.simTotal,
        outstandingSims: g.outstandingSims,
        notStarted: g.notStarted,
        incomplete: g.incomplete,
        completedAbove: g.completedAbove,
        completedBelow: g.completedBelow,
        avgScore: avgScore,
        pipelineStatus: pipelineStatus,
        offerSentAt: offer && offer.offerSentAt ? offer.offerSentAt : '',
        reminderSentAt: offer && offer.reminderSentAt ? offer.reminderSentAt : '',
        bookedWindow: offer && offer.bookedWindow ? offer.bookedWindow : '',
        outstandingSimNames: g.outstandingSimNames.join('; ')
      });

      summary.totalReps++;
      summary.totalSimRows += g.simTotal;
      summary.outstandingSimRows += g.outstandingSims;
      summary.notStartedSimRows += g.notStarted;
      summary.incompleteSimRows += g.incomplete;
      summary.completedAboveSimRows += g.completedAbove;
      summary.completedBelowSimRows += g.completedBelow;
      if (g.outstandingSims > 0) summary.outstandingReps++;
      if (pipelineStatus === 'Scheduled') summary.scheduledReps++;
      if (pipelineStatus === 'Offer Open') summary.openOfferReps++;
      if (pipelineStatus === 'Escalated') summary.escalatedReps++;
      if (pipelineStatus === 'Needs Scheduling') summary.needsSchedulingReps++;
      completedScoreTotal += g.completedScoreTotal;
      completedScoreCount += g.completedScoreCount;
    });
  });

  if (completedScoreCount) {
    summary.avgCompletedScore = completedScoreTotal / completedScoreCount;
  }

  reps.sort(function(a, b) {
    var mgr = String(a.managerName || '').localeCompare(String(b.managerName || ''));
    if (mgr !== 0) return mgr;
    if (b.outstandingSims !== a.outstandingSims) return b.outstandingSims - a.outstandingSims;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });

  var managerMap = {};
  reps.forEach(function(r) {
    var mgr = r.managerName || 'Unassigned';
    if (!managerMap[mgr]) {
      managerMap[mgr] = {
        manager: mgr,
        reps: 0,
        outstandingReps: 0,
        scheduledReps: 0,
        openOfferReps: 0,
        escalatedReps: 0,
        needsSchedulingReps: 0,
        outstandingSims: 0,
        notStartedSims: 0,
        incompleteSims: 0,
        completedAboveSims: 0,
        completedBelowSims: 0,
        completedScoreTotal: 0,
        completedScoreCount: 0
      };
    }
    var m = managerMap[mgr];
    m.reps++;
    m.outstandingSims += r.outstandingSims;
    m.notStartedSims += r.notStarted;
    m.incompleteSims += r.incomplete;
    m.completedAboveSims += r.completedAbove;
    m.completedBelowSims += r.completedBelow;
    if (r.outstandingSims > 0) m.outstandingReps++;
    if (r.pipelineStatus === 'Scheduled') m.scheduledReps++;
    if (r.pipelineStatus === 'Offer Open') m.openOfferReps++;
    if (r.pipelineStatus === 'Escalated') m.escalatedReps++;
    if (r.pipelineStatus === 'Needs Scheduling') m.needsSchedulingReps++;
    if (r.avgScore !== null && (r.completedAbove + r.completedBelow) > 0) {
      m.completedScoreTotal += r.avgScore * (r.completedAbove + r.completedBelow);
      m.completedScoreCount += r.completedAbove + r.completedBelow;
    }
  });

  var managers = Object.keys(managerMap).map(function(k) {
    var m = managerMap[k];
    m.avgCompletedScore = m.completedScoreCount ? m.completedScoreTotal / m.completedScoreCount : null;
    return m;
  }).sort(function(a, b) {
    if (b.outstandingReps !== a.outstandingReps) return b.outstandingReps - a.outstandingReps;
    return String(a.manager || '').localeCompare(String(b.manager || ''));
  });

  return { summary: summary, managers: managers, reps: reps };
}

function tsParseReflexTabDetailed_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  var header = values[0].map(function(c) { return String(c || '').trim().toLowerCase(); });
  var colName = tsFindCol_(header, ['user name', 'name', 'consultant']);
  var colEmail = tsFindCol_(header, ['user email', 'email']);
  var colSim = tsFindCol_(header, ['simulation name', 'simulation', 'sim']);
  var colStatus = tsFindCol_(header, ['status']);
  var colScore = tsFindCol_(header, ['best score (%)', 'best score', 'score']);
  if (colName < 0) colName = 0;
  if (colEmail < 0) colEmail = 1;
  if (colSim < 0) colSim = 2;
  if (colStatus < 0) colStatus = 3;

  var groups = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var name = String(row[colName] || '').trim();
    var email = String(row[colEmail] || '').trim().toLowerCase();
    if (!name && !email) continue;
    var key = email || name.toLowerCase();
    if (!groups[key]) {
      groups[key] = {
        name: name,
        email: email,
        simTotal: 0,
        outstandingSims: 0,
        notStarted: 0,
        incomplete: 0,
        completedAbove: 0,
        completedBelow: 0,
        completedScoreTotal: 0,
        completedScoreCount: 0,
        outstandingSimNames: []
      };
    }
    var g = groups[key];
    g.simTotal++;
    var status = String(row[colStatus] || '').trim();
    var score = colScore >= 0 ? tsParseScorePercent_(row[colScore]) : NaN;
    var bucket = tsClassifySimRow_(status, score);
    if (bucket === 'notStarted') {
      g.notStarted++;
      g.outstandingSims++;
      g.outstandingSimNames.push(String(row[colSim] || 'Sim').trim() || 'Sim');
    } else if (bucket === 'incomplete') {
      g.incomplete++;
      g.outstandingSims++;
      g.outstandingSimNames.push(String(row[colSim] || 'Sim').trim() || 'Sim');
    } else if (bucket === 'completedAbove') {
      g.completedAbove++;
      if (!isNaN(score)) {
        g.completedScoreTotal += score;
        g.completedScoreCount++;
      }
    } else if (bucket === 'completedBelow') {
      g.completedBelow++;
      if (!isNaN(score)) {
        g.completedScoreTotal += score;
        g.completedScoreCount++;
      }
    }
    if (!g.name && name) g.name = name;
    if (!g.email && email) g.email = email;
  }
  return groups;
}

function tsLoadLatestOfferByRepKey_() {
  var map = {};
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return map;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var email = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var group = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    if (!email) continue;
    var key = tsRepAnalyticsKey_(email, group);
    var created = row[TS.OFFER_COLS.CREATED_AT - 1];
    var prev = map[key];
    var prevTime = prev && prev.created instanceof Date ? prev.created.getTime() : 0;
    var curTime = created instanceof Date ? created.getTime() : 0;
    if (!prev || curTime >= prevTime) {
      map[key] = {
        created: created,
        status: String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase(),
        offerSentAt: row[TS.OFFER_COLS.OFFER_SENT_AT - 1],
        reminderSentAt: row[TS.OFFER_COLS.REMINDER_SENT_AT - 1],
        bookedWindow: String(row[TS.OFFER_COLS.BOOKED_WINDOW - 1] || '')
      };
    }
  }
  return map;
}

function tsRepAnalyticsKey_(email, salesGroup) {
  return String(email || '').trim().toLowerCase() + '|' + String(salesGroup || '').trim().toLowerCase();
}

function tsDerivePipelineStatus_(outstandingSims, offer) {
  if (!outstandingSims) return 'All Complete';
  if (!offer) return 'Needs Scheduling';
  if (offer.status === 'BOOKED') return 'Scheduled';
  if (offer.status === 'ESCALATED') return 'Escalated';
  if (offer.status === 'OFFERED') return 'Offer Open';
  if (offer.status === 'PENDING') return 'Queued';
  return 'Needs Scheduling';
}

function tsParseScorePercent_(value) {
  if (value === '' || value === null || value === undefined) return NaN;
  if (typeof value === 'number') return value <= 1 ? value * 100 : value;
  var parsed = Number(String(value).replace('%', '').trim());
  return isFinite(parsed) ? parsed : NaN;
}

function tsClassifySimRow_(status, score) {
  var s = String(status || '').trim().toLowerCase();
  if (s === 'not started') return 'notStarted';
  if (s === 'incomplete') return 'incomplete';
  if (s === 'completed') {
    if (isNaN(score)) return 'completedBelow';
    return score >= TS.PASSING_SCORE_PERCENT ? 'completedAbove' : 'completedBelow';
  }
  return 'other';
}

function tsAppendOfferRow_(need) {
  var sheet = tsEnsureOffersSheet_(tsGetSpreadsheet_());
  var escalateToken = Utilities.getUuid().replace(/-/g, '');
  need.escalateToken = escalateToken;
  sheet.appendRow([
    new Date(),
    need.name,
    need.email,
    need.salesGroup,
    need.queue,
    need.managerName,
    need.incompleteCount,
    need.sims.join(', '),
    need.durationMin,
    'PENDING',
    '',
    '',
    '',
    '',
    '',
    escalateToken
  ]);
  SpreadsheetApp.flush();
  return sheet.getLastRow();
}

function tsCreateBookingToken_(opts) {
  tsEnsureBookingTokensSheet_(tsGetSpreadsheet_());
  var token = Utilities.getUuid().replace(/-/g, '');
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.BOOKING_TOKENS);
  sheet.appendRow([
    token,
    new Date(),
    'PENDING',
    opts.consultantName || '',
    opts.consultantEmail || '',
    opts.managerName || '',
    opts.salesGroup || '',
    opts.queue || '',
    opts.simsCsv || '',
    opts.durationMin || 30,
    opts.dateStr || '',
    opts.startStr || '',
    opts.endStr || '',
    opts.offerRow || ''
  ]);
  var lr = sheet.getLastRow();
  var rDt = sheet.getRange(lr, TS.BOOK_COLS.DATE_STR, 1, 3);
  rDt.setNumberFormat('@');
  rDt.setValues([[String(opts.dateStr || ''), String(opts.startStr || ''), String(opts.endStr || '')]]);
  SpreadsheetApp.flush();
  return token;
}

function tsLookupBookingTokenRow_(token) {
  var t = String(token || '').trim().toLowerCase();
  if (!t) return null;
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.BOOKING_TOKENS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === t) {
      return { rowNum: i + 1, row: values[i] };
    }
  }
  return null;
}

function tsUpdateBookingTokenStatus_(rowNum, status) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.BOOKING_TOKENS);
  if (!sheet || !rowNum) return;
  sheet.getRange(rowNum, TS.BOOK_COLS.STATUS).setValue(status);
  SpreadsheetApp.flush();
}

/***************************************
 * CONFIG + UTIL
 ***************************************/

/** Web app doGet has no active spreadsheet — always open container by ID. */
function tsGetSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (ignore) {}
  try {
    return SpreadsheetApp.openById(TS.SPREADSHEET_ID);
  } catch (err) {
    throw new Error('Cannot open workbook: ' + err);
  }
}

function tsLoadConfig_() {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) return {};
  var values = sheet.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) continue;
    if (key === 'Manager Name') break;
    out[key] = values[i][1];
  }
  return out;
}

function tsConfigBool_(config, key, defaultVal) {
  var val = config[key];
  if (val === undefined || val === null || String(val).trim() === '') return defaultVal;
  return String(val).trim().toUpperCase() === 'TRUE';
}

function tsIsTestMode_() {
  return tsConfigBool_(tsLoadConfig_(), 'TEST_MODE', false);
}

function tsLoadManagerSlackAliasMap_() {
  var map = {};
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) return map;
  var values = sheet.getDataRange().getValues();
  var inSection = false;
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0] || '').trim();
    var alias = String(values[i][1] || '').trim();
    if (name === 'Manager Name') { inSection = true; continue; }
    if (inSection) {
      if (!name) break;
      if (name.indexOf('(no managers found') === 0) continue;
      map[name] = alias;
    }
  }
  return map;
}

function tsResolveManagerSlackAlias_(managerName, aliasMap) {
  if (normalizeManagerDisplayName_(managerName) === 'Liz Berkley') return 'elizabeth.berkley';
  if (!managerName) return '';
  if (aliasMap[managerName]) return aliasMap[managerName];
  var key = tsNormalizePersonName_(managerName);
  var lower = managerName.toLowerCase();
  for (var exactKey in aliasMap) {
    if (exactKey.toLowerCase() === lower) return aliasMap[exactKey];
    if (tsNormalizePersonName_(exactKey) === key && aliasMap[exactKey]) return aliasMap[exactKey];
  }
  var parts = key.split(' ').filter(Boolean);
  if (parts.length >= 2) {
    var fuzzy = [];
    for (var name in aliasMap) {
      if (!aliasMap[name]) continue;
      var mk = tsNormalizePersonName_(name);
      if (mk === key || mk.indexOf(key) === 0 || key.indexOf(mk) === 0) fuzzy.push(aliasMap[name]);
    }
    if (fuzzy.length === 1) return fuzzy[0];
  }
  return '';
}

function tsSlackDmManager_(managerName, message) {
  if (!managerName) return;
  var aliasMap = tsLoadManagerSlackAliasMap_();
  var alias = tsResolveManagerSlackAlias_(managerName, aliasMap) || managerName;
  if (tsIsTestMode_()) {
    tsSlackDmTestRecipient_('*[TEST MODE — manager ' + managerName + ']*\n\n' + message);
    return;
  }
  tsSlackDmAlias_(alias, message);
}

function tsAudit_(event, ref, details, result) {
  try {
    var sheet = tsEnsureAuditSheet_(tsGetSpreadsheet_());
    sheet.appendRow([new Date(), event, ref, details, result]);
  } catch (ignore) {}
}

function tsRemoveTrigger_(fnName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
  });
}

function tsGetApiKey_() {
  var key = (PropertiesService.getScriptProperties().getProperty('ASSEMBLED_API_KEY') || '').trim();
  if (!key) key = String(tsLoadConfig_().ASSEMBLED_API_KEY || '').trim();
  if (!key) throw new Error('ASSEMBLED_API_KEY not set (Script Properties or TS Config).');
  return key;
}

function tsGetSlackToken_() {
  var token = (PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || '').trim();
  if (!token) token = String(tsLoadConfig_().SLACK_BOT_TOKEN || '').trim();
  if (!token) throw new Error('SLACK_BOT_TOKEN not set (Script Properties or TS Config).');
  return token;
}

function tsAuthHeaders_(apiKey) {
  return {
    Authorization: 'Basic ' + Utilities.base64Encode(apiKey + ':'),
    'Content-Type': 'application/json'
  };
}

function tsBuildUrl_(base, params) {
  var keys = Object.keys(params || {}).filter(function(k) {
    return params[k] !== undefined && params[k] !== null && params[k] !== '';
  });
  if (!keys.length) return base;
  return base + '?' + keys.map(function(k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k]));
  }).join('&');
}

function tsBuildDateTime_(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  var iso = String(dateStr).trim() + 'T' + String(timeStr).trim() + ':00';
  var d = Utilities.parseDate(iso, TS.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  return (!d || isNaN(d.getTime())) ? null : d;
}

function tsParseDateStr_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, TS.TZ, 'yyyy-MM-dd');
  }
  var s = String(raw || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  return s;
}

function tsParseTimeInt_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, TS.TZ, 'HH:mm');
  }
  var s = String(raw || '').trim();
  if (/^\d{1,2}:\d{2}/.test(s)) {
    var parts = s.split(':');
    return tsPad2_(parseInt(parts[0], 10)) + ':' + tsPad2_(parseInt(parts[1], 10));
  }
  return s;
}

function tsAlignDown_(n, mod) { return n - (n % mod); }
function tsPad2_(n) { var x = Number(n); return x < 10 ? '0' + x : String(x); }
function tsNum_(v) { var n = Number(v); return isFinite(n) ? n : 0; }
function tsIsNum_(v) { return v !== null && v !== undefined && v !== '' && isFinite(Number(v)); }
function tsNormalizeToken_(s) {
  return String(s || '').toLowerCase().trim()
    .replace(/[^a-z0-9]+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '');
}
function tsFriendlyDate_(dateStr) {
  var d = tsBuildDateTime_(dateStr, '12:00');
  return d ? Utilities.formatDate(d, TS.TZ, 'EEE, MMM d') : dateStr;
}
function tsFriendlyTime_(timeStr) {
  var d = tsBuildDateTime_('2000-01-01', timeStr);
  return d ? Utilities.formatDate(d, TS.TZ, 'h:mm a') : timeStr;
}
function tsEscHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
