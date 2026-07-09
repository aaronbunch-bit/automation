/**
 * Resumable Send It runner for the Reflex-AI Training Scheduler.
 *
 * Apps Script has a hard execution limit. This wrapper lets operators click
 * Send It > Send repeatedly and continue from the last processed consultant
 * instead of starting over.
 *
 * Requires the main scheduler file to be present because it uses:
 * - tsLoadConsultantsNeedingTraining_
 * - processTrainingOffers
 * - tsAudit_
 */

var TS_RESUME_BATCH_STARTED_AT_KEY = 'TS_RESUME_BATCH_STARTED_AT';
var TS_RESUME_BATCH_PROCESSED_KEYS = 'TS_RESUME_BATCH_PROCESSED_KEYS';
var TS_RESUME_BATCH_SALES_GROUP_KEY = 'TS_RESUME_BATCH_SALES_GROUP';
var TS_RESUME_RUN_BUDGET_MS = 4.25 * 60 * 1000;

function sendItSendResumable() {
  sendItSendResumableForSalesGroup_('');
}

function sendItSendAdultLearning() {
  sendItSendResumableForSalesGroup_('Adult Learning');
}

function sendItSendAllSupergroups() {
  sendItSendResumableForSalesGroup_('');
}

function sendItSendCollege() {
  sendItSendResumableForSalesGroup_('College');
}

function sendItSendELD() {
  sendItSendResumableForSalesGroup_('ELD');
}

function sendItSendHighSchool() {
  sendItSendResumableForSalesGroup_('High School');
}

function sendItSendProfCerts() {
  sendItSendResumableForSalesGroup_('Prof Certs');
}

function sendItSendResumableForSalesGroup_(salesGroupFilter) {
  var normalizedSalesGroupFilter = String(salesGroupFilter || '').trim();
  var needs = tsLoadConsultantsNeedingTraining_().filter(function(need) {
    return !normalizedSalesGroupFilter ||
      String(need.salesGroup || '').trim().toLowerCase() === normalizedSalesGroupFilter.toLowerCase();
  });
  var batchLabel = normalizedSalesGroupFilter || 'All Supergroups';

  if (!needs.length) {
    tsClearResumeBatchState_();
    SpreadsheetApp.getUi().alert(
      'Send',
      'No consultants with outstanding Reflex-AI training were found for ' + batchLabel + '.',
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

  var state = tsLoadResumeBatchState_();
  if (String(state.salesGroup || '') !== normalizedSalesGroupFilter) {
    state = {
      startedAt: '',
      processed: {},
      salesGroup: normalizedSalesGroupFilter
    };
  }
  var processed = state.processed;
  var startedAt = state.startedAt || new Date().toISOString();
  var runStart = Date.now();
  var processedThisRun = 0;
  var skippedAlreadyProcessed = 0;
  var total = needs.length;
  var processingContext = tsCreateResumeProcessingContext_();

  tsSaveResumeBatchState_(startedAt, processed, normalizedSalesGroupFilter);

  for (var i = 0; i < needs.length; i++) {
    if (Date.now() - runStart > TS_RESUME_RUN_BUDGET_MS) {
      tsSaveResumeBatchState_(startedAt, processed, normalizedSalesGroupFilter);
      tsAudit_(
        'SEND_IT_RESUME',
        batchLabel,
        'Paused after processing ' + processedThisRun + ' this run; ' +
          Object.keys(processed).length + '/' + total + ' total processed.',
        'WARN'
      );
      SpreadsheetApp.getUi().alert(
        'Send',
        'Paused before Apps Script timeout for ' + batchLabel + '.\n\n' +
          'Processed this click: ' + processedThisRun + '\n' +
          'Already completed in this batch: ' + Object.keys(processed).length + ' of ' + total + '\n\n' +
          'Click Send It > Send again to continue from where it left off.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    var need = needs[i];
    var key = tsResumeNeedKey_(need);
    if (processed[key]) {
      skippedAlreadyProcessed++;
      continue;
    }

    tsProcessSingleTrainingNeedForResume_(need, processingContext);

    processed[key] = true;
    processedThisRun++;
    tsSaveResumeBatchState_(startedAt, processed, normalizedSalesGroupFilter);
  }

  tsClearResumeBatchState_();
  tsAudit_(
    'SEND_IT_RESUME',
    batchLabel,
    'Batch complete. Processed this run=' + processedThisRun +
      '; skipped previously processed=' + skippedAlreadyProcessed +
      '; total=' + total,
    'OK'
  );

  var testNote = tsIsTestMode_()
    ? '\n\nTEST_MODE is ON — Slack DMs go to ' +
      (tsLoadConfig_().TEST_MODE_SLACK_EMAIL || 'robert.sorrell@varsitytutors.com') + '.'
    : '';

  SpreadsheetApp.getUi().alert(
    'Send',
    'Training offer batch complete for ' + batchLabel + '.\n\n' +
      'Processed this click: ' + processedThisRun + '\n' +
      'Skipped previously processed: ' + skippedAlreadyProcessed + '\n' +
      'Total consultants in batch: ' + total + testNote,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function resetSendItResumeBatch() {
  tsClearResumeBatchState_();
  SpreadsheetApp.getUi().alert(
    'Send',
    'Send It resume state cleared. The next Send It run will start a fresh batch.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function tsLoadResumeBatchState_() {
  var props = PropertiesService.getScriptProperties();
  var startedAt = String(props.getProperty(TS_RESUME_BATCH_STARTED_AT_KEY) || '').trim();
  var salesGroup = String(props.getProperty(TS_RESUME_BATCH_SALES_GROUP_KEY) || '').trim();
  var raw = String(props.getProperty(TS_RESUME_BATCH_PROCESSED_KEYS) || '{}').trim();
  var processed = {};

  try {
    processed = JSON.parse(raw) || {};
  } catch (ignore) {
    processed = {};
  }

  return {
    startedAt: startedAt,
    salesGroup: salesGroup,
    processed: processed
  };
}

function tsSaveResumeBatchState_(startedAt, processed, salesGroup) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(TS_RESUME_BATCH_STARTED_AT_KEY, startedAt || new Date().toISOString());
  props.setProperty(TS_RESUME_BATCH_SALES_GROUP_KEY, String(salesGroup || '').trim());
  props.setProperty(TS_RESUME_BATCH_PROCESSED_KEYS, JSON.stringify(processed || {}));
}

function tsClearResumeBatchState_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(TS_RESUME_BATCH_STARTED_AT_KEY);
  props.deleteProperty(TS_RESUME_BATCH_SALES_GROUP_KEY);
  props.deleteProperty(TS_RESUME_BATCH_PROCESSED_KEYS);
}

function tsResumeNeedKey_(need) {
  return [
    String(need.email || '').trim().toLowerCase(),
    String(need.salesGroup || '').trim().toLowerCase()
  ].join('|');
}

function tsProcessSingleTrainingNeedForResume_(need, processingContext) {
  if (tsConsultantAlreadyBooked_(need.email, need.salesGroup)) {
    tsAudit_('SEND_IT_RESUME', tsResumeNeedKey_(need), 'Already booked — skip', 'INFO');
    return;
  }

  if (tsHasActivePendingOffer_(need.email, need.salesGroup)) {
    tsAudit_('SEND_IT_RESUME', tsResumeNeedKey_(need), 'Active pending offer/tokens — skip', 'INFO');
    return;
  }

  try {
    var headers = processingContext.headers;
    var config = processingContext.config;
    var siteId = processingContext.siteId;
    var queueId = processingContext.queueIds[need.queue] || tsResolveQueueId_(headers, need.queue);

    if (!queueId) {
      tsAudit_('SEND_IT_RESUME', need.email, 'Queue not found: ' + need.queue, 'FAILED');
      return;
    }

    processingContext.queueIds[need.queue] = queueId;

    if (!processingContext.forecastCaches[need.queue]) {
      processingContext.forecastCaches[need.queue] = tsBuildForecastCache_(headers, siteId, queueId, config);
    }

    var ctx = {
      forecastCache: processingContext.forecastCaches[need.queue],
      scheduleIdx: processingContext.scheduleIdx
    };

    tsProcessOneConsultantOffer_(need, headers, siteId, config, false, { force: true }, ctx);
  } catch (err) {
    tsAudit_('SEND_IT_RESUME', need.email || need.name || '', String(err), 'FAILED');
  }
}

function tsCreateResumeProcessingContext_() {
  var config = tsLoadConfig_();
  var apiKey = tsGetApiKey_();
  var headers = tsAuthHeaders_(apiKey);
  var siteId = tsResolveSiteId_(headers, TS.ASSEMBLED.SITE_NAME);
  var searchRange = tsComputeSearchRange_(config);

  return {
    config: config,
    headers: headers,
    siteId: siteId,
    scheduleIdx: tsPullPhoneScheduleIndex_(headers, searchRange.start, searchRange.end),
    queueIds: {},
    forecastCaches: {}
  };
}
