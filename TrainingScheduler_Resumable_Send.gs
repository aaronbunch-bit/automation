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
var TS_RESUME_RUN_BUDGET_MS = 4.25 * 60 * 1000;

function sendItSendResumable() {
  var needs = tsLoadConsultantsNeedingTraining_();

  if (!needs.length) {
    tsClearResumeBatchState_();
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

  var state = tsLoadResumeBatchState_();
  var processed = state.processed;
  var startedAt = state.startedAt || new Date().toISOString();
  var runStart = Date.now();
  var processedThisRun = 0;
  var skippedAlreadyProcessed = 0;
  var total = needs.length;

  tsSaveResumeBatchState_(startedAt, processed);

  for (var i = 0; i < needs.length; i++) {
    if (Date.now() - runStart > TS_RESUME_RUN_BUDGET_MS) {
      tsSaveResumeBatchState_(startedAt, processed);
      tsAudit_(
        'SEND_IT_RESUME',
        '',
        'Paused after processing ' + processedThisRun + ' this run; ' +
          Object.keys(processed).length + '/' + total + ' total processed.',
        'WARN'
      );
      SpreadsheetApp.getUi().alert(
        'Send',
        'Paused before Apps Script timeout.\n\n' +
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

    processTrainingOffers({
      force: true,
      consultantFilter: need.email || need.name
    });

    processed[key] = true;
    processedThisRun++;
    tsSaveResumeBatchState_(startedAt, processed);
  }

  tsClearResumeBatchState_();
  tsAudit_(
    'SEND_IT_RESUME',
    '',
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
    'Training offer batch complete.\n\n' +
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
  var raw = String(props.getProperty(TS_RESUME_BATCH_PROCESSED_KEYS) || '{}').trim();
  var processed = {};

  try {
    processed = JSON.parse(raw) || {};
  } catch (ignore) {
    processed = {};
  }

  return {
    startedAt: startedAt,
    processed: processed
  };
}

function tsSaveResumeBatchState_(startedAt, processed) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(TS_RESUME_BATCH_STARTED_AT_KEY, startedAt || new Date().toISOString());
  props.setProperty(TS_RESUME_BATCH_PROCESSED_KEYS, JSON.stringify(processed || {}));
}

function tsClearResumeBatchState_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(TS_RESUME_BATCH_STARTED_AT_KEY);
  props.deleteProperty(TS_RESUME_BATCH_PROCESSED_KEYS);
}

function tsResumeNeedKey_(need) {
  return [
    String(need.email || '').trim().toLowerCase(),
    String(need.salesGroup || '').trim().toLowerCase()
  ].join('|');
}
