/**
 * Weekly assigned scheduling mode for Reflex-AI Training Scheduler.
 *
 * Uses a new tab:
 *   Weekly Sim Schedule
 *
 * Columns:
 *   Week Start | Supergroup | Simulation Name
 *
 * Behavior:
 * - Schedules one configured sim per rep per week.
 * - Uses a fixed 20-minute block by default.
 * - Chooses the best on-shift window by capacity score, even when staffing is short.
 * - In test mode, sends the rep-facing message to aaron.bunch and does not book.
 */

var WEEKLY_SIM_SCHEDULE_SHEET_NAME = 'Weekly Sim Schedule';
var WEEKLY_SIM_SCHEDULE_HEADERS = ['Week Start', 'Supergroup', 'Simulation Name'];

function setupWeeklyAssignedScheduling() {
  var ss = tsGetSpreadsheet_();
  var sheet = ss.getSheetByName(WEEKLY_SIM_SCHEDULE_SHEET_NAME) || ss.insertSheet(WEEKLY_SIM_SCHEDULE_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, WEEKLY_SIM_SCHEDULE_HEADERS.length).setValues([WEEKLY_SIM_SCHEDULE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, WEEKLY_SIM_SCHEDULE_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1F4E78')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, WEEKLY_SIM_SCHEDULE_HEADERS.length);
  }

  tsEnsureWeeklyConfigKeys_();
  SpreadsheetApp.getUi().alert(
    'Weekly assigned scheduling setup complete.\n\n' +
    'Fill Weekly Sim Schedule with:\n' +
    'Week Start | Supergroup | Simulation Name'
  );
}

function testWeeklyAssignedScheduling() {
  runWeeklyAssignedScheduling_(true);
}

function runWeeklyAssignedScheduling() {
  runWeeklyAssignedScheduling_(false);
}

function runWeeklyAssignedScheduling_(testMode) {
  var ss = tsGetSpreadsheet_();
  var config = tsLoadConfig_();
  var weekStart = tsWeeklyTargetWeekStart_(new Date());
  var weekStartKey = Utilities.formatDate(weekStart, TS.TZ, 'yyyy-MM-dd');
  var assignments = tsLoadWeeklyAssignments_(ss, weekStartKey);

  if (!Object.keys(assignments).length) {
    SpreadsheetApp.getUi().alert(
      'Weekly Assigned Scheduling',
      'No Weekly Sim Schedule rows found for week start ' + weekStartKey + '.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var needs = tsLoadConsultantsNeedingTraining_();
  var eligibleNeeds = [];

  needs.forEach(function(need) {
    var assignedSim = assignments[String(need.salesGroup || '').trim().toLowerCase()];
    if (!assignedSim) return;

    var sims = (need.sims || []).map(function(value) { return String(value || '').trim().toLowerCase(); });
    if (sims.indexOf(String(assignedSim).trim().toLowerCase()) === -1) return;

    var weeklyNeed = {};
    Object.keys(need).forEach(function(key) {
      weeklyNeed[key] = need[key];
    });
    weeklyNeed.sims = [assignedSim];
    weeklyNeed.incompleteCount = 1;
    weeklyNeed.durationMin = Number(config.WEEKLY_TRAINING_DURATION_MIN || 20);
    weeklyNeed.weekStart = weekStartKey;
    eligibleNeeds.push(weeklyNeed);
  });

  if (!eligibleNeeds.length) {
    SpreadsheetApp.getUi().alert(
      'Weekly Assigned Scheduling',
      'No reps need the assigned weekly simulations for week start ' + weekStartKey + '.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var apiKey = tsGetApiKey_();
  var headers = tsAuthHeaders_(apiKey);
  var siteId = tsResolveSiteId_(headers, TS.ASSEMBLED.SITE_NAME);
  var range = tsWeeklyRange_(weekStart);
  var scheduleIdx = tsPullPhoneScheduleIndex_(headers, range.start, range.end);
  var queueCache = {};
  var processed = 0;
  var booked = 0;
  var testSent = 0;
  var escalated = 0;
  var skippedAlreadyScheduled = 0;
  var failed = 0;

  eligibleNeeds.forEach(function(need) {
    try {
      if (tsWeeklyAlreadyScheduled_(need)) {
        skippedAlreadyScheduled++;
        return;
      }

      var queueId = queueCache[need.queue] && queueCache[need.queue].queueId;
      if (!queueId) {
        queueId = tsResolveQueueId_(headers, need.queue);
      }
      if (!queueId) {
        tsAudit_('WEEKLY_ASSIGNED', need.email, 'Queue not found: ' + need.queue, 'FAILED');
        failed++;
        return;
      }

      if (!queueCache[need.queue]) {
        queueCache[need.queue] = {
          queueId: queueId,
          forecastCache: tsBuildWeeklyForecastCache_(headers, siteId, queueId, range, config)
        };
      }

      var bestWindow = tsFindBestWeeklyTrainingWindow_(
        queueCache[need.queue].forecastCache,
        scheduleIdx,
        need.email,
        Number(need.durationMin || 20) * 60 * 1000,
        config
      );

      if (!bestWindow) {
        if (typeof tsAutoEscalateNoCapacity_ === 'function' && !testMode) {
          tsAutoEscalateNoCapacity_(need, 'No on-shift weekly assigned window found', config);
          escalated++;
        } else {
          tsAudit_('WEEKLY_ASSIGNED', need.email, 'No on-shift weekly assigned window found', 'WARN');
          failed++;
        }
        return;
      }

      if (testMode) {
        tsSendWeeklyAssignedTestMessage_(need, bestWindow);
        testSent++;
      } else {
        tsBookWeeklyAssignedWindow_(headers, need, bestWindow);
        booked++;
      }
      processed++;
    } catch (err) {
      failed++;
      tsAudit_('WEEKLY_ASSIGNED', need.email || need.name || '', String(err), 'FAILED');
    }
  });

  SpreadsheetApp.getUi().alert(
    'Weekly Assigned Scheduling',
    (testMode ? 'Test run complete.\n\n' : 'Live run complete.\n\n') +
      'Eligible reps: ' + eligibleNeeds.length + '\n' +
      'Processed: ' + processed + '\n' +
      'Booked: ' + booked + '\n' +
      'Test messages sent: ' + testSent + '\n' +
      'Escalated/no-window: ' + escalated + '\n' +
      'Skipped already scheduled: ' + skippedAlreadyScheduled + '\n' +
      'Failed: ' + failed,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function tsLoadWeeklyAssignments_(ss, weekStartKey) {
  var sheet = ss.getSheetByName(WEEKLY_SIM_SCHEDULE_SHEET_NAME);
  var assignments = {};
  if (!sheet || sheet.getLastRow() < 2) return assignments;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(value) { return String(value || '').trim(); });
  var col = {};
  headers.forEach(function(header, index) {
    col[header] = index;
  });

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowWeekStart = tsWeeklyDateKey_(row[col['Week Start']]);
    var supergroup = String(row[col['Supergroup']] || '').trim();
    var simulationName = String(row[col['Simulation Name']] || '').trim();

    if (!rowWeekStart || !supergroup || !simulationName) continue;
    if (rowWeekStart !== weekStartKey) continue;

    assignments[supergroup.toLowerCase()] = simulationName;
  }

  return assignments;
}

function tsWeeklyTargetWeekStart_(today) {
  var d = new Date(today.getTime());
  d.setHours(0, 0, 0, 0);
  var day = d.getDay();
  var daysUntilSunday = (7 - day) % 7;
  if (daysUntilSunday === 0 && day !== 0) daysUntilSunday = 7;
  d.setDate(d.getDate() + daysUntilSunday);
  return d;
}

function tsWeeklyRange_(weekStart) {
  var start = new Date(weekStart.getTime());
  start.setHours(0, 0, 0, 0);
  var end = new Date(start.getTime());
  end.setDate(end.getDate() + 7);
  end.setHours(0, 0, 0, 0);
  return { start: start, end: end };
}

function tsBuildWeeklyForecastCache_(headers, siteId, queueId, range, config) {
  var cache = {};
  var cursor = new Date(range.start.getTime());

  while (cursor < range.end) {
    var dateStr = Utilities.formatDate(cursor, TS.TZ, 'yyyy-MM-dd');
    var bandStart = tsBuildDateTime_(dateStr, tsPad2_(Number(config.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START)) + ':00');
    var bandEnd = tsBuildDateTime_(dateStr, tsPad2_(Number(config.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END)) + ':00');

    if (bandStart && bandEnd) {
      var startSec = Math.floor(tsAlignDown_(bandStart.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
      var endSec = Math.floor(tsAlignDown_(bandEnd.getTime() / 1000, TS.ASSEMBLED.INTERVAL));
      cache[dateStr] = tsFetchForecastIntervals_(headers, siteId, queueId, startSec, endSec);
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return cache;
}

function tsFindBestWeeklyTrainingWindow_(forecastCache, scheduleIdx, consultantEmail, durationMs, config) {
  var best = null;
  var bandStartHour = Number(config.BUSINESS_HOUR_START || TS.BUSINESS_HOUR_START);
  var bandEndHour = Number(config.BUSINESS_HOUR_END || TS.BUSINESS_HOUR_END);

  Object.keys(forecastCache).sort().forEach(function(dateStr) {
    var intervals = forecastCache[dateStr] || [];
    if (!intervals.length) return;

    var bandStart = tsBuildDateTime_(dateStr, tsPad2_(bandStartHour) + ':00');
    var bandEnd = tsBuildDateTime_(dateStr, tsPad2_(bandEndHour) + ':00');
    if (!bandStart || !bandEnd) return;

    for (var i = 0; i < intervals.length; i++) {
      var interval = intervals[i];
      if (!interval.start_time) continue;

      var slotStart = new Date(interval.start_time * 1000);
      var slotEnd = new Date(slotStart.getTime() + durationMs);
      if (slotStart < bandStart || slotEnd > bandEnd) continue;
      if (slotStart.getTime() < new Date().getTime() + 2 * 60 * 60 * 1000) continue;
      if (!tsConsultantOnShiftDuringWindow_(scheduleIdx, consultantEmail, slotStart, slotEnd)) continue;

      var minPostNet = Infinity;
      for (var j = i; j < intervals.length; j++) {
        var current = intervals[j];
        if (!current.start_time) continue;
        var currentStart = new Date(current.start_time * 1000);
        if (currentStart >= slotEnd) break;

        var scheduled = tsNum_(current.staffing_scheduled);
        var required = tsNum_(current.staffing_required && current.staffing_required.forecasted);
        var net = tsIsNum_(current.staffing_net) ? Number(current.staffing_net) : scheduled - required;
        var postNet = net - 1;
        if (postNet < minPostNet) minPostNet = postNet;
      }

      if (!isFinite(minPostNet)) continue;
      if (!best || minPostNet > best.postNet || (minPostNet === best.postNet && slotStart < best.start)) {
        best = {
          start: slotStart,
          end: slotEnd,
          dateStr: dateStr,
          startStr: Utilities.formatDate(slotStart, TS.TZ, 'HH:mm'),
          endStr: Utilities.formatDate(slotEnd, TS.TZ, 'HH:mm'),
          postNet: minPostNet
        };
      }
    }
  });

  return best;
}

function tsWeeklyAlreadyScheduled_(need) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var values = sheet.getDataRange().getValues();
  var targetEmail = String(need.email || '').trim().toLowerCase();
  var targetGroup = String(need.salesGroup || '').trim().toLowerCase();
  var targetSim = String((need.sims || [])[0] || '').trim().toLowerCase();
  var weekStart = tsBuildDateTime_(need.weekStart, '00:00');
  var weekEnd = weekStart ? new Date(weekStart.getTime() + 7 * 24 * 60 * 60 * 1000) : null;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var email = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var salesGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var simsCsv = String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (email !== targetEmail || salesGroup !== targetGroup) continue;
    if (status !== 'BOOKED') continue;
    if (simsCsv !== targetSim) continue;
    if (!weekStart || !weekEnd) return true;

    var bookedWindow = String(row[TS.OFFER_COLS.BOOKED_WINDOW - 1] || '').trim();
    var bookedDate = tsBuildDateTime_(bookedWindow.substring(0, 10), '00:00');
    if (bookedDate && bookedDate >= weekStart && bookedDate < weekEnd) return true;
  }

  return false;
}

function tsBookWeeklyAssignedWindow_(headers, need, window) {
  var description = 'Weekly Reflex-AI Sim: ' + need.sims.join(', ');
  var commit = tsCommitTrainingToAssembled_(headers, need.email, need.name, window.start, window.end, description, 'weekly-assigned');
  if (!commit.ok) {
    tsAudit_('WEEKLY_ASSIGNED', need.email, 'Assembled commit failed', 'FAILED');
    return;
  }

  var offerRow = tsAppendOfferRow_(need);
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  sheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('BOOKED');
  sheet.getRange(offerRow, TS.OFFER_COLS.OFFER_SENT_AT).setValue(new Date());
  sheet.getRange(offerRow, TS.OFFER_COLS.BOOKED_AT).setValue(new Date());
  sheet.getRange(offerRow, TS.OFFER_COLS.BOOKED_WINDOW).setValue(window.dateStr + ' ' + window.startStr + '-' + window.endStr);
  SpreadsheetApp.flush();

  tsCreateConsultantCalendarEvent_(need.email, need.name, need.sims.join(', '), need.salesGroup, window.start, window.end, 'weekly-assigned');
  tsSlackDmConsultant_(need.name, need.email, tsBuildWeeklyAssignedSlackMessage_(need, window, false));
  tsAudit_('WEEKLY_ASSIGNED', need.email, 'Booked weekly assigned sim ' + need.sims.join(', ') + ' at ' + window.dateStr + ' ' + window.startStr, 'OK');
}

function tsSendWeeklyAssignedTestMessage_(need, window) {
  var alias = String(tsLoadConfig_().WEEKLY_TEST_SLACK_ALIAS || 'aaron.bunch').trim();
  tsSlackDmAlias_(alias, '*[TEST MODE — would have gone to ' + (need.name || need.email) + ']*\n\n' + tsBuildWeeklyAssignedSlackMessage_(need, window, true));
  tsAudit_('WEEKLY_ASSIGNED_TEST', need.email, 'Sent weekly assigned preview to ' + alias, 'OK');
}

function tsBuildWeeklyAssignedSlackMessage_(need, window, testMode) {
  return [
    '*Reflex-AI weekly training scheduled*',
    '',
    'You have been scheduled for a *20-minute* Reflex-AI training block.',
    '',
    '*Simulation to complete:* ' + need.sims.join(', '),
    '*Time:* ' + tsFriendlyDate_(window.dateStr) + ', ' + tsFriendlyTime_(window.startStr) + '–' + tsFriendlyTime_(window.endStr) + ' CT',
    '',
    'Please complete only this assigned simulation during the scheduled block.'
  ].join('\n');
}

function tsWeeklyDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TS.TZ, 'yyyy-MM-dd');
  }
  var raw = String(value || '').trim();
  if (!raw) return '';
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, TS.TZ, 'yyyy-MM-dd');
  }
  return raw.substring(0, 10);
}

function tsEnsureWeeklyConfigKeys_() {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet) return;
  tsEnsureOptionalConfigKey_(sheet, 'SCHEDULING_MODE', 'AUTO', 'AUTO, WEEKLY_ASSIGNED, or OFFER_BASED');
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_MODE_START_DATE', '2026-08-01', 'Weekly assigned mode start date');
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_MODE_END_DATE', '2026-10-31', 'Weekly assigned mode end date');
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_TRAINING_DURATION_MIN', '20', 'Weekly assigned training block length');
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_TEST_SLACK_ALIAS', 'aaron.bunch', 'Slack alias for weekly assigned test messages');
}
