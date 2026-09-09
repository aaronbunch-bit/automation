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
  runWeeklyAssignedScheduling_(true, '', 'NEXT');
}

function runWeeklyAssignedScheduling() {
  runWeeklyAssignedScheduling_(false, '', 'NEXT');
}

function runWeeklyAssignedSchedulingAllSupergroups() {
  runWeeklyAssignedScheduling_(false, '', 'NEXT');
}

function runWeeklyAssignedSchedulingAdultLearning() {
  runWeeklyAssignedScheduling_(false, 'Adult Learning', 'NEXT');
}

function runWeeklyAssignedSchedulingCollege() {
  runWeeklyAssignedScheduling_(false, 'College', 'NEXT');
}

function runWeeklyAssignedSchedulingELD() {
  runWeeklyAssignedScheduling_(false, 'ELD', 'NEXT');
}

function runWeeklyAssignedSchedulingHighSchool() {
  runWeeklyAssignedScheduling_(false, 'High School', 'NEXT');
}

function runWeeklyAssignedSchedulingProfCerts() {
  runWeeklyAssignedScheduling_(false, 'Prof Certs', 'NEXT');
}

function runCurrentWeeklyAssignedSchedulingAllSupergroups() {
  runWeeklyAssignedScheduling_(false, '', 'CURRENT');
}

function runCurrentWeeklyAssignedSchedulingAdultLearning() {
  runWeeklyAssignedScheduling_(false, 'Adult Learning', 'CURRENT');
}

function runCurrentWeeklyAssignedSchedulingCollege() {
  runWeeklyAssignedScheduling_(false, 'College', 'CURRENT');
}

function runCurrentWeeklyAssignedSchedulingELD() {
  runWeeklyAssignedScheduling_(false, 'ELD', 'CURRENT');
}

function runCurrentWeeklyAssignedSchedulingHighSchool() {
  runWeeklyAssignedScheduling_(false, 'High School', 'CURRENT');
}

function runCurrentWeeklyAssignedSchedulingProfCerts() {
  runWeeklyAssignedScheduling_(false, 'Prof Certs', 'CURRENT');
}

function runWeeklyAssignedScheduling_(testMode, salesGroupFilter, weekMode) {
  var ss = tsGetSpreadsheet_();
  var config = tsLoadConfig_();
  var normalizedWeekMode = String(weekMode || 'NEXT').toUpperCase() === 'CURRENT' ? 'CURRENT' : 'NEXT';
  var weekStart = normalizedWeekMode === 'CURRENT'
    ? tsWeeklyCurrentWeekStart_(new Date())
    : tsWeeklyTargetWeekStart_(new Date());
  var weekStartKey = Utilities.formatDate(weekStart, TS.TZ, 'yyyy-MM-dd');
  var assignments = tsLoadWeeklyAssignments_(ss, weekStartKey);
  var normalizedSalesGroupFilter = String(salesGroupFilter || '').trim();
  var normalizedSalesGroupKey = tsWeeklySalesGroupKey_(normalizedSalesGroupFilter);
  var batchLabel = normalizedSalesGroupFilter || 'All Supergroups';
  var weekModeLabel = normalizedWeekMode === 'CURRENT' ? 'Current Week' : 'Next Week';

  if (!Object.keys(assignments).length) {
    SpreadsheetApp.getUi().alert(
      'Weekly Assigned Scheduling',
      'No Weekly Sim Schedule rows found for ' + weekModeLabel + ' week start ' + weekStartKey + '.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  if (normalizedSalesGroupFilter && !assignments[normalizedSalesGroupKey]) {
    SpreadsheetApp.getUi().alert(
      'Weekly Assigned Scheduling',
      'No Weekly Sim Schedule row found for ' + normalizedSalesGroupFilter + ' and ' + weekModeLabel + ' week start ' + weekStartKey + '.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var needs = tsLoadConsultantsNeedingTraining_();
  var eligibleNeeds = [];
  var groupNeedCount = 0;
  var assignedSimNeedCount = 0;
  var simMismatchSamples = [];

  needs.forEach(function(need) {
    var needSalesGroupKey = tsWeeklySalesGroupKey_(need.salesGroup);
    if (normalizedSalesGroupKey && needSalesGroupKey !== normalizedSalesGroupKey) return;
    groupNeedCount++;

    var assignedSim = assignments[needSalesGroupKey];
    if (!assignedSim) return;
    assignedSimNeedCount++;

    if (!tsWeeklySimulationMatches_(assignedSim, need.sims || [])) {
      if (simMismatchSamples.length < 5) {
        simMismatchSamples.push((need.name || need.email || 'Unknown rep') + ': ' + (need.sims || []).join(' | '));
      }
      return;
    }

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
    tsAudit_(
      'WEEKLY_ASSIGNED_DIAG',
      batchLabel,
      'No eligible weekly reps for ' + weekModeLabel + ' week ' + weekStartKey +
        '; assignedSim=' + (assignments[normalizedSalesGroupKey] || 'Multiple/none') +
        '; loadedNeeds=' + needs.length +
        '; groupNeeds=' + groupNeedCount +
        '; needsWithAssignment=' + assignedSimNeedCount +
        '; simMismatchSamples=' + simMismatchSamples.join(' || '),
      'WARN'
    );
    SpreadsheetApp.getUi().alert(
      'Weekly Assigned Scheduling',
      'No reps matched the assigned weekly simulation for ' + batchLabel + ' and ' + weekModeLabel + ' week start ' + weekStartKey + '.\n\n' +
        'Loaded needs in this group: ' + groupNeedCount + '\n' +
        'Assigned sim: ' + (assignments[normalizedSalesGroupKey] || 'See Weekly Sim Schedule') + '\n\n' +
        'Check TS Audit for sample sim names from the CSV.',
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
  var slotOccupancy = tsBuildWeeklySlotOccupancy_(weekStartKey);
  var slotCap = Number(config.WEEKLY_MAX_REPS_PER_SIM_SLOT || 3);
  if (!isFinite(slotCap) || slotCap < 1) slotCap = 3;
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
        tsAudit_(
          'WEEKLY_ASSIGNED',
          need.email,
          'Already scheduled for ' + (need.sims || []).join(', ') + ' in week ' + need.weekStart + ' — skip',
          'INFO'
        );
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
        need,
        need.email,
        Number(need.durationMin || 20) * 60 * 1000,
        config,
        slotOccupancy,
        slotCap
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
        tsIncrementWeeklySlotOccupancy_(slotOccupancy, need, bestWindow);
      } else {
        if (tsBookWeeklyAssignedWindow_(headers, need, bestWindow)) {
          tsIncrementWeeklySlotOccupancy_(slotOccupancy, need, bestWindow);
          booked++;
        } else {
          failed++;
        }
      }
      processed++;
    } catch (err) {
      failed++;
      tsAudit_('WEEKLY_ASSIGNED', need.email || need.name || '', String(err), 'FAILED');
    }
  });

  SpreadsheetApp.getUi().alert(
    'Weekly Assigned Scheduling',
    (testMode ? 'Test run complete for ' : 'Live run complete for ') + batchLabel + ' (' + weekModeLabel + ').\n\n' +
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

    assignments[tsWeeklySalesGroupKey_(supergroup)] = simulationName;
  }

  return assignments;
}

function tsWeeklySalesGroupKey_(value) {
  var raw = String(value || '').trim().toLowerCase();
  var compact = raw.replace(/[^a-z0-9]+/g, '');

  if (!compact) return '';
  if (compact.indexOf('highschool') !== -1) return 'high school';
  if (compact.indexOf('adultlearner') !== -1 || compact.indexOf('adultlearning') !== -1) return 'adult learning';
  if (compact.indexOf('elementary') !== -1 || compact === 'eld' || compact.indexOf('eld') !== -1) return 'eld';
  if (compact.indexOf('college') !== -1) return 'college';
  if (compact.indexOf('profcert') !== -1 || compact.indexOf('professionalcert') !== -1) return 'prof certs';

  return raw;
}

function tsWeeklySimulationMatches_(assignedSim, consultantSims) {
  var target = tsWeeklyNormalizeSimulationName_(assignedSim);
  if (!target) return false;

  return (consultantSims || []).some(function(sim) {
    var candidate = tsWeeklyNormalizeSimulationName_(sim);
    if (!candidate) return false;
    if (candidate === target) return true;

    if (Math.min(candidate.length, target.length) >= 12) {
      if (candidate.indexOf(target) !== -1 || target.indexOf(candidate) !== -1) return true;
    }

    if (tsWeeklyTokenMatch_(target, candidate)) return true;

    var maxLength = Math.max(target.length, candidate.length);
    var distance = typeof editDistance_ === 'function'
      ? editDistance_(target, candidate)
      : tsWeeklyEditDistance_(target, candidate);

    return maxLength >= 20 && distance / maxLength <= 0.12;
  });
}

function tsWeeklyNormalizeSimulationName_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[’‘]/g, "'")
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(reflex|reflexai|ai|simulation|sim)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tsWeeklyTokenMatch_(target, candidate) {
  var targetTokens = tsWeeklyMeaningfulTokens_(target);
  var candidateTokens = tsWeeklyMeaningfulTokens_(candidate);
  if (targetTokens.length < 3 || candidateTokens.length < 3) return false;

  var candidateMap = {};
  candidateTokens.forEach(function(token) {
    candidateMap[token] = true;
  });

  var matched = targetTokens.filter(function(token) {
    return !!candidateMap[token];
  }).length;

  return matched / targetTokens.length >= 0.85 && matched / candidateTokens.length >= 0.7;
}

function tsWeeklyMeaningfulTokens_(normalizedName) {
  var stopWords = {
    the: true,
    and: true,
    for: true,
    with: true,
    your: true,
    you: true,
    a: true,
    an: true,
    to: true,
    of: true,
    in: true,
    on: true
  };

  return String(normalizedName || '')
    .split(' ')
    .filter(function(token) {
      return token && !stopWords[token];
    });
}

function tsWeeklyEditDistance_(left, right) {
  left = String(left || '');
  right = String(right || '');

  var matrix = [];
  for (var i = 0; i <= left.length; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= right.length; j++) {
    matrix[0][j] = j;
  }

  for (i = 1; i <= left.length; i++) {
    for (j = 1; j <= right.length; j++) {
      var cost = left.charAt(i - 1) === right.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  return matrix[left.length][right.length];
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

function tsWeeklyCurrentWeekStart_(today) {
  var d = new Date(today.getTime());
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - d.getDay());
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

function tsBuildWeeklySlotOccupancy_(weekStartKey) {
  var occupancy = {};
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() < 2) return occupancy;

  var values = sheet.getDataRange().getValues();
  var weekStart = tsBuildDateTime_(weekStartKey, '00:00');
  var weekEnd = weekStart ? new Date(weekStart.getTime() + 7 * 24 * 60 * 60 * 1000) : null;
  var activeColumn = typeof tsGetOfferActiveColumn_ === 'function'
    ? tsGetOfferActiveColumn_(sheet)
    : 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (status !== 'BOOKED') continue;
    if (activeColumn && typeof tsOfferRowIsInactive_ === 'function' && tsOfferRowIsInactive_(row, activeColumn)) continue;

    var bookedWindow = String(row[TS.OFFER_COLS.BOOKED_WINDOW - 1] || '').trim();
    var windowParts = tsWeeklyBookedWindowParts_(bookedWindow);
    if (!windowParts.dateStr || !windowParts.startStr) continue;

    var bookedDate = tsBuildDateTime_(windowParts.dateStr, '00:00');
    if (weekStart && weekEnd && (!bookedDate || bookedDate < weekStart || bookedDate >= weekEnd)) continue;

    String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '')
      .split(',')
      .map(function(sim) { return sim.trim(); })
      .filter(Boolean)
      .forEach(function(sim) {
        var key = tsWeeklySlotKey_(sim, windowParts.dateStr, windowParts.startStr);
        occupancy[key] = (occupancy[key] || 0) + 1;
      });
  }

  return occupancy;
}

function tsWeeklyBookedWindowParts_(bookedWindow) {
  var raw = String(bookedWindow || '').trim();
  var match = raw.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}:\d{2})/);
  if (!match) {
    return {
      dateStr: raw.substring(0, 10),
      startStr: ''
    };
  }

  return {
    dateStr: match[1],
    startStr: match[2]
  };
}

function tsWeeklySlotCount_(occupancy, need, dateStr, startStr) {
  return occupancy[tsWeeklySlotKey_((need.sims || [])[0], dateStr, startStr)] || 0;
}

function tsIncrementWeeklySlotOccupancy_(occupancy, need, window) {
  var key = tsWeeklySlotKey_((need.sims || [])[0], window.dateStr, window.startStr);
  occupancy[key] = (occupancy[key] || 0) + 1;
}

function tsWeeklySlotKey_(simulationName, dateStr, startStr) {
  return [
    tsWeeklyNormalizeSimulationName_(simulationName),
    String(dateStr || '').trim(),
    String(startStr || '').trim()
  ].join('|');
}

function tsFindBestWeeklyTrainingWindow_(forecastCache, scheduleIdx, need, consultantEmail, durationMs, config, slotOccupancy, slotCap) {
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
      var slotStartStr = Utilities.formatDate(slotStart, TS.TZ, 'HH:mm');
      if (tsWeeklySlotCount_(slotOccupancy, need, dateStr, slotStartStr) >= slotCap) continue;

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
          startStr: slotStartStr,
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
  var activeColumn = typeof tsGetOfferActiveColumn_ === 'function'
    ? tsGetOfferActiveColumn_(sheet)
    : 0;
  var cutoff = typeof tsGetOfferIgnoreBeforeDate_ === 'function'
    ? tsGetOfferIgnoreBeforeDate_()
    : null;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var email = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var salesGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (email !== targetEmail || salesGroup !== targetGroup) continue;
    if (status !== 'BOOKED') continue;
    if (activeColumn && typeof tsOfferRowIsInactive_ === 'function' && tsOfferRowIsInactive_(row, activeColumn)) continue;
    if (cutoff && typeof tsDateIsBeforeCutoff_ === 'function' && tsDateIsBeforeCutoff_(row[TS.OFFER_COLS.CREATED_AT - 1], cutoff)) continue;
    if (tsWeeklyOfferRowMatchesNeed_(row, need)) return true;
  }

  return false;
}

function tsWeeklyOfferRowMatchesNeed_(row, need) {
  var rowSims = String(row[TS.OFFER_COLS.SIMS_CSV - 1] || '')
    .split(',')
    .map(function(sim) { return String(sim || '').trim().toLowerCase(); })
    .filter(Boolean);
  var targetSims = (need.sims || [])
    .map(function(sim) { return String(sim || '').trim().toLowerCase(); })
    .filter(Boolean);

  var matchesAssignedSim = targetSims.some(function(sim) {
    return rowSims.some(function(rowSim) {
      return rowSim === sim || tsWeeklySimulationMatches_(sim, [rowSim]);
    });
  });
  if (!matchesAssignedSim) return false;

  var weekStart = tsBuildDateTime_(need.weekStart, '00:00');
  if (!weekStart) return false;
  var weekEnd = new Date(weekStart.getTime() + 7 * 24 * 60 * 60 * 1000);
  var bookedWindow = String(row[TS.OFFER_COLS.BOOKED_WINDOW - 1] || '').trim();
  var windowParts = tsWeeklyBookedWindowParts_(bookedWindow);
  var bookedDate = tsBuildDateTime_(windowParts.dateStr || bookedWindow.substring(0, 10), '00:00');
  var bookedStart = windowParts.startStr
    ? tsBuildDateTime_(windowParts.dateStr, windowParts.startStr)
    : bookedDate;

  if (!bookedDate || bookedDate < weekStart || bookedDate >= weekEnd) return false;

  // If the scheduled block has already passed and the rep still appears in the
  // needs list, allow the current-week catch-up run to find a new slot.
  return !!bookedStart && bookedStart.getTime() >= new Date().getTime();
}

function tsBookWeeklyAssignedWindow_(headers, need, window) {
  var description = 'Weekly Reflex-AI Sim: ' + need.sims.join(', ');
  var commit = tsCommitTrainingToAssembled_(headers, need.email, need.name, window.start, window.end, description, 'weekly-assigned');
  if (!commit.ok) {
    tsAudit_('WEEKLY_ASSIGNED', need.email, 'Assembled commit failed', 'FAILED');
    return false;
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
  return true;
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
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_MAX_REPS_PER_SIM_SLOT', '3', 'Maximum reps assigned the same sim at the same date/time');
  tsEnsureOptionalConfigKey_(sheet, 'WEEKLY_TEST_SLACK_ALIAS', 'aaron.bunch', 'Slack alias for weekly assigned test messages');
}
