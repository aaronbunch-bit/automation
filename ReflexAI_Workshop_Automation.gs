var CSV_DUMP_SHEET_NAME = 'ReflexAI CSV Dump';
var CSV_DUMP_SHEET_PREFIX = 'ReflexAI CSV - ';
var MANAGER_ROSTER_SHEET_NAME = 'Manager Roster';
var LOOKER_MANAGER_LOOKUP_SHEET_NAME = 'Looker Manager Lookup';
var NAME_MATCH_OVERRIDES_SHEET_NAME = 'Name Match Overrides';
var REMOVED_REPS_SHEET_NAME = 'Removed Reps';
var MISSING_LOOKER_PAIRINGS_SHEET_NAME = 'Missing Looker Manager Pairing';
var EXCEPTION_SHEET_NAME = 'Simulation Exceptions';
var RUN_LOG_SHEET_NAME = 'Run Log';
var RUN_SETTINGS_SHEET_NAME = 'Run Settings';
var LOOKER_MANAGER_SOURCE_SPREADSHEET_ID = '1Mj5bLfdC3fVhn41I82g15CpoPL-pA-CpMp97FTctfzc';
var LOOKER_MANAGER_SOURCE_RANGE = 'B:D';

var PASSING_SCORE_PERCENT = 80;
var DEFAULT_JOURNEY_NAME = 'High School Year Round Workshops - Jun';

var COMPLETE_SIMULATION_ACTION = 'Ask representative to complete this simulation and schedule time via Assembled for representative to complete the simulation adhering to capacity constraints.';
var RETAKE_SIMULATION_ACTION = 'Ask representative to retake this simulation and coach on missed skills.';
var NO_FOLLOW_UP_ACTION = 'No follow-up required.';
var REFLEXAI_PLATFORM_RESOURCE_URL = 'https://drive.google.com/file/d/18X5z6iRGRk-fKY4bAvIxswwys3ne2z3r/view';
var REFLEXAI_LOGIN_URL = 'https://varsitytutors.reflexai.com/home';
var EMAIL_SUBTITLE = 'Varsity Tutors Quality Assurance Pillar';
var COMPLETED_CLEARED_LABEL = 'Completed & Cleared 80% Threshold';
var COMPLETED_NOT_CLEARED_LABEL = 'Completed & Not Cleared 80% Threshold';
var NOT_STARTED_LABEL = 'Not Started';
var COMPLETED_CLEARED_COLOR = '#4f8cff';
var COMPLETED_NOT_CLEARED_COLOR = '#ffcc33';
var NOT_STARTED_COLOR = '#ff3ec8';
var TAKE_RATE_BOOKED_COLOR = '#4f8cff';
var TAKE_RATE_ESCALATED_COLOR = '#ffcc33';
var TAKE_RATE_OFFERED_COLOR = '#ff3ec8';

var EMAIL_NAME_OVERRIDES = {
  'angie damon': 'angela.damon@varsitytutors.com',
  'angela damon': 'angela.damon@varsitytutors.com',
  'john wright ii': 'john.wright@varsitytutors.com',
  'jennifer volugamore': 'jen.vulgamore@varsitytutors.com',
  'jennifer vulgamore': 'jen.vulgamore@varsitytutors.com'
};

var TEST_EMAIL_RECIPIENTS = [
  'aaron.bunch@varsitytutors.com'
];

var DIRECTOR_EMAIL_RECIPIENTS = [
  'joshua.langford@varsitytutors.com',
  'aaron.bunch@varsitytutors.com',
  'aftynn.peters@varsitytutors.com',
  'taylor.wisnasky@varsitytutors.com'
];

var SENIOR_LEADER_SUPERGROUP_RECIPIENTS = [
  {
    supergroupName: 'Prof Certs',
    leaderName: 'Taylor Wisnasky',
    email: 'taylor.wisnasky@varsitytutors.com'
  },
  {
    supergroupName: 'ELD',
    leaderName: 'Aftynn Peters',
    email: 'aftynn.peters@varsitytutors.com'
  },
  {
    supergroupName: 'High School',
    leaderName: 'Aaron Bunch',
    email: 'aaron.bunch@varsitytutors.com'
  },
  {
    supergroupName: 'College',
    leaderName: 'Joshua Langford',
    email: 'joshua.langford@varsitytutors.com'
  },
  {
    supergroupName: 'Adult Learning',
    leaderName: 'Joshua Langford',
    email: 'joshua.langford@varsitytutors.com'
  }
];

var BULK_CSV_SHEET_NAMES = [
  'ReflexAI CSV - High School',
  'ReflexAI CSV - ELD',
  'ReflexAI CSV - College',
  'ReflexAI CSV - Adult Learning',
  'ReflexAI CSV - Prof Certs'
];

var CSV_DUMP_HEADERS = [
  'User Name',
  'User Email',
  'Simulation Name',
  'Status',
  'Best Score (%)',
  'Passing Score (%)',
  'Passed',
  'Attempts Used',
  'Max Attempts',
  'Completed At',
  'Journey Completion'
];

var LOOKER_MANAGER_LOOKUP_HEADERS = [
  'Rep Name',
  'Rep Id',
  'Rep Manager'
];

var MISSING_LOOKER_PAIRINGS_HEADERS = [
  'Representative'
];

var NAME_MATCH_OVERRIDES_HEADERS = [
  'ReflexAI Name',
  'Looker Name'
];

var REMOVED_REPS_HEADERS = [
  'Representative Email',
  'Representative Name',
  'Reason',
  'Notes'
];

var EXCEPTION_HEADERS = [
  'Report Run At',
  'Representative',
  'Representative Email',
  'Manager',
  'Manager Email',
  'Senior / Team Lead',
  'Senior / Team Lead Email',
  'Journey',
  'Simulation',
  'Completion Status',
  'Score',
  'Required Follow-up Action',
  'Manager Email Sent',
  'Manager Email Sent At',
  'Test Sent'
];

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('ReflexAI Emails')
    .addSubMenu(
      ui.createMenu('Setup')
        .addItem('Set Up Sheets', 'createSetupSheets')
        .addItem('Connect Looker Report', 'connectLookerManagerImport')
        .addItem('Set Current Journey Name', 'setCurrentJourneyName')
    )
    .addSubMenu(
      ui.createMenu('Build Report')
        .addItem('Run CSV Dump', 'runWeeklySimulationExceptionReport')
    )
    .addSubMenu(
      ui.createMenu('Test Emails')
        .addItem('Test Manager Email', 'sendTestManagerExceptionEmails')
        .addItem('Test Senior Leader Email', 'sendTestSeniorLeadershipRecap')
        .addItem('Test Director Email', 'sendTestDirectorEmail')
    )
    .addSubMenu(
      ui.createMenu('Send Emails')
        .addItem('Send Manager Email', 'sendManagerExceptionEmails')
        .addItem('Send Senior Leader Email', 'sendSeniorLeadershipRecap')
        .addItem('Send Director Email', 'sendDirectorEmail')
    )
    .addToUi();

  ui.createMenu('ReflexAI Scheduling')
    .addSubMenu(
      ui.createMenu('Send Offers')
        .addItem('Setup Weekly Assigned Scheduling', 'setupWeeklyAssignedScheduling')
        .addItem('Test Weekly Assigned Scheduling', 'testWeeklyAssignedScheduling')
        .addItem('Run Weekly Assigned - All Supergroups', 'runWeeklyAssignedSchedulingAllSupergroups')
        .addItem('Run Weekly Assigned - Adult Learning', 'runWeeklyAssignedSchedulingAdultLearning')
        .addItem('Run Weekly Assigned - College', 'runWeeklyAssignedSchedulingCollege')
        .addItem('Run Weekly Assigned - ELD', 'runWeeklyAssignedSchedulingELD')
        .addItem('Run Weekly Assigned - High School', 'runWeeklyAssignedSchedulingHighSchool')
        .addItem('Run Weekly Assigned - Prof Certs', 'runWeeklyAssignedSchedulingProfCerts')
        .addSeparator()
        .addItem('Send All Supergroups', 'sendItSendAllSupergroups')
        .addItem('Send Adult Learning', 'sendItSendAdultLearning')
        .addItem('Send College', 'sendItSendCollege')
        .addItem('Send ELD', 'sendItSendELD')
        .addItem('Send High School', 'sendItSendHighSchool')
        .addItem('Send Prof Certs', 'sendItSendProfCerts')
        .addItem('Reset Send Progress', 'resetSendItResumeBatch')
    )
    .addSubMenu(
      ui.createMenu('Manager + Coach Nudges')
        .addItem('Setup Manager Coach Columns', 'setupManagerCoachColumns')
        .addItem('Test Nudge Managers + Coaches', 'testNudgeManagersAndCoaches')
        .addItem('Nudge Managers + Coaches', 'nudgeManagersAndCoaches')
    )
    .addSubMenu(
      ui.createMenu('Repairs')
        .addItem('Repair TS Config Sheet', 'repairTsConfigSheet')
        .addItem('Repair Escalation Links', 'repairEscalationLinks')
        .addItem('Resend Active Offers with Repaired Links', 'resendActiveOffersWithCurrentLinks')
        .addItem('Expire Old Pending Offers', 'expireOldPendingOffers')
    )
    .addSubMenu(
      ui.createMenu('Testing')
        .addItem('Send Test Offer (Syla Doronina)', 'sendTestOfferSylaDoronina')
        .addItem('Resend Syla Slack', 'resendTestOfferSlackSylaDoronina')
        .addItem('Test Slack DM to Me', 'sendTestSlackPingToMe')
    )
    .addSubMenu(
      ui.createMenu('Analytics')
        .addItem('Update TS Analytics', 'refreshTsAnalytics')
        .addItem('Refresh Manager List from Looker', 'refreshManagerAliasRows')
    )
    .addSubMenu(
      ui.createMenu('Admin')
        .addItem('Setup Workbook + Triggers', 'setupTrainingScheduler')
        .addItem('Run Offers Now', 'runTrainingOffersNow')
        .addItem('Run Reminders Now', 'runTrainingRemindersNow')
        .addItem('Dry-run', 'runTrainingOffersDryRun')
    )
    .addToUi();
}

function createSetupSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  createSheetWithHeaders_(spreadsheet, CSV_DUMP_SHEET_NAME, CSV_DUMP_HEADERS);
  BULK_CSV_SHEET_NAMES.forEach(function(sheetName) {
    createSheetWithHeaders_(spreadsheet, sheetName, CSV_DUMP_HEADERS);
  });
  createSheetWithHeaders_(spreadsheet, LOOKER_MANAGER_LOOKUP_SHEET_NAME, LOOKER_MANAGER_LOOKUP_HEADERS);
  createSheetWithHeaders_(spreadsheet, NAME_MATCH_OVERRIDES_SHEET_NAME, NAME_MATCH_OVERRIDES_HEADERS);
  createSheetWithHeaders_(spreadsheet, REMOVED_REPS_SHEET_NAME, REMOVED_REPS_HEADERS);
  createSheetWithHeaders_(spreadsheet, MISSING_LOOKER_PAIRINGS_SHEET_NAME, MISSING_LOOKER_PAIRINGS_HEADERS);
  createSheetWithHeaders_(spreadsheet, EXCEPTION_SHEET_NAME, EXCEPTION_HEADERS);
  createRunSettingsSheet_(spreadsheet);

  var exceptionSheet = spreadsheet.getSheetByName(EXCEPTION_SHEET_NAME);
  ensureExceptionTrackingHeaders_(exceptionSheet);

  var runLogSheet = getOrCreateSheet_(spreadsheet, RUN_LOG_SHEET_NAME);
  if (runLogSheet.getLastRow() === 0) {
    runLogSheet.getRange(1, 1, 1, 4).setValues([[
      'Report Run At',
      'Rows Processed',
      'Exceptions Found',
      'Emails Sent'
    ]]);
    runLogSheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('Setup complete.');
}

function connectLookerManagerImport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(spreadsheet, LOOKER_MANAGER_LOOKUP_SHEET_NAME);
  var formula = '=IMPORTRANGE("' + LOOKER_MANAGER_SOURCE_SPREADSHEET_ID + '","' + LOOKER_MANAGER_SOURCE_RANGE + '")';

  sheet.clear();
  sheet.getRange('A1').setFormula(formula);
  sheet.autoResizeColumns(1, LOOKER_MANAGER_LOOKUP_HEADERS.length);

  SpreadsheetApp.getUi().alert(
    'Looker manager import formula added to "' + LOOKER_MANAGER_LOOKUP_SHEET_NAME + '".\n\n' +
    'If Google Sheets shows #REF!, click the cell and choose Allow access.'
  );
}

function setCurrentJourneyName() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getRunSettings_(spreadsheet);
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Set current journey name',
    'Enter the journey name to use for this CSV when the CSV does not include a Journey column.\n\nCurrent value: ' + settings.currentJourneyName,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var newJourneyName = response.getResponseText().trim();
  if (!newJourneyName) {
    ui.alert('Journey name was not changed because the value was blank.');
    return;
  }

  setRunSetting_(spreadsheet, 'Current Journey Name', newJourneyName);
  ui.alert('Current journey name updated to:\n\n' + newJourneyName);
}

function runWeeklySimulationExceptionReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runSettings = getRunSettings_(spreadsheet);

  var csvRows = loadReflexAiCsvRows_(spreadsheet, runSettings);
  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import ReflexAI CSV data into "' + CSV_DUMP_SHEET_NAME + '" or tabs named "' + CSV_DUMP_SHEET_PREFIX + '[Journey Name]".');
    return;
  }

  var managerRoster = buildLookerManagerRoster_(spreadsheet);
  var exceptions = [];
  var missingLookerPairings = {};

  csvRows.forEach(function(row) {
    var userName = getValue_(row, 'User Name');
    var userEmail = getValue_(row, 'User Email');
    var simulationName = getValue_(row, 'Simulation Name');
    var status = getValue_(row, 'Status');
    var score = getValue_(row, 'Best Score (%)');
    var managerInfo = getManagerInfoForRep_(managerRoster, userEmail, userName);

    if (!managerInfo.managerName) {
      if (String(userName || '').trim()) {
        missingLookerPairings[String(userName).trim()] = true;
      }
      return;
    }
    if (isBlockedManagerForReporting_(managerInfo.managerName, managerInfo.managerEmail)) {
      return;
    }

    var scoreNumber = parseScore_(score);
    var statusLower = String(status).toLowerCase().trim();

    var needsFollowUp = false;
    var action = '';

    if (
      statusLower === 'not started' ||
      statusLower === 'not attempted' ||
      statusLower === 'in progress' ||
      statusLower === 'started' ||
      statusLower === 'pending'
    ) {
      needsFollowUp = true;
      action = COMPLETE_SIMULATION_ACTION;
    } else if (
      statusLower === 'completed' &&
      !isNaN(scoreNumber) &&
      scoreNumber < PASSING_SCORE_PERCENT
    ) {
      needsFollowUp = true;
      action = RETAKE_SIMULATION_ACTION;
    }

    if (!needsFollowUp) {
      return;
    }

    exceptions.push({
      reportRunAt: new Date(),
      repName: userName,
      repEmail: userEmail,
      managerName: managerInfo.managerName || '',
      managerEmail: managerInfo.managerEmail || '',
      seniorName: managerInfo.seniorName || '',
      seniorEmail: managerInfo.seniorEmail || '',
      journeyName: getJourneyNameForRow_(row, runSettings),
      simulationName: simulationName,
      status: status,
      score: isNaN(scoreNumber) ? '' : scoreNumber / 100,
      action: action
    });
  });

  writeMissingLookerPairings_(spreadsheet, Object.keys(missingLookerPairings));
  writeExceptions_(spreadsheet, exceptions);
  writeRunLog_(spreadsheet, csvRows.length, exceptions.length, false);

  SpreadsheetApp.getUi().alert(
    'Report complete.\n\nRows checked: ' +
    csvRows.length +
    '\nExceptions found: ' +
    exceptions.length +
    '\nMissing Looker pairings excluded: ' +
    Object.keys(missingLookerPairings).length
  );
}

function sendTestManagerExceptionEmails() {
  sendManagerEmailBatches_(true);
}

function sendManagerExceptionEmails() {
  sendManagerEmailBatches_(false);
}

function sendManagerEmailBatches_(testMode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runSettings = getRunSettings_(spreadsheet);
  var managerRoster = buildLookerManagerRoster_(spreadsheet);
  var csvRows = loadReflexAiCsvRows_(spreadsheet, runSettings);
  var assignedRows = [];
  var grouped = {};
  var skippedMissingManager = 0;
  var skippedNoManagerEmail = 0;
  var skippedBlockedRecipients = 0;

  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import ReflexAI CSV data into "' + CSV_DUMP_SHEET_NAME + '" or tabs named "' + CSV_DUMP_SHEET_PREFIX + '[Journey Name]".');
    return;
  }

  assignedRows = buildPeakCadenceEmailRows_(spreadsheet, csvRows, managerRoster, runSettings);
  if (!assignedRows.length) {
    SpreadsheetApp.getUi().alert(
      'No peak-cadence assignments found.\n\n' +
      'Manager emails now report TS Offers rows with Status = BOOKED or ESCALATED for the current/prior Weekly Sim Schedule weeks.'
    );
    return;
  }

  assignedRows.forEach(function(rowItem) {
    var managerEmail = String(rowItem.managerEmail || '').toLowerCase().trim();

    if (!rowItem.managerName) {
      skippedMissingManager++;
      return;
    }

    if (!managerEmail) {
      skippedNoManagerEmail++;
      return;
    }

    if (isBlockedManagerForReporting_(rowItem.managerName, managerEmail)) {
      skippedBlockedRecipients++;
      return;
    }

    if (!grouped[managerEmail]) {
      grouped[managerEmail] = {
        managerName: rowItem.managerName || '',
        ccEmails: {},
        metricBucket: newMetricBucket_(rowItem.managerName || ''),
        metricRows: [],
        metrics: null,
        allRows: [],
        rows: []
      };
    }

    var seniorEmail = String(rowItem.seniorEmail || '').toLowerCase().trim();
    if (seniorEmail && seniorEmail !== managerEmail && isAllowedRecipientEmail_(seniorEmail)) {
      grouped[managerEmail].ccEmails[seniorEmail] = true;
    }

    grouped[managerEmail].allRows.push(rowItem);
    grouped[managerEmail].rows.push(rowItem);
    addMetricOutcome_(grouped[managerEmail].metricBucket, rowItem.category, rowItem.scorePercent);
    grouped[managerEmail].metricRows.push(rowItem);
  });

  var managerEmails = Object.keys(grouped).filter(function(managerEmail) {
    return grouped[managerEmail].allRows.length;
  });

  if (!managerEmails.length) {
    SpreadsheetApp.getUi().alert(
      'No manager email batches found.\n\n' +
      'Rows missing Looker manager pairing: ' +
      skippedMissingManager +
      '\nRows missing manager email: ' +
      skippedNoManagerEmail
    );
    return;
  }

  var batchesSent = 0;
  var rowsIncluded = 0;

  managerEmails.forEach(function(managerEmail) {
    var batch = grouped[managerEmail];
    batch.metrics = {
      bucket: finalizeMetricBucket_(batch.metricBucket),
      rows: batch.metricRows
    };
    var intendedCcRecipients = Object.keys(batch.ccEmails).join(',');
    var takeRate = buildOfferTakeRateForManager_(spreadsheet, batch.managerName);
    var inlineImages = {};
    var takeRateChartCid = '';
    if (takeRate.total) {
      takeRateChartCid = 'takeRateChart';
      inlineImages[takeRateChartCid] = buildTakeRatePieChartBlob_(takeRate, 'Manager Take Rate');
    }

    var recipients = testMode
      ? TEST_EMAIL_RECIPIENTS.join(',')
      : managerEmail;
    var ccRecipients = testMode ? '' : intendedCcRecipients;

    if (!recipients) {
      skippedBlockedRecipients += batch.allRows.length;
      return;
    }

    var emailOptions = {
      to: recipients,
      cc: ccRecipients,
      subject: (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ReflexAI Simulation Follow-Up',
      body: buildManagerEmailBody_(batch.managerName, batch.allRows, testMode, managerEmail, intendedCcRecipients),
      htmlBody: buildManagerEmailHtml_(batch.managerName, batch.allRows, testMode, managerEmail, batch.metrics, intendedCcRecipients, takeRate, takeRateChartCid)
    };
    if (takeRateChartCid) {
      emailOptions.inlineImages = inlineImages;
    }
    MailApp.sendEmail(emailOptions);

    batchesSent++;
    rowsIncluded += batch.allRows.length;
  });

  SpreadsheetApp.getUi().alert(
    (testMode ? 'Test emails sent.' : 'Manager emails sent.') +
    '\n\nBatches sent: ' +
    batchesSent +
    '\nRows included: ' +
    rowsIncluded +
    '\nRows missing Looker manager pairing: ' +
    skippedMissingManager +
    '\nRows missing manager email: ' +
    skippedNoManagerEmail +
    '\nRows skipped due to blocked John Paul/Riordan recipient: ' +
    skippedBlockedRecipients
  );
}

function buildManagerMetricBuckets_(csvRows, managerRoster) {
  var buckets = {};

  csvRows.forEach(function(row) {
    var managerInfo = getManagerInfoForRep_(
      managerRoster,
      getValue_(row, 'User Email'),
      getValue_(row, 'User Name')
    );
    var managerEmail = String(managerInfo.managerEmail || '').toLowerCase().trim();

    if (!managerInfo.managerName || !managerEmail || isBlockedManagerForReporting_(managerInfo.managerName, managerEmail)) {
      return;
    }

    if (!buckets[managerEmail]) {
      buckets[managerEmail] = {
        bucket: newMetricBucket_(managerInfo.managerName || ''),
        rows: []
      };
    }

    var score = parseScore_(getValue_(row, 'Best Score (%)'));
    var category = classifySimulationOutcome_(getValue_(row, 'Status'), score);
    addMetricOutcome_(buckets[managerEmail].bucket, category, score);
    buckets[managerEmail].rows.push({
      simulationName: getValue_(row, 'Simulation Name') || 'Unknown Simulation',
      score: isNaN(score) ? '' : score / 100
    });
  });

  Object.keys(buckets).forEach(function(managerEmail) {
    buckets[managerEmail].bucket = finalizeMetricBucket_(buckets[managerEmail].bucket);
  });

  return buckets;
}

function sendTestSeniorLeadershipRecap() {
  sendSeniorLeadershipRecaps_(true);
}

function sendSeniorLeadershipRecap() {
  sendSeniorLeadershipRecaps_(false);
}

function sendSeniorLeadershipRecaps_(testMode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runSettings = getRunSettings_(spreadsheet);
  var csvRows = loadReflexAiCsvRows_(spreadsheet, runSettings);

  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import ReflexAI CSV data into "' + CSV_DUMP_SHEET_NAME + '" or tabs named "' + CSV_DUMP_SHEET_PREFIX + '[Journey Name]".');
    return;
  }

  var managerRoster = buildLookerManagerRoster_(spreadsheet);
  var assignedRows = buildPeakCadenceEmailRows_(spreadsheet, csvRows, managerRoster, runSettings);
  if (!assignedRows.length) {
    SpreadsheetApp.getUi().alert(
      'No peak-cadence assignments found.\n\n' +
      'Senior leader emails now report TS Offers rows with Status = BOOKED or ESCALATED for the current/prior Weekly Sim Schedule weeks.'
    );
    return;
  }
  var sentCount = 0;

  SENIOR_LEADER_SUPERGROUP_RECIPIENTS.forEach(function(config) {
    var filteredRows = filterAssignedRowsForSupergroup_(assignedRows, config.supergroupName);
    var recap = buildAssignedSeniorLeadershipRecap_(filteredRows, config.supergroupName);
    recap.supergroupName = config.supergroupName;
    var recipients = testMode ? TEST_EMAIL_RECIPIENTS : [config.email];
    var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Supergroup Recap';
    var takeRate = buildOfferTakeRateForSupergroup_(spreadsheet, config.supergroupName);
    var inlineImages = {};
    var takeRateChartCid = '';
    if (takeRate.total) {
      takeRateChartCid = 'takeRateChart';
      inlineImages[takeRateChartCid] = buildTakeRatePieChartBlob_(takeRate, config.supergroupName + ' Take Rate');
    }

    var emailOptions = {
      to: recipients.join(','),
      subject: subject,
      body: buildSeniorLeadershipRecapText_(recap, testMode, config),
      htmlBody: buildSeniorLeadershipRecapHtml_(recap, testMode, config, takeRate, takeRateChartCid)
    };
    if (takeRateChartCid) {
      emailOptions.inlineImages = inlineImages;
    }
    MailApp.sendEmail(emailOptions);

    sentCount++;
  });

  SpreadsheetApp.getUi().alert(
    (testMode ? 'Test senior leader recaps sent.' : 'Senior leader recaps sent.') +
    '\n\nEmails sent: ' +
    sentCount +
    '\nRecipients: ' +
    (testMode ? TEST_EMAIL_RECIPIENTS.join(', ') : 'configured senior leaders by supergroup')
  );
}

function filterRowsForSupergroup_(rows, supergroupName, runSettings) {
  var target = normalizePersonKey_(supergroupName);

  return rows.filter(function(row) {
    var rowSupergroup = deriveSupergroupName_(getJourneyNameForRow_(row, runSettings || {}));
    return normalizePersonKey_(rowSupergroup) === target;
  });
}

function filterAssignedRowsForSupergroup_(rows, supergroupName) {
  var target = peakSalesGroupKey_(supergroupName);
  return (rows || []).filter(function(row) {
    return peakSalesGroupKey_(row.supergroupName || row.salesGroup) === target;
  });
}

function buildAssignedSeniorLeadershipRecap_(assignedRows, supergroupName) {
  var recap = {
    generatedAt: new Date(),
    totalRows: assignedRows.length,
    supergroupName: supergroupName || 'Supergroup',
    counts: {
      notStarted: 0,
      completedBelow: 0,
      completedAbove: 0
    },
    bySimulation: {},
    byManager: {},
    overdueRows: []
  };

  assignedRows.forEach(function(row) {
    if (isBlockedManagerForReporting_(row.managerName, row.managerEmail)) return;

    recap.counts[row.category]++;

    var managerKey = getManagerRecapKey_(row.managerName, row.managerEmail);
    if (!managerKey) managerKey = 'unassigned';
    if (!recap.byManager[managerKey]) {
      recap.byManager[managerKey] = {
        managerName: row.managerName || 'Unassigned',
        managerEmail: row.managerEmail || '',
        notStarted: 0,
        completedBelow: 0,
        completedAbove: 0,
        total: 0
      };
    }

    recap.byManager[managerKey][row.category]++;
    recap.byManager[managerKey].total++;

    if (isCompletedStatus_(row.status) && isFinite(row.scorePercent)) {
      if (!recap.bySimulation[row.assignedSimulationName]) {
        recap.bySimulation[row.assignedSimulationName] = {
          simulationName: row.assignedSimulationName,
          completedCount: 0,
          scoreTotal: 0
        };
      }
      recap.bySimulation[row.assignedSimulationName].completedCount++;
      recap.bySimulation[row.assignedSimulationName].scoreTotal += row.scorePercent;
    }

    if (row.overdue) {
      recap.overdueRows.push(row);
    }
  });

  recap.simulationAverages = Object.keys(recap.bySimulation)
    .map(function(key) {
      var item = recap.bySimulation[key];
      return {
        simulationName: item.simulationName,
        completedCount: item.completedCount,
        averageScore: item.completedCount ? item.scoreTotal / item.completedCount : null
      };
    })
    .sort(function(a, b) {
      return String(a.simulationName).localeCompare(String(b.simulationName));
    });

  recap.managerRows = Object.keys(recap.byManager)
    .map(function(key) {
      return recap.byManager[key];
    })
    .sort(function(a, b) {
      return String(a.managerName).localeCompare(String(b.managerName));
    });

  return recap;
}

function getSeniorLeaderRecipients_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(EXCEPTION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(header) {
    return String(header).trim();
  });
  var col = buildColumnIndex_(headers);
  var seniorEmailColumn = col['Senior / Team Lead Email'] !== undefined
    ? col['Senior / Team Lead Email']
    : 6;
  var seen = {};
  var recipients = [];

  values.slice(1).forEach(function(row) {
    var email = String(row[seniorEmailColumn] || '').toLowerCase().trim();
    if (!email || seen[email]) {
      return;
    }

    seen[email] = true;
    recipients.push(email);
  });

  return recipients;
}

function buildPeakCadenceEmailRows_(spreadsheet, csvRows, managerRoster, runSettings) {
  var scheduleIndex = buildWeeklyAssignmentScheduleIndex_(spreadsheet);
  var csvLookup = buildCsvRowsLookup_(csvRows);
  var removalLookup = buildRemovedRepLookup_(spreadsheet);
  var offerRows = readSheetRows_(spreadsheet, 'TS Offers');
  var rows = [];
  var seen = {};
  var now = new Date();
  var currentDayEnd = new Date(now.getTime());
  currentDayEnd.setHours(23, 59, 59, 999);

  offerRows.forEach(function(offerRow) {
    var status = String(getValue_(offerRow, 'Status') || '').trim().toUpperCase();
    var isEscalated = status === 'ESCALATED';
    if (status !== 'BOOKED' && !isEscalated) return;
    if (tsOfferRowInactiveForEmail_(offerRow)) return;

    var repEmail = String(getFirstNonBlankValue_(offerRow, [
      'Consultant Email',
      'Email',
      'Representative Email',
      'Rep Email',
      'User Email'
    ]) || '').toLowerCase().trim();
    var repName = getFirstNonBlankValue_(offerRow, [
      'Consultant',
      'Consultant Name',
      'Representative',
      'Representative Name',
      'Rep',
      'User Name'
    ]) || '';
    if (isRemovedRepForEmail_(removalLookup, repEmail, repName)) return;

    var salesGroup = getValue_(offerRow, 'Sales Group') || '';
    var sims = String(getValue_(offerRow, 'Sims') || getValue_(offerRow, 'Sims CSV') || '')
      .split(',')
      .map(function(value) { return value.trim(); })
      .filter(Boolean);
    var bookedWindow = getValue_(offerRow, 'Booked Window');
    var createdAt = getValue_(offerRow, 'Created At') || getValue_(offerRow, 'Offer Sent At');
    var bookedStart = parseBookedWindowStart_(bookedWindow) || parsePeakDate_(createdAt);

    sims.forEach(function(simName) {
      var assignment = findWeeklyAssignmentForOffer_(scheduleIndex, salesGroup, simName, bookedStart);
      var dueWeekStart = assignment && assignment.weekStart
        ? assignment.weekStart
        : peakWeekStartFromDate_(bookedStart);

      if (dueWeekStart && dueWeekStart.getTime() > currentDayEnd.getTime()) return;

      var dedupeKey = [
        repEmail || normalizePersonKey_(repName),
        peakNormalizeSimulationName_(simName),
        String(bookedWindow || ''),
        dueWeekStart ? dueWeekStart.getTime() : ''
      ].join('|');
      if (seen[dedupeKey]) return;
      seen[dedupeKey] = true;

      var csvRow = findCsvRowForAssignedSim_(getCsvCandidatesForOffer_(csvLookup, repEmail, repName), simName);
      if (!csvRow) return;

      var csvStatus = getValue_(csvRow, 'Status');
      var score = parseScore_(getValue_(csvRow, 'Best Score (%)'));
      var displayStatus = isEscalated && !isCompletedStatus_(csvStatus) ? 'Escalated' : csvStatus;
      var category = classifySimulationOutcome_(displayStatus, score);
      var csvRepName = getValue_(csvRow, 'User Name');
      var csvJourneyName = getJourneyNameForRow_(csvRow, runSettings || {});
      var managerInfo = getManagerInfoForRep_(managerRoster, repEmail, csvRepName || repName);
      if (!managerInfo.managerName) return;

      var managerName = normalizeManagerDisplayName_(managerInfo.managerName);
      var managerEmail = String(managerInfo.managerEmail || emailFromName_(managerName) || '').toLowerCase().trim();

      if (isBlockedManagerForReporting_(managerName, managerEmail)) return;

      rows.push({
        repName: csvRepName || repName || repEmail,
        repEmail: repEmail,
        managerName: managerName,
        managerEmail: managerEmail,
        seniorName: managerInfo.seniorName || '',
        seniorEmail: managerInfo.seniorEmail || '',
        salesGroup: salesGroup,
        supergroupName: deriveSupergroupName_(csvJourneyName || salesGroup),
        journeyName: csvJourneyName || salesGroup,
        simulationName: simName,
        assignedSimulationName: assignment && assignment.simulationName ? assignment.simulationName : simName,
        dueWeekStart: dueWeekStart,
        dueWeekLabel: formatWeekLabel_(dueWeekStart),
        dateAssigned: bookedStart,
        dateAssignedLabel: (isEscalated ? 'Escalated ' : '') + formatAssignedDate_(bookedStart, bookedWindow || createdAt),
        status: displayStatus || 'Not Started',
        score: isNaN(score) ? '' : score / 100,
        scorePercent: score,
        category: category,
        escalated: isEscalated,
        overdue: !!(!isEscalated && bookedStart && bookedStart.getTime() < now.getTime() && !isCompletedStatus_(csvStatus))
      });
    });
  });

  return rows.sort(function(a, b) {
    var managerCompare = String(a.managerName).localeCompare(String(b.managerName));
    if (managerCompare) return managerCompare;
    var dateCompare = (a.dateAssigned ? a.dateAssigned.getTime() : 0) - (b.dateAssigned ? b.dateAssigned.getTime() : 0);
    if (dateCompare) return dateCompare;
    return String(a.repName).localeCompare(String(b.repName));
  });
}

function buildWeeklyAssignmentScheduleIndex_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, 'Weekly Sim Schedule');
  var index = {
    byGroupSim: {},
    rows: []
  };

  rows.forEach(function(row) {
    var weekStart = parsePeakDate_(getValue_(row, 'Week Start'));
    var supergroup = getValue_(row, 'Supergroup');
    var simulationName = getValue_(row, 'Simulation Name');
    if (!weekStart || !supergroup || !simulationName) return;

    var item = {
      weekStart: weekStart,
      supergroup: supergroup,
      simulationName: simulationName,
      groupKey: peakSalesGroupKey_(supergroup),
      simKey: peakNormalizeSimulationName_(simulationName)
    };
    index.rows.push(item);
    index.byGroupSim[item.groupKey + '|' + item.simKey] = item;
  });

  return index;
}

function findWeeklyAssignmentForOffer_(scheduleIndex, salesGroup, simName, bookedStart) {
  var groupKey = peakSalesGroupKey_(salesGroup);
  var simKey = peakNormalizeSimulationName_(simName);
  var candidates = scheduleIndex.rows.filter(function(item) {
    return item.groupKey === groupKey && (item.simKey === simKey || peakSimulationNamesMatch_(item.simulationName, simName));
  });
  if (!candidates.length) return null;

  if (bookedStart) {
    candidates.sort(function(a, b) {
      return Math.abs(a.weekStart.getTime() - bookedStart.getTime()) - Math.abs(b.weekStart.getTime() - bookedStart.getTime());
    });
  }

  return candidates[0];
}

function buildCsvRowsLookup_(csvRows) {
  var lookup = {
    byEmail: {},
    byName: {}
  };
  csvRows.forEach(function(row) {
    var email = String(getValue_(row, 'User Email') || '').toLowerCase().trim();
    var nameKey = normalizePersonKey_(getValue_(row, 'User Name'));

    if (email) {
      if (!lookup.byEmail[email]) lookup.byEmail[email] = [];
      lookup.byEmail[email].push(row);
    }

    if (nameKey) {
      if (!lookup.byName[nameKey]) lookup.byName[nameKey] = [];
      lookup.byName[nameKey].push(row);
    }
  });
  return lookup;
}

function getCsvCandidatesForOffer_(csvLookup, repEmail, repName) {
  var candidates = [];
  var seen = {};

  function addRows(rows) {
    (rows || []).forEach(function(row) {
      var key = [
        String(getValue_(row, 'User Email') || '').toLowerCase().trim(),
        normalizePersonKey_(getValue_(row, 'User Name')),
        peakNormalizeSimulationName_(getValue_(row, 'Simulation Name'))
      ].join('|');
      if (seen[key]) return;
      seen[key] = true;
      candidates.push(row);
    });
  }

  addRows(csvLookup.byEmail[String(repEmail || '').toLowerCase().trim()]);
  addRows(csvLookup.byName[normalizePersonKey_(repName)]);

  return candidates;
}

function findCsvRowForAssignedSim_(candidateRows, simName) {
  var exactKey = peakNormalizeSimulationName_(simName);
  var fallback = null;

  candidateRows.forEach(function(row) {
    var candidateName = getValue_(row, 'Simulation Name');
    var candidateKey = peakNormalizeSimulationName_(candidateName);
    if (candidateKey === exactKey) {
      fallback = row;
      return;
    }
    if (!fallback && peakSimulationNamesMatch_(candidateName, simName)) {
      fallback = row;
    }
  });

  return fallback;
}

function peakSimulationNamesMatch_(left, right) {
  var leftKey = peakNormalizeSimulationName_(left);
  var rightKey = peakNormalizeSimulationName_(right);
  if (!leftKey || !rightKey) return false;
  if (leftKey === rightKey) return true;
  if (Math.min(leftKey.length, rightKey.length) >= 12 && (leftKey.indexOf(rightKey) !== -1 || rightKey.indexOf(leftKey) !== -1)) return true;
  var maxLength = Math.max(leftKey.length, rightKey.length);
  return maxLength >= 20 && editDistance_(leftKey, rightKey) / maxLength <= 0.12;
}

function peakNormalizeSimulationName_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[’‘]/g, "'")
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(reflex|reflexai|ai|simulation|sim)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function peakSalesGroupKey_(value) {
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

function parseBookedWindowStart_(value) {
  var raw = String(value || '').trim();
  var match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4]), Number(match[5]), 0);
  }

  return parsePeakDate_(value);
}

function parsePeakDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var raw = String(value || '').trim();
  if (!raw) return null;

  var iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }

  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function peakWeekStartFromDate_(dateValue) {
  if (!dateValue || isNaN(dateValue.getTime())) return null;
  var d = new Date(dateValue.getTime());
  d.setHours(0, 0, 0, 0);
  var day = d.getDay();
  d.setDate(d.getDate() - day);
  return d;
}

function formatWeekLabel_(dateValue) {
  if (!dateValue || isNaN(dateValue.getTime())) return '';
  return 'Week of ' + Utilities.formatDate(dateValue, getEmailTimeZone_(), 'MMM d, yyyy');
}

function formatAssignedDate_(dateValue, fallback) {
  if (!dateValue || isNaN(dateValue.getTime())) return String(fallback || '').trim();
  return Utilities.formatDate(dateValue, getEmailTimeZone_(), 'MMM d, h:mm a');
}

function getEmailTimeZone_() {
  return (typeof TS !== 'undefined' && TS.TZ) || Session.getScriptTimeZone() || 'America/Chicago';
}

function isCompletedStatus_(status) {
  return String(status || '').toLowerCase().trim() === 'completed';
}

function isBlockedManagerForReporting_(managerName, managerEmail) {
  return isJohnRiordanName_(managerName) ||
    isBlockedJohnRiordanEmail_(managerEmail) ||
    isMarkDefrancoName_(managerName) ||
    isBlockedMarkDefrancoEmail_(managerEmail) ||
    isIgnoredManagerName_(managerName) ||
    isIgnoredManagerEmail_(managerEmail) ||
    isUnassignedManagerName_(managerName);
}

function isIgnoredManagerName_(value) {
  var normalized = normalizePersonKey_(value);
  return [
    'billy vorgias',
    'john paul riordan',
    'johnpaul riordan',
    'john riordan',
    'ashley roos',
    'margaret etzel',
    'ronaldo felix',
    'tamaira kaster'
  ].indexOf(normalized) !== -1;
}

function isIgnoredManagerEmail_(email) {
  return isIgnoredManagerName_(String(email || '').split('@')[0]);
}

function isMarkDefrancoName_(value) {
  var normalized = normalizePersonKey_(value);
  return [
    'mark defranco',
    'mark de franco',
    'mark defranko',
    'mark de franko'
  ].indexOf(normalized) !== -1;
}

function isBlockedMarkDefrancoEmail_(email) {
  var normalized = normalizePersonKey_(String(email || '').split('@')[0]);
  return isMarkDefrancoName_(normalized);
}

function isUnassignedManagerName_(value) {
  var normalized = normalizePersonKey_(value);
  return !normalized || [
    'unassigned',
    'unassigned manager',
    'missing',
    'missing manager',
    'no manager',
    'none',
    'na',
    'n a'
  ].indexOf(normalized) !== -1;
}

function isRemovedRepForEmail_(removalLookup, email, name) {
  if (!removalLookup || !removalLookup.hasEntries) return false;
  var normalizedEmail = String(email || '').toLowerCase().trim();
  var normalizedName = normalizePersonKey_(name);
  return !!((normalizedEmail && removalLookup.emails[normalizedEmail]) || (normalizedName && removalLookup.names[normalizedName]));
}

function buildOfferTakeRateForManager_(spreadsheet, managerName) {
  return buildOfferTakeRate_(spreadsheet, function(row) {
    return normalizePersonKey_(getValue_(row, 'Manager')) === normalizePersonKey_(managerName);
  });
}

function buildOfferTakeRateForSupergroup_(spreadsheet, supergroupName) {
  return buildOfferTakeRate_(spreadsheet, function(row) {
    return normalizePersonKey_(getValue_(row, 'Sales Group')) === normalizePersonKey_(supergroupName);
  });
}

function buildOfferTakeRateOverall_(spreadsheet) {
  return buildOfferTakeRate_(spreadsheet, function() {
    return true;
  });
}

function applyOpenOfferCountsToManagerRows_(managerRows, takeRate) {
  return applyOpenOfferCountsToMetricRows_(managerRows, takeRate, 'managerName');
}

function applyOpenOfferCountsToMetricRows_(rows, takeRate, nameProperty) {
  var countsByManager = {};

  if (takeRate && takeRate.offeredRows) {
    takeRate.offeredRows.forEach(function(row) {
      var managerName = getValue_(row, 'Manager') || 'Unassigned';
      countsByManager[normalizePersonKey_(managerName)] = (countsByManager[normalizePersonKey_(managerName)] || 0) + 1;
    });
  }

  (rows || []).forEach(function(row) {
    var name = row[nameProperty] || row.name || '';
    row.openOfferCount = countsByManager[normalizePersonKey_(name)] || 0;
  });

  return rows;
}

function buildOfferTakeRate_(spreadsheet, predicate) {
  var removalLookup = buildRemovedRepLookup_(spreadsheet);
  var rows = readSheetRows_(spreadsheet, 'TS Offers').filter(function(row) {
    return !tsOfferRowInactiveForEmail_(row) &&
      !isRemovedRepForEmail_(
        removalLookup,
        getFirstNonBlankValue_(row, ['Consultant Email', 'Email', 'Representative Email', 'Rep Email', 'User Email']),
        getFirstNonBlankValue_(row, ['Consultant', 'Consultant Name', 'Representative', 'Representative Name', 'Rep', 'User Name'])
      ) &&
      !isBlockedManagerForReporting_(getValue_(row, 'Manager') || getValue_(row, 'Manager Name'), '') &&
      predicate(row);
  });
  var takeRate = {
    booked: 0,
    escalated: 0,
    offered: 0,
    total: 0,
    offeredRows: []
  };

  rows.forEach(function(row) {
    var status = String(getValue_(row, 'Status') || '').trim().toUpperCase();
    if (status === 'BOOKED') {
      takeRate.booked++;
      takeRate.total++;
    } else if (status === 'ESCALATED') {
      takeRate.escalated++;
      takeRate.total++;
    } else if (status === 'OFFERED' || status === 'PENDING') {
      takeRate.offered++;
      takeRate.total++;
      takeRate.offeredRows.push(row);
    }
  });

  return takeRate;
}

function tsOfferRowInactiveForEmail_(row) {
  var value = String(
    getValue_(row, 'Active or Inactive') ||
    getValue_(row, 'Active/Inactive') ||
    getValue_(row, 'Active') ||
    ''
  ).trim().toLowerCase();
  return value === 'inactive';
}

function buildTakeRatePieChartBlob_(takeRate, title) {
  var data = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Status')
    .addColumn(Charts.ColumnType.NUMBER, 'Count')
    .addRow(['Booked', takeRate.booked])
    .addRow(['Escalated', takeRate.escalated])
    .addRow(['Offered', takeRate.offered])
    .build();

  return Charts.newPieChart()
    .setDataTable(data)
    .setTitle(title)
    .setDimensions(420, 260)
    .setColors([TAKE_RATE_BOOKED_COLOR, TAKE_RATE_ESCALATED_COLOR, TAKE_RATE_OFFERED_COLOR])
    .build()
    .getAs('image/png')
    .setName('take-rate.png');
}

function buildTakeRateSectionHtml_(eyebrow, title, takeRate, chartCid, detailsHtml) {
  if (!takeRate || !takeRate.total) {
    return '';
  }

  var chartHtml = chartCid
    ? '<div style="text-align:center;margin:8px 0 14px 0;"><img src="cid:' + chartCid + '" style="max-width:420px;width:100%;height:auto;border:0;"></div>'
    : '';

  return sectionCard_(
    eyebrow,
    title,
    metricTiles_([
      { label: 'Booked', subLabel: 'Selected time block via Ops Bot', value: formatCountPercent_(takeRate.booked, takeRate.booked / takeRate.total), color: TAKE_RATE_BOOKED_COLOR },
      { label: 'Escalated', subLabel: 'Manager notified; manual scheduling required', value: formatCountPercent_(takeRate.escalated, takeRate.escalated / takeRate.total), color: TAKE_RATE_ESCALATED_COLOR },
      { label: 'Offered', subLabel: 'Yet to action Ops Bot offerings', value: formatCountPercent_(takeRate.offered, takeRate.offered / takeRate.total), color: TAKE_RATE_OFFERED_COLOR }
    ]) +
    chartHtml +
    (detailsHtml || '')
  );
}

function buildManagerOpenOfferDetailsHtml_(takeRate) {
  if (!takeRate || !takeRate.offeredRows.length) {
    return '';
  }

  var rows = takeRate.offeredRows.map(function(row) {
    return '<tr>' +
      '<td>' + escapeHtml_(getValue_(row, 'Consultant')) + '</td>' +
      '<td>' + escapeHtml_(formatOfferSentDate_(getValue_(row, 'Offer Sent At') || getValue_(row, 'Created At'))) + '</td>' +
      '</tr>';
  }).join('');

  return '<div style="font-size:13px;line-height:19px;color:#4d49a3;margin:0 0 10px 0;"><strong>Open offerings still needing action:</strong></div>' +
    styledTable_(['Representative', 'Offer Sent'], rows);
}

function buildSeniorOpenOfferDetailsHtml_(takeRate) {
  if (!takeRate || !takeRate.offeredRows.length) {
    return '';
  }

  var managerMap = {};
  takeRate.offeredRows.forEach(function(row) {
    var managerName = getValue_(row, 'Manager') || 'Unassigned';
    if (!managerMap[managerName]) managerMap[managerName] = [];
    managerMap[managerName].push(getValue_(row, 'Consultant'));
  });

  var rows = Object.keys(managerMap).sort().map(function(managerName) {
    var reps = managerMap[managerName].filter(Boolean).sort();
    return '<tr>' +
      '<td>' + escapeHtml_(managerName) + '</td>' +
      '<td>' + escapeHtml_(reps.join(', ')) + '</td>' +
      '<td>' + reps.length + '</td>' +
      '</tr>';
  }).join('');

  return '<div style="font-size:13px;line-height:19px;color:#4d49a3;margin:0 0 10px 0;"><strong>Open offerings by manager:</strong></div>' +
    styledTable_(['Manager', 'Representatives', 'Count'], rows);
}

function buildDirectorOpenOfferDetailsHtml_(takeRate) {
  if (!takeRate || !takeRate.offeredRows.length) {
    return '';
  }

  var managerMap = {};
  takeRate.offeredRows.forEach(function(row) {
    var managerName = getValue_(row, 'Manager') || 'Unassigned';
    managerMap[managerName] = (managerMap[managerName] || 0) + 1;
  });

  var rows = Object.keys(managerMap).sort().map(function(managerName) {
    return '<tr>' +
      '<td>' + escapeHtml_(managerName) + '</td>' +
      '<td>' + managerMap[managerName] + '</td>' +
      '</tr>';
  }).join('');

  return '<div style="font-size:13px;line-height:19px;color:#4d49a3;margin:0 0 10px 0;"><strong>Open offerings by manager:</strong></div>' +
    styledTable_(['Manager', 'Reps Not Actioned'], rows);
}

function sendTestDirectorEmail() {
  sendDirectorEmail_(TEST_EMAIL_RECIPIENTS, true);
}

function sendDirectorEmail() {
  sendDirectorEmail_(DIRECTOR_EMAIL_RECIPIENTS, false);
}

function sendDirectorEmail_(recipients, testMode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runSettings = getRunSettings_(spreadsheet);
  var csvRows = loadReflexAiCsvRows_(spreadsheet, runSettings);

  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import ReflexAI CSV data into "' + CSV_DUMP_SHEET_NAME + '" or tabs named "' + CSV_DUMP_SHEET_PREFIX + '[Journey Name]".');
    return;
  }

  var managerRoster = buildLookerManagerRoster_(spreadsheet);
  var assignedRows = buildPeakCadenceEmailRows_(spreadsheet, csvRows, managerRoster, runSettings);
  if (!assignedRows.length) {
    SpreadsheetApp.getUi().alert(
      'No peak-cadence assignments found.\n\n' +
      'Director emails now report TS Offers rows with Status = BOOKED or ESCALATED for the current/prior Weekly Sim Schedule weeks.'
    );
    return;
  }
  var recap = buildDirectorRecapFromAssignedRows_(assignedRows, runSettings);
  var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Director Recap';
  var takeRate = buildOfferTakeRateOverall_(spreadsheet);
  var inlineImages = {};
  var takeRateChartCid = '';
  if (takeRate.total) {
    takeRateChartCid = 'takeRateChart';
    inlineImages[takeRateChartCid] = buildTakeRatePieChartBlob_(takeRate, 'Director Take Rate');
  }

  var emailOptions = {
    to: recipients.join(','),
    subject: subject,
    body: buildDirectorEmailText_(recap, testMode),
    htmlBody: buildDirectorEmailHtml_(recap, testMode, takeRate, takeRateChartCid)
  };
  if (takeRateChartCid) {
    emailOptions.inlineImages = inlineImages;
  }
  MailApp.sendEmail(emailOptions);

  SpreadsheetApp.getUi().alert(
    'Director recap sent.\n\nRecipients: ' +
    recipients.join(', ') +
    '\nRows included: ' +
    recap.totalRows
  );
}

function buildDirectorRecapFromAssignedRows_(assignedRows, runSettings) {
  var recap = {
    totalRows: 0,
    supergroupName: getCurrentSupergroupNameFromAssignedRows_(assignedRows, runSettings || {}),
    company: newMetricBucket_(),
    bySupergroup: {},
    byManager: {},
    overdueRows: []
  };

  assignedRows.forEach(function(row) {
    if (isBlockedManagerForReporting_(row.managerName, row.managerEmail)) return;

    var supergroupName = row.supergroupName || deriveSupergroupName_(row.journeyName || row.salesGroup);
    var managerName = normalizeManagerDisplayName_(row.managerName || 'Unassigned');
    var managerEmail = row.managerEmail || '';
    var managerKey = getManagerRecapKey_(managerName, managerEmail);
    if (!managerKey) managerKey = 'unassigned';

    recap.totalRows++;
    addMetricOutcome_(recap.company, row.category, row.scorePercent);

    if (!recap.bySupergroup[supergroupName]) {
      recap.bySupergroup[supergroupName] = newMetricBucket_(supergroupName);
    }
    addMetricOutcome_(recap.bySupergroup[supergroupName], row.category, row.scorePercent);

    if (!recap.byManager[managerKey]) {
      recap.byManager[managerKey] = newMetricBucket_(managerName);
      recap.byManager[managerKey].managerEmail = managerEmail;
    }
    addMetricOutcome_(recap.byManager[managerKey], row.category, row.scorePercent);

    if (row.overdue) {
      recap.overdueRows.push(row);
    }
  });

  recap.supergroupRows = Object.keys(recap.bySupergroup)
    .map(function(key) {
      return finalizeMetricBucket_(recap.bySupergroup[key]);
    })
    .sort(function(a, b) {
      return String(a.name).localeCompare(String(b.name));
    });

  recap.managerRows = Object.keys(recap.byManager)
    .map(function(key) {
      return finalizeMetricBucket_(recap.byManager[key]);
    })
    .sort(metricRankSort_);

  recap.company = finalizeMetricBucket_(recap.company);
  recap.notStartedRows = buildDirectorNotStartedRows_({});
  return recap;
}

function getCurrentSupergroupNameFromAssignedRows_(assignedRows, runSettings) {
  var names = {};
  (assignedRows || []).forEach(function(row) {
    var supergroupName = row.supergroupName || deriveSupergroupName_(row.journeyName || row.salesGroup);
    if (supergroupName) names[supergroupName.toLowerCase()] = supergroupName;
  });

  var distinct = Object.keys(names).map(function(key) {
    return names[key];
  }).sort();

  if (distinct.length > 1) return 'All Supergroups';
  if (distinct.length === 1) return distinct[0];
  return deriveSupergroupName_((runSettings && runSettings.currentJourneyName) || DEFAULT_JOURNEY_NAME);
}

function getDirectorRecipients_(csvRows, managerRoster) {
  var seen = {};
  var recipients = [];

  csvRows.forEach(function(row) {
    var managerInfo = getManagerInfoForRep_(
      managerRoster,
      getValue_(row, 'User Email'),
      getValue_(row, 'User Name')
    );
    var email = String(managerInfo.managerEmail || '').toLowerCase().trim();

    if (!email || seen[email]) {
      return;
    }

    seen[email] = true;
    recipients.push(email);
  });

  return recipients;
}

function buildDirectorRecap_(csvRows, managerRoster, runSettings) {
  var recap = {
    totalRows: 0,
    supergroupName: getCurrentSupergroupName_(csvRows, runSettings || {}),
    company: newMetricBucket_(),
    bySupergroup: {},
    byManager: {},
    notStartedDetails: {}
  };

  csvRows.forEach(function(row) {
    var userName = getValue_(row, 'User Name');
    var userEmail = getValue_(row, 'User Email');
    var managerInfo = getManagerInfoForRep_(managerRoster, userEmail, userName);

    if (!managerInfo.managerName || isBlockedManagerForReporting_(managerInfo.managerName, managerInfo.managerEmail)) {
      return;
    }

    var score = parseScore_(getValue_(row, 'Best Score (%)'));
    var category = classifySimulationOutcome_(getValue_(row, 'Status'), score);
    var supergroupName = deriveSupergroupName_(getJourneyNameForRow_(row, runSettings || {}));
    var managerName = normalizeManagerDisplayName_(managerInfo.managerName);
    var managerEmail = managerInfo.managerEmail || '';
    var managerKey = getManagerRecapKey_(managerName, managerEmail);

    recap.totalRows++;
    addMetricOutcome_(recap.company, category, score);

    if (!recap.bySupergroup[supergroupName]) {
      recap.bySupergroup[supergroupName] = newMetricBucket_(supergroupName);
    }
    addMetricOutcome_(recap.bySupergroup[supergroupName], category, score);

    if (!recap.byManager[managerKey]) {
      recap.byManager[managerKey] = newMetricBucket_(managerName);
      recap.byManager[managerKey].managerEmail = managerEmail;
    }
    addMetricOutcome_(recap.byManager[managerKey], category, score);

    if (category === 'notStarted') {
      addDirectorNotStartedDetail_(recap.notStartedDetails, supergroupName, managerName, userName);
    }
  });

  recap.supergroupRows = Object.keys(recap.bySupergroup)
    .map(function(key) {
      return finalizeMetricBucket_(recap.bySupergroup[key]);
    })
    .sort(function(a, b) {
      return String(a.name).localeCompare(String(b.name));
    });

  recap.managerRows = Object.keys(recap.byManager)
    .map(function(key) {
      return finalizeMetricBucket_(recap.byManager[key]);
    })
    .sort(metricRankSort_);

  recap.company = finalizeMetricBucket_(recap.company);
  recap.notStartedRows = buildDirectorNotStartedRows_(recap.notStartedDetails);
  return recap;
}

function addDirectorNotStartedDetail_(details, supergroupName, managerName, repName) {
  if (!repName) return;

  if (!details[supergroupName]) {
    details[supergroupName] = {};
  }

  if (!details[supergroupName][managerName]) {
    details[supergroupName][managerName] = {};
  }

  details[supergroupName][managerName][repName] = true;
}

function buildDirectorNotStartedRows_(details) {
  var rows = [];

  Object.keys(details).sort().forEach(function(supergroupName) {
    Object.keys(details[supergroupName]).sort().forEach(function(managerName) {
      rows.push({
        supergroupName: supergroupName,
        managerName: managerName,
        reps: Object.keys(details[supergroupName][managerName]).sort()
      });
    });
  });

  return rows;
}

function newMetricBucket_(name) {
  return {
    name: name || '',
    completedAbove: 0,
    completedBelow: 0,
    notStarted: 0,
    total: 0,
    completedAboveScoreTotal: 0,
    completedBelowScoreTotal: 0
  };
}

function addMetricOutcome_(bucket, category, score) {
  bucket[category]++;
  bucket.total++;

  if (category === 'completedAbove' && !isNaN(score)) {
    bucket.completedAboveScoreTotal += score;
  }

  if (category === 'completedBelow' && !isNaN(score)) {
    bucket.completedBelowScoreTotal += score;
  }
}

function finalizeMetricBucket_(bucket) {
  bucket.completedAbovePercent = bucket.total ? bucket.completedAbove / bucket.total : 0;
  bucket.completedBelowPercent = bucket.total ? bucket.completedBelow / bucket.total : 0;
  bucket.notStartedPercent = bucket.total ? bucket.notStarted / bucket.total : 0;
  bucket.completedAboveAverageScore = bucket.completedAbove ? bucket.completedAboveScoreTotal / bucket.completedAbove : null;
  bucket.completedBelowAverageScore = bucket.completedBelow ? bucket.completedBelowScoreTotal / bucket.completedBelow : null;
  bucket.completedAverageScore = (bucket.completedAbove + bucket.completedBelow)
    ? (bucket.completedAboveScoreTotal + bucket.completedBelowScoreTotal) / (bucket.completedAbove + bucket.completedBelow)
    : null;
  return bucket;
}

function metricRankSort_(a, b) {
  if (b.completedAbovePercent !== a.completedAbovePercent) {
    return b.completedAbovePercent - a.completedAbovePercent;
  }

  if (b.completedAbove !== a.completedAbove) {
    return b.completedAbove - a.completedAbove;
  }

  if (a.notStartedPercent !== b.notStartedPercent) {
    return a.notStartedPercent - b.notStartedPercent;
  }

  return String(a.name).localeCompare(String(b.name));
}

function buildDirectorEmailText_(recap, testMode) {
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - Director recap preview.');
    lines.push('');
  }

  lines.push('Hi Directors,');
  lines.push('');
  lines.push('Below is the ' + getCurrentMonthName_() + ' ReflexAI director recap thus far for ' + recap.supergroupName + '.');
  lines.push('');
  appendDirectorTextSection_(lines, 'Supergroup Breakdown', recap.supergroupRows.concat([companySummaryRow_(recap.company)]));
  appendDirectorTextSection_(lines, 'Manager Breakdown', recap.managerRows);
  appendOverdueAssignedTextSection_(lines, recap.overdueRows);

  return lines.join('\n');
}

function appendDirectorTextSection_(lines, title, rows) {
  lines.push(title);
  rows.forEach(function(row) {
    lines.push(
      (row.name || 'Company') +
      ' | ' + COMPLETED_CLEARED_LABEL + ': ' + formatCountPercent_(row.completedAbove, row.completedAbovePercent) +
      ' | ' + COMPLETED_NOT_CLEARED_LABEL + ': ' + formatCountPercent_(row.completedBelow, row.completedBelowPercent) +
      ' | ' + NOT_STARTED_LABEL + ': ' + formatCountPercent_(row.notStarted, row.notStartedPercent) +
      ' | Average Completed Score: ' + formatScore_(row.completedAverageScore)
    );
  });
  lines.push('');
}

function appendDirectorNotStartedTextSection_(lines, rows) {
  if (!rows || !rows.length) {
    return;
  }

  lines.push('Not Started Details');
  lines.push('Reference list for the Manager Breakdown section, grouped by supergroup and manager.');

  rows.forEach(function(row) {
    lines.push(row.supergroupName + ' | ' + row.managerName + ' | ' + row.reps.join(', '));
  });

  lines.push('');
}

function emailShell_(title, subtitle, bodyHtml) {
  return '<div style="margin:0;padding:0;background:linear-gradient(135deg,#fbf2ff 0%,#edf4ff 52%,#f7f5ff 100%);font-family:Arial,Helvetica,sans-serif;color:#24205f;">' +
    '<div style="max-width:960px;margin:0 auto;padding:24px;">' +
      '<div style="background:#ffffff;border-radius:22px;overflow:hidden;border:1px solid #dedaf8;box-shadow:0 14px 34px rgba(36,32,95,0.16);">' +
        '<div style="height:12px;background:linear-gradient(90deg,#ffcc33 0%,#ff6b9c 25%,#ff3ec8 48%,#7957ff 72%,#20d5d2 100%);"></div>' +
        '<div style="background:linear-gradient(135deg,#24205f 0%,#353082 68%,#5c4be8 100%);color:#ffffff;padding:30px 34px;">' +
          '<div style="display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.28);border-radius:999px;padding:6px 12px;font-size:12px;letter-spacing:1.2px;text-transform:uppercase;color:#ffffff;font-weight:800;">Varsity Tutors</div>' +
          '<div style="font-size:28px;line-height:34px;font-weight:800;margin-top:8px;">' + escapeHtml_(title) + '</div>' +
          '<div style="font-size:14px;line-height:20px;color:#e8e6ff;margin-top:8px;">' + escapeHtml_(subtitle || '') + '</div>' +
          '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;border-collapse:collapse;"><tr>' +
            '<td style="width:33%;padding-right:8px;"><div style="height:5px;border-radius:999px;background:#ffcc33;"></div></td>' +
            '<td style="width:34%;padding:0 8px;"><div style="height:5px;border-radius:999px;background:#ff3ec8;"></div></td>' +
            '<td style="width:33%;padding-left:8px;"><div style="height:5px;border-radius:999px;background:#20d5d2;"></div></td>' +
          '</tr></table>' +
        '</div>' +
        '<div style="padding:28px 32px;background:#f7f5ff;">' + bodyHtml + '</div>' +
        '<div style="background:linear-gradient(135deg,#24205f 0%,#353082 68%,#5c4be8 100%);color:#ffffff;padding:18px 32px;text-align:center;font-size:14px;font-weight:700;">Questions? Contact Aaron Bunch for Support.</div>' +
      '</div>' +
    '</div>' +
  '</div>';
}

function testBanner_(testMode, message) {
  if (!testMode) return '';

  return '<div style="background:#fff3cd;border:1px solid #ffd966;color:#5f4500;border-radius:14px;padding:12px 16px;margin:0 0 18px 0;font-size:14px;line-height:20px;">' +
    '<strong>TEST MODE</strong> - ' + escapeHtml_(message || 'Preview email.') +
    '</div>';
}

function introCard_(title, message) {
  return '<div style="background:#ffffff;border:1px solid #dedaf8;border-radius:18px;padding:18px 20px;margin:0 0 18px 0;box-shadow:0 6px 18px rgba(36,32,95,0.07);">' +
    '<div style="font-size:17px;font-weight:800;color:#24205f;margin-bottom:8px;">' + escapeHtml_(title) + '</div>' +
    '<div style="font-size:14px;line-height:21px;color:#4d49a3;">' + escapeHtml_(message) + '</div>' +
    '</div>';
}

function sectionCard_(eyebrow, title, contentHtml) {
  var accent = accentColorForSection_(eyebrow, title);

  return '<div style="background:#ffffff;border:1px solid #dedaf8;border-radius:20px;padding:0;margin:0 0 24px 0;overflow:hidden;box-shadow:0 14px 30px rgba(36,32,95,0.12);">' +
    '<div style="height:5px;background:' + accent + ';"></div>' +
    '<div style="padding:18px 22px;border-bottom:1px solid #3e398d;background:linear-gradient(135deg,#24205f 0%,#353082 78%,#4f46b7 100%);">' +
      '<div style="display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.22);border-radius:999px;padding:5px 10px;font-size:11px;letter-spacing:1.1px;text-transform:uppercase;color:#ffffff;font-weight:900;">' + escapeHtml_(eyebrow || '') + '</div>' +
      '<div style="font-size:21px;line-height:27px;color:#ffffff;font-weight:900;margin-top:8px;">' + escapeHtml_(title) + '</div>' +
    '</div>' +
    '<div style="padding:20px 22px;background:linear-gradient(180deg,#ffffff 0%,#fbfaff 100%);">' + contentHtml + '</div>' +
    '</div>';
}

function metricTiles_(tiles) {
  return '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>' +
    tiles.map(function(tile) {
      var glow = hexToRgba_(tile.color, 0.18);
      var wash = hexToRgba_(tile.color, 0.07);

      return '<td style="width:33.33%;padding:6px;vertical-align:top;">' +
        '<div style="border:1px solid ' + hexToRgba_(tile.color, 0.38) + ';border-radius:18px;padding:20px 14px;background:linear-gradient(180deg,#ffffff 0%,' + wash + ' 100%);text-align:center;box-shadow:0 14px 30px ' + glow + ';min-height:86px;">' +
          '<div style="font-size:36px;font-weight:900;color:' + tile.color + ';line-height:40px;text-align:center;letter-spacing:-0.5px;">' + escapeHtml_(String(tile.value)) + '</div>' +
          '<div style="font-size:15px;line-height:18px;color:#24205f;font-weight:900;margin-top:10px;letter-spacing:-0.15px;">' + escapeHtml_(shortMetricLabel_(tile.label)) + '</div>' +
          (tile.subLabel ? '<div style="font-size:10px;line-height:13px;color:#6a62d2;font-weight:700;margin-top:5px;">' + escapeHtml_(tile.subLabel) + '</div>' : '') +
        '</div>' +
      '</td>';
    }).join('') +
    '</tr></table>';
}

function shortMetricLabel_(label) {
  if (label === COMPLETED_CLEARED_LABEL) return 'Cleared 80%';
  if (label === COMPLETED_NOT_CLEARED_LABEL) return 'Not Cleared';
  if (label === NOT_STARTED_LABEL) return 'Not Started';
  return label;
}

function hexToRgba_(hex, alpha) {
  var value = String(hex || '').replace('#', '');
  if (value.length !== 6) {
    return 'rgba(36,32,95,' + alpha + ')';
  }

  var red = parseInt(value.slice(0, 2), 16);
  var green = parseInt(value.slice(2, 4), 16);
  var blue = parseInt(value.slice(4, 6), 16);

  return 'rgba(' + red + ',' + green + ',' + blue + ',' + alpha + ')';
}

function progressBar_(completedAbove, completedBelow, notStarted) {
  var total = completedAbove + completedBelow + notStarted;
  if (!total) {
    return '';
  }

  var abovePct = Math.round((completedAbove / total) * 100);
  var belowPct = Math.round((completedBelow / total) * 100);
  var notStartedPct = Math.max(0, 100 - abovePct - belowPct);

  return '<div style="margin:16px 6px 4px 6px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>' +
      '<td width="' + abovePct + '%" style="height:14px;background:' + COMPLETED_CLEARED_COLOR + ';border-radius:999px 0 0 999px;font-size:1px;line-height:1px;">&nbsp;</td>' +
      '<td width="' + belowPct + '%" style="height:14px;background:' + COMPLETED_NOT_CLEARED_COLOR + ';font-size:1px;line-height:1px;">&nbsp;</td>' +
      '<td width="' + notStartedPct + '%" style="height:14px;background:' + NOT_STARTED_COLOR + ';border-radius:0 999px 999px 0;font-size:1px;line-height:1px;">&nbsp;</td>' +
    '</tr></table>' +
  '</div>';
}

function managerSimulationAveragesHtml_(rows) {
  var bySimulation = {};

  rows.forEach(function(row) {
    var simulationName = row.simulationName || 'Unknown Simulation';
    if (!bySimulation[simulationName]) {
      bySimulation[simulationName] = {
        total: 0,
        count: 0,
        assignedCount: 0
      };
    }

    bySimulation[simulationName].assignedCount++;

    var scorePercent = parseEmailScorePercent_(row.score);
    if (isFinite(scorePercent)) {
      bySimulation[simulationName].total += scorePercent;
      bySimulation[simulationName].count++;
    }
  });

  var simulationNames = Object.keys(bySimulation).sort(function(a, b) {
    if (bySimulation[b].assignedCount !== bySimulation[a].assignedCount) {
      return bySimulation[b].assignedCount - bySimulation[a].assignedCount;
    }

    return String(a).localeCompare(String(b));
  }).slice(0, 3);
  if (!simulationNames.length) {
    return '';
  }

  return '<div style="margin:16px 6px 2px 6px;">' +
    '<div style="font-size:12px;letter-spacing:0.7px;text-transform:uppercase;color:#6a62d2;font-weight:900;margin-bottom:8px;">Average Score by Simulation</div>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>' +
      simulationNames.map(function(simulationName) {
        var item = bySimulation[simulationName];
        var average = item.count ? item.total / item.count : null;
        return '<td style="padding:5px;vertical-align:top;">' +
          '<div style="background:#fbfaff;border:1px solid #ebe8ff;border-radius:14px;padding:10px;text-align:center;">' +
            '<div style="font-size:12px;color:#4d49a3;font-weight:800;">' + escapeHtml_(simulationName) + '</div>' +
            '<div style="font-size:18px;color:#24205f;font-weight:900;margin-top:4px;">' + (average === null ? 'N/A' : escapeHtml_(formatScore_(average))) + '</div>' +
          '</div>' +
        '</td>';
      }).join('') +
    '</tr></table>' +
  '</div>';
}

function styledTable_(headers, bodyRowsHtml) {
  var styledBody = bodyRowsHtml
    .replace(/<tr>/g, '<tr style="background:#ffffff;">')
    .replace(/<td([^>]*)>/g, '<td$1 style="padding:11px 10px;border-bottom:1px solid #ebe8ff;vertical-align:middle;">');

  return '<div style="border:1px solid #ebe8ff;border-radius:14px;overflow:hidden;box-shadow:inset 0 1px 0 rgba(255,255,255,0.8);">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;color:#24205f;background:#ffffff;">' +
    '<thead><tr>' +
      headers.map(function(header) {
        return '<th style="text-align:left;background:linear-gradient(180deg,#f3f0ff 0%,#ebe8ff 100%);color:#24205f;padding:12px 10px;border-bottom:1px solid #dedaf8;font-weight:900;">' + escapeHtml_(header) + '</th>';
      }).join('') +
    '</tr></thead>' +
    '<tbody>' + styledBody + '</tbody>' +
    '</table>' +
    '</div>';
}

function accentColorForSection_(eyebrow, title) {
  var text = String(eyebrow || '') + ' ' + String(title || '');
  text = text.toLowerCase();

  if (text.indexOf('not cleared') !== -1 || text.indexOf('priority') !== -1) return COMPLETED_NOT_CLEARED_COLOR;
  if (text.indexOf('action') !== -1 || text.indexOf('not started') !== -1) return NOT_STARTED_COLOR;
  if (text.indexOf('performance') !== -1 || text.indexOf('average') !== -1 || text.indexOf('detail') !== -1) return '#20d5d2';
  if (text.indexOf('snapshot') !== -1 || text.indexOf('cleared') !== -1) return COMPLETED_CLEARED_COLOR;
  if (text.indexOf('manager') !== -1) return '#7957ff';
  return '#6a62d2';
}

function buildDirectorEmailHtml_(recap, testMode, takeRate, takeRateChartCid) {
  var bodyHtml = testBanner_(testMode, 'Director recap preview.') +
    introCard_(
      'Hi Directors,',
      'Below is the ' + getCurrentMonthName_() + ' ReflexAI director recap thus far for ' + recap.supergroupName + '.'
    ) +
    buildTakeRateSectionHtml_(
      'OFFER TAKE RATE',
      'Current Offer Status',
      takeRate,
      takeRateChartCid,
      ''
    ) +
    sectionCard_(
      'CONSUMER SALES SNAPSHOT',
      'Overall Simulation Counts',
      metricTiles_([
        { label: COMPLETED_CLEARED_LABEL, value: formatCountPercent_(recap.company.completedAbove, recap.company.completedAbovePercent), color: COMPLETED_CLEARED_COLOR },
        { label: COMPLETED_NOT_CLEARED_LABEL, value: formatCountPercent_(recap.company.completedBelow, recap.company.completedBelowPercent), color: COMPLETED_NOT_CLEARED_COLOR },
        { label: NOT_STARTED_LABEL, value: formatCountPercent_(recap.company.notStarted, recap.company.notStartedPercent), color: NOT_STARTED_COLOR }
      ])
    ) +
    buildDirectorTable_('Supergroup Breakdown', recap.supergroupRows.concat([companySummaryRow_(recap.company)]), 'Supergroup') +
    buildDirectorTable_('Manager Breakdown', recap.managerRows, 'Manager') +
    buildOverdueAssignedManagerTable_(recap.overdueRows, 'Assigned Time Passed - Not Completed');

  return emailShell_(
    getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Director Recap',
    EMAIL_SUBTITLE,
    bodyHtml
  );
}

function buildDirectorTable_(title, rows, firstColumnLabel) {
  var isManagerBreakdown = title === 'Manager Breakdown';
  var tableRows = rows.map(function(row) {
    var rowStyle = row.isCompanySummary ? ' style="background-color:#e7f7ec;font-weight:bold;"' : '';

    return '<tr' + rowStyle + '>' +
      '<td>' + escapeHtml_(row.name || 'Company') + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedAbove, row.completedAbovePercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedBelow, row.completedBelowPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.notStarted, row.notStartedPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatScore_(row.completedAverageScore)) + '</td>' +
      (isManagerBreakdown ? '' : '<td>' + row.total + '</td>') +
      '</tr>';
  }).join('');

  return sectionCard_(
    title === 'Manager Breakdown' ? 'STACK-RANKED BY CLEARED THRESHOLD' : 'DASHBOARD VIEW',
    title,
    styledTable_([
      firstColumnLabel || 'Segment',
      COMPLETED_CLEARED_LABEL,
      COMPLETED_NOT_CLEARED_LABEL,
      NOT_STARTED_LABEL,
      'Average Completed Score',
      ...(isManagerBreakdown ? [] : ['Total'])
    ], tableRows)
  );
}

function buildDirectorNotStartedTable_(rows) {
  if (!rows || !rows.length) {
    return '';
  }

  var rowsBySupergroup = {};
  rows.forEach(function(row) {
    if (!rowsBySupergroup[row.supergroupName]) {
      rowsBySupergroup[row.supergroupName] = [];
    }
    rowsBySupergroup[row.supergroupName].push(row);
  });

  var tableRows = Object.keys(rowsBySupergroup).sort().map(function(supergroupName) {
    var managerRows = rowsBySupergroup[supergroupName];

    return managerRows.map(function(row, index) {
      var supergroupCell = index === 0
        ? '<td rowspan="' + managerRows.length + '">' + escapeHtml_(supergroupName) + '</td>'
        : '';

      return '<tr>' +
        supergroupCell +
        '<td>' + escapeHtml_(row.managerName) + '</td>' +
        '<td>' + escapeHtml_(row.reps.join(', ')) + '</td>' +
        '<td>' + row.reps.length + '</td>' +
        '</tr>';
    }).join('');
  }).join('');

  return sectionCard_(
    'REFERENCE LIST',
    'Not Started Details',
    '<div style="font-size:13px;line-height:19px;color:#4d49a3;margin-bottom:12px;">Available as supporting detail for the Manager Breakdown section. Grouped by supergroup and manager to keep the recap compact.</div>' +
    styledTable_(['Supergroup', 'Manager', 'Representatives', 'Count'], tableRows)
  );
}

function companySummaryRow_(company) {
  var row = {};
  Object.keys(company).forEach(function(key) {
    row[key] = company[key];
  });
  row.name = 'Consumer Sales';
  row.isCompanySummary = true;
  return row;
}

function buildSeniorLeadershipRecap_(csvRows, managerRoster, runSettings) {
  var recap = {
    generatedAt: new Date(),
    totalRows: csvRows.length,
    supergroupName: getCurrentSupergroupName_(csvRows, runSettings || {}),
    counts: {
      notStarted: 0,
      completedBelow: 0,
      completedAbove: 0
    },
    bySimulation: {},
    byManager: {}
  };

  csvRows.forEach(function(row) {
    var userName = getValue_(row, 'User Name');
    var userEmail = getValue_(row, 'User Email');
    var simulationName = getValue_(row, 'Simulation Name') || 'Unknown Simulation';
    var status = getValue_(row, 'Status');
    var score = parseScore_(getValue_(row, 'Best Score (%)'));
    var category = classifySimulationOutcome_(status, score);
    var managerInfo = getManagerInfoForRep_(managerRoster, userEmail, userName);

    if (!managerInfo.managerName || isBlockedManagerForReporting_(managerInfo.managerName, managerInfo.managerEmail)) {
      return;
    }

    var managerName = normalizeManagerDisplayName_(managerInfo.managerName || 'Unassigned');
    var managerEmail = managerInfo.managerEmail || '';
    var managerKey = getManagerRecapKey_(managerName, managerEmail);

    if (!managerKey) managerKey = 'unassigned';

    recap.counts[category]++;

    if (!recap.byManager[managerKey]) {
      recap.byManager[managerKey] = {
        managerName: managerName,
        managerEmail: managerEmail,
        notStarted: 0,
        completedBelow: 0,
        completedAbove: 0,
        total: 0
      };
    }

    if (!recap.byManager[managerKey].managerEmail && managerEmail) {
      recap.byManager[managerKey].managerEmail = managerEmail;
    }

    recap.byManager[managerKey][category]++;
    recap.byManager[managerKey].total++;

    if (String(status).toLowerCase().trim() === 'completed' && !isNaN(score)) {
      if (!recap.bySimulation[simulationName]) {
        recap.bySimulation[simulationName] = {
          simulationName: simulationName,
          completedCount: 0,
          scoreTotal: 0
        };
      }

      recap.bySimulation[simulationName].completedCount++;
      recap.bySimulation[simulationName].scoreTotal += score;
    }
  });

  recap.simulationAverages = Object.keys(recap.bySimulation)
    .map(function(key) {
      var item = recap.bySimulation[key];
      return {
        simulationName: item.simulationName,
        completedCount: item.completedCount,
        averageScore: item.completedCount ? item.scoreTotal / item.completedCount : null
      };
    })
    .sort(function(a, b) {
      return String(a.simulationName).localeCompare(String(b.simulationName));
    });

  recap.managerRows = Object.keys(recap.byManager)
    .map(function(key) {
      return recap.byManager[key];
    })
    .sort(function(a, b) {
      return String(a.managerName).localeCompare(String(b.managerName));
    });

  return recap;
}

function classifySimulationOutcome_(status, score) {
  var statusLower = String(status).toLowerCase().trim();

  if (statusLower === 'completed' && !isNaN(score)) {
    return score >= PASSING_SCORE_PERCENT ? 'completedAbove' : 'completedBelow';
  }

  return 'notStarted';
}

function getCurrentSupergroupName_(csvRows, runSettings) {
  var supergroupNames = getDistinctSupergroupNames_(csvRows);
  if (supergroupNames.length > 1) {
    return 'All Supergroups';
  }
  if (supergroupNames.length === 1) {
    return supergroupNames[0];
  }

  return deriveSupergroupName_((runSettings && runSettings.currentJourneyName) || DEFAULT_JOURNEY_NAME);
}

function getDistinctSupergroupNames_(csvRows) {
  var names = {};

  csvRows.forEach(function(row) {
    var journeyName = String(row.__journeyName || '').trim() || getFirstNonBlankValue_(row, [
      'Journey',
      'Journey Name',
      'Journey Title',
      'Workshop',
      'Workshop Name',
      'Course',
      'Course Name'
    ]);
    if (!journeyName) return;

    var supergroupName = deriveSupergroupName_(journeyName);
    names[supergroupName.toLowerCase()] = supergroupName;
  });

  return Object.keys(names).map(function(key) {
    return names[key];
  }).sort();
}

function deriveSupergroupName_(journeyName) {
  var value = String(journeyName || '').trim();
  if (!value) return 'Supergroup';

  value = value.replace(/\s*[-–]\s*(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\s*\d*$/i, '');
  value = value.replace(/\s+Year Round Workshops.*$/i, '');
  value = value.replace(/\s+Workshops.*$/i, '');
  value = value.trim();

  return value || String(journeyName || '').trim() || 'Supergroup';
}

function normalizeManagerDisplayName_(managerName) {
  var value = String(managerName || '').trim();
  var normalized = normalizePersonKey_(value);

  if (isJohnRiordanName_(normalized)) {
    return 'John Riordan';
  }

  return value || 'Unassigned';
}

function normalizePersonKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function personNameTokens_(value) {
  return normalizePersonKey_(value)
    .split(' ')
    .filter(Boolean);
}

function firstLastKey_(value) {
  var tokens = personNameTokens_(value);
  if (tokens.length < 2) return '';
  return tokens[0] + '|' + tokens[tokens.length - 1];
}

function firstInitialLastKey_(value) {
  var tokens = personNameTokens_(value);
  if (tokens.length < 2) return '';
  return tokens[0].charAt(0) + '|' + tokens[tokens.length - 1];
}

function allFirstSurnameKeys_(value) {
  var tokens = personNameTokens_(value);
  if (tokens.length < 2) return [];

  var keys = {};
  tokens.slice(1).forEach(function(token) {
    keys[tokens[0] + '|' + token] = true;
  });

  return Object.keys(keys);
}

function allFirstInitialSurnameKeys_(value) {
  var tokens = personNameTokens_(value);
  if (tokens.length < 2) return [];

  var keys = {};
  tokens.slice(1).forEach(function(token) {
    keys[tokens[0].charAt(0) + '|' + token] = true;
  });

  return Object.keys(keys);
}

function findUniqueFuzzyNameMatch_(managerRoster, repName) {
  var requestedTokens = personNameTokens_(repName);
  var candidates = managerRoster.__candidates || [];

  if (requestedTokens.length < 2) {
    return null;
  }

  var requestedFirst = requestedTokens[0];
  var requestedLast = requestedTokens[requestedTokens.length - 1];
  var matches = [];

  candidates.forEach(function(candidate) {
    var candidateTokens = candidate.tokens || [];
    if (candidateTokens.length < 2 || candidateTokens[0] !== requestedFirst) {
      return;
    }

    var bestDistance = candidateTokens.slice(1).reduce(function(best, token) {
      return Math.min(best, editDistance_(requestedLast, token));
    }, 999);

    if (bestDistance <= 2) {
      matches.push({
        distance: bestDistance,
        managerInfo: candidate.managerInfo
      });
    }
  });

  if (!matches.length) {
    return null;
  }

  var bestDistance = matches.reduce(function(best, match) {
    return Math.min(best, match.distance);
  }, 999);
  var bestMatches = matches.filter(function(match) {
    return match.distance === bestDistance;
  });
  var uniqueCandidates = uniqueManagerCandidates_(bestMatches.map(function(match) {
    return match.managerInfo;
  }));

  return uniqueCandidates.length === 1 ? uniqueCandidates[0] : null;
}

function editDistance_(left, right) {
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

function isJohnRiordanName_(value) {
  var normalized = normalizePersonKey_(value);
  return [
    'john riordan',
    'johnpaul riordan',
    'john paul riordan',
    'john ridorian',
    'johnpaul ridorian',
    'john paul ridorian'
  ].indexOf(normalized) !== -1;
}

function isAllowedRecipientEmail_(email) {
  return !isBlockedJohnRiordanEmail_(email);
}

function isBlockedJohnRiordanEmail_(email) {
  var normalized = normalizePersonKey_(String(email || '').split('@')[0]);
  return isJohnRiordanName_(normalized);
}

function getManagerRecapKey_(managerName, managerEmail) {
  var normalizedName = normalizeManagerDisplayName_(managerName);
  if (normalizedName === 'John Riordan') {
    return 'manager:john-riordan';
  }

  return String(managerEmail || normalizedName || 'Unassigned').toLowerCase().trim();
}

function buildSeniorLeadershipRecapText_(recap, testMode, recipientConfig) {
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - Senior leadership recap preview.');
    if (recipientConfig && recipientConfig.email) {
      lines.push('This email would have gone to: ' + recipientConfig.email);
    }
    lines.push('');
  }

  lines.push('Hi ' + ((recipientConfig && recipientConfig.leaderName) || 'Senior Leaders') + ',');
  lines.push('');
  lines.push('Below is the ' + getCurrentMonthName_() + ' ReflexAI recap thus far for ' + recap.supergroupName + '.');
  lines.push('');
  lines.push('Overall simulation counts');
  lines.push(NOT_STARTED_LABEL + ': ' + recap.counts.notStarted);
  lines.push(COMPLETED_NOT_CLEARED_LABEL + ': ' + recap.counts.completedBelow);
  lines.push(COMPLETED_CLEARED_LABEL + ': ' + recap.counts.completedAbove);
  lines.push('');
  lines.push('Average score on completed simulations by simulation');

  recap.simulationAverages.forEach(function(item) {
    lines.push(item.simulationName + ': ' + formatScore_(item.averageScore) + ' across ' + item.completedCount + ' completed simulations');
  });

  lines.push('');
  lines.push('Manager breakdown');
  recap.managerRows.forEach(function(manager) {
    lines.push(
      manager.managerName +
      ' | ' + COMPLETED_CLEARED_LABEL + ': ' +
      formatPercent_(manager.completedAbove, manager.total) +
      ' | ' + COMPLETED_NOT_CLEARED_LABEL + ': ' +
      formatPercent_(manager.completedBelow, manager.total) +
      ' | ' + NOT_STARTED_LABEL + ': ' +
      formatPercent_(manager.notStarted, manager.total)
    );
  });

  appendOverdueAssignedTextSection_(lines, recap.overdueRows);

  return lines.join('\n');
}

function buildSeniorLeadershipRecapHtml_(recap, testMode, recipientConfig, takeRate, takeRateChartCid) {
  var intended = recipientConfig && recipientConfig.email ? 'This email would have gone to: ' + recipientConfig.email + '.' : '';
  var bodyHtml = testBanner_(testMode, 'Senior leadership recap preview. ' + intended) +
    introCard_(
      'Hi ' + ((recipientConfig && recipientConfig.leaderName) || 'Senior Leaders') + ',',
      'Below is the ' + getCurrentMonthName_() + ' ReflexAI recap thus far for ' + recap.supergroupName + '.'
    ) +
    buildTakeRateSectionHtml_(
      'OFFER TAKE RATE',
      'Current Offer Status',
      takeRate,
      takeRateChartCid,
      ''
    ) +
    buildLeadershipCountsTable_(recap) +
    buildOverdueAssignedManagerTable_(recap.overdueRows, 'Assigned Time Passed - Not Completed') +
    buildSimulationAverageTable_(recap.simulationAverages) +
    buildManagerRecapTable_(recap.managerRows);

  return emailShell_(
    getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Supergroup Recap',
    EMAIL_SUBTITLE,
    bodyHtml
  );
}

function appendOverdueAssignedTextSection_(lines, rows) {
  var managerRows = buildOverdueAssignedManagerRows_(rows);
  if (!managerRows.length) return;

  lines.push('');
  lines.push('Assigned time passed - not completed');
  managerRows.forEach(function(row) {
    lines.push(row.managerName + ': ' + row.reps.join(', '));
  });
}

function buildLeadershipCountsTable_(recap) {
  return sectionCard_(
    'SUPERGROUP SNAPSHOT',
    'Overall Simulation Counts',
    metricTiles_([
      { label: COMPLETED_CLEARED_LABEL, value: recap.counts.completedAbove, color: COMPLETED_CLEARED_COLOR },
      { label: COMPLETED_NOT_CLEARED_LABEL, value: recap.counts.completedBelow, color: COMPLETED_NOT_CLEARED_COLOR },
      { label: NOT_STARTED_LABEL, value: recap.counts.notStarted, color: NOT_STARTED_COLOR }
    ]) +
    progressBar_(recap.counts.completedAbove, recap.counts.completedBelow, recap.counts.notStarted)
  );
}

function buildSimulationAverageTable_(simulationAverages) {
  var rows = simulationAverages.map(function(item) {
    return '<tr>' +
      '<td>' + escapeHtml_(item.simulationName) + '</td>' +
      '<td>' + item.completedCount + '</td>' +
      '<td>' + escapeHtml_(formatScore_(item.averageScore)) + '</td>' +
      '</tr>';
  }).join('');

  return sectionCard_(
    'PERFORMANCE DETAIL',
    'Average Score on Completed Simulations by Simulation',
    styledTable_(['Simulation', 'Completed Count', 'Average Score'], rows)
  );
}

function buildManagerRecapTable_(managerRows) {
  var rows = managerRows.map(function(manager) {
    return '<tr>' +
      '<td>' + escapeHtml_(manager.managerName) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.completedAbove, manager.total)) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.completedBelow, manager.total)) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.notStarted, manager.total)) + '</td>' +
      '<td>' + manager.total + '</td>' +
      '</tr>';
  }).join('');

  return sectionCard_(
    'MANAGER VIEW',
    'Manager Breakdown',
    styledTable_([
      'Manager',
      '% ' + COMPLETED_CLEARED_LABEL,
      '% ' + COMPLETED_NOT_CLEARED_LABEL,
      '% ' + NOT_STARTED_LABEL,
      'Total Simulations'
    ], rows)
  );
}

function buildOverdueAssignedManagerTable_(rows, title) {
  var managerRows = buildOverdueAssignedManagerRows_(rows);
  if (!managerRows.length) return '';

  var tableRows = managerRows.map(function(row) {
    return '<tr>' +
      '<td>' + escapeHtml_(row.managerName) + '</td>' +
      '<td>' + escapeHtml_(row.reps.join(', ')) + '</td>' +
      '</tr>';
  }).join('');

  return sectionCard_(
    'QUICK GLANCE',
    title || 'Assigned Time Passed - Not Completed',
    styledTable_(['Manager', 'Reps'], tableRows)
  );
}

function buildOverdueAssignedManagerRows_(rows) {
  var byManager = {};
  (rows || []).forEach(function(row) {
    if (!row.overdue) return;
    if (isBlockedManagerForReporting_(row.managerName, row.managerEmail)) return;
    var managerName = row.managerName || 'Unassigned';
    if (!byManager[managerName]) byManager[managerName] = {};
    if (row.repName) byManager[managerName][row.repName] = true;
  });

  return Object.keys(byManager).sort().map(function(managerName) {
    return {
      managerName: managerName,
      reps: Object.keys(byManager[managerName]).sort()
    };
  });
}

function buildLookerManagerRoster_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, LOOKER_MANAGER_LOOKUP_SHEET_NAME);
  var roster = {};
  roster.__candidates = [];
  var exactByName = {};
  var firstLastCandidates = {};
  var firstInitialLastCandidates = {};
  var seniorRoster = buildManagerRosterSeniorLookup_(spreadsheet);

  rows.forEach(function(row) {
    var repName = getFirstNonBlankValue_(row, [
      'Rep Name',
      'Representative Name',
      'User Name',
      'Name',
      'Manager'
    ]);
    var repEmail = getFirstNonBlankValue_(row, [
      'Rep Email',
      'Representative Email',
      'User Email',
      'Email',
      'Employee Email',
      'Manager Email'
    ]);
    var managerName = getFirstNonBlankValue_(row, [
      'Rep Manager',
      'Representative Manager',
      'Manager Name',
      'Regional Director',
      'Regional Directo',
      'Regional Dir',
      'Senior Leader'
    ]);

    if (!repName || !managerName) {
      return;
    }

    var managerEmail = getFirstNonBlankValue_(row, [
      'Regional Director Email',
      'Regional Directo Email',
      'Regional Dir Email',
      'Senior Leader Email'
    ]) || emailFromName_(managerName);

    var managerInfo = {
      managerName: managerName,
      managerEmail: managerEmail,
      seniorName: '',
      seniorEmail: ''
    };
    var seniorInfo = getSeniorInfoForRep_(seniorRoster, repEmail, repName);

    if (seniorInfo.seniorName) {
      managerInfo.seniorName = seniorInfo.seniorName;
      managerInfo.seniorEmail = seniorInfo.seniorEmail;
    }

    var normalizedRepName = normalizePersonKey_(repName);
    if (normalizedRepName) {
      roster[normalizedRepName] = managerInfo;
      exactByName[normalizedRepName] = managerInfo;
      allFirstSurnameKeys_(repName).forEach(function(key) {
        addCandidate_(firstLastCandidates, key, managerInfo);
      });
      allFirstInitialSurnameKeys_(repName).forEach(function(key) {
        addCandidate_(firstInitialLastCandidates, key, managerInfo);
      });
      roster.__candidates.push({
        repName: repName,
        tokens: personNameTokens_(repName),
        managerInfo: managerInfo
      });
    }

    repEmail = String(repEmail || '').toLowerCase().trim();

    if (repEmail && repEmail.indexOf('@') !== -1) {
      roster[repEmail] = managerInfo;
    }
  });

  addUniqueCandidateMatches_(roster, 'firstlast:', firstLastCandidates);
  addUniqueCandidateMatches_(roster, 'initiallast:', firstInitialLastCandidates);
  applyNameMatchOverrides_(spreadsheet, roster, exactByName);

  return roster;
}

function buildManagerRosterSeniorLookup_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, MANAGER_ROSTER_SHEET_NAME);
  var lookup = {};
  var nameToEmail = {};

  rows.forEach(function(row) {
    var name = getValue_(row, 'Name');
    var email = String(getValue_(row, 'Email') || '').toLowerCase().trim();
    var normalizedName = normalizePersonKey_(name);

    if (normalizedName && email) {
      nameToEmail[normalizedName] = email;
    }
  });

  rows.forEach(function(row) {
    var repEmail = String(getValue_(row, 'Email') || '').toLowerCase().trim();
    var repName = getValue_(row, 'Name');
    var seniorName = getValue_(row, 'Senior');

    if (!seniorName) {
      return;
    }

    var seniorEmail = nameToEmail[normalizePersonKey_(seniorName)] || emailFromName_(seniorName);
    var seniorInfo = {
      seniorName: seniorName,
      seniorEmail: seniorEmail
    };

    if (repEmail) {
      lookup[repEmail] = seniorInfo;
    }

    if (repName) {
      lookup[normalizePersonKey_(repName)] = seniorInfo;
    }
  });

  return lookup;
}

function getSeniorInfoForRep_(seniorRoster, repEmail, repName) {
  return seniorRoster[String(repEmail || '').toLowerCase().trim()] ||
    seniorRoster[normalizePersonKey_(repName)] ||
    {};
}

function getManagerInfoForRep_(managerRoster, repEmail, repName) {
  var emailKey = String(repEmail || '').toLowerCase().trim();
  var exactNameKey = normalizePersonKey_(repName);
  var firstLast = firstLastKey_(repName);
  var initialLast = firstInitialLastKey_(repName);

  return managerRoster[emailKey] ||
    managerRoster[exactNameKey] ||
    managerRoster['firstlast:' + firstLast] ||
    managerRoster['initiallast:' + initialLast] ||
    findUniqueFuzzyNameMatch_(managerRoster, repName) ||
    {};
}

function addCandidate_(candidateMap, key, managerInfo) {
  if (!key) return;
  if (!candidateMap[key]) candidateMap[key] = [];
  candidateMap[key].push(managerInfo);
}

function addUniqueCandidateMatches_(roster, prefix, candidateMap) {
  Object.keys(candidateMap).forEach(function(key) {
    var uniqueCandidates = uniqueManagerCandidates_(candidateMap[key]);
    if (uniqueCandidates.length === 1) {
      roster[prefix + key] = uniqueCandidates[0];
    }
  });
}

function uniqueManagerCandidates_(candidates) {
  var seen = {};
  var unique = [];

  candidates.forEach(function(candidate) {
    var key = [
      candidate.managerName,
      candidate.managerEmail,
      candidate.seniorName,
      candidate.seniorEmail
    ].map(function(value) {
      return String(value || '').toLowerCase().trim();
    }).join('|');

    if (seen[key]) return;
    seen[key] = true;
    unique.push(candidate);
  });

  return unique;
}

function applyNameMatchOverrides_(spreadsheet, roster, exactByName) {
  var rows = readSheetRows_(spreadsheet, NAME_MATCH_OVERRIDES_SHEET_NAME);

  rows.forEach(function(row) {
    var reflexAiName = getValue_(row, 'ReflexAI Name');
    var lookerName = getValue_(row, 'Looker Name');
    var managerInfo = exactByName[normalizePersonKey_(lookerName)];

    if (!reflexAiName || !managerInfo) {
      return;
    }

    roster[normalizePersonKey_(reflexAiName)] = managerInfo;
  });
}

function createRunSettingsSheet_(spreadsheet) {
  var sheet = getOrCreateSheet_(spreadsheet, RUN_SETTINGS_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    sheet.getRange(2, 1, 1, 2).setValues([['Current Journey Name', DEFAULT_JOURNEY_NAME]]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
    return;
  }

  var settings = getRunSettings_(spreadsheet);
  if (!settings.currentJourneyName) {
    setRunSetting_(spreadsheet, 'Current Journey Name', DEFAULT_JOURNEY_NAME);
  }
}

function getRunSettings_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(RUN_SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      currentJourneyName: DEFAULT_JOURNEY_NAME
    };
  }

  var values = sheet.getDataRange().getValues();
  var settings = {};

  values.slice(1).forEach(function(row) {
    var key = String(row[0] || '').trim();
    var value = String(row[1] || '').trim();
    if (key) {
      settings[key] = value;
    }
  });

  return {
    currentJourneyName: settings['Current Journey Name'] || DEFAULT_JOURNEY_NAME
  };
}

function setRunSetting_(spreadsheet, settingName, settingValue) {
  var sheet = getOrCreateSheet_(spreadsheet, RUN_SETTINGS_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    sheet.setFrozenRows(1);
  }

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === settingName) {
      sheet.getRange(i + 1, 2).setValue(settingValue);
      sheet.autoResizeColumns(1, 2);
      return;
    }
  }

  sheet.appendRow([settingName, settingValue]);
  sheet.autoResizeColumns(1, 2);
}

function getJourneyNameForRow_(row, runSettings) {
  if (String(row.__journeyName || '').trim()) {
    return row.__journeyName;
  }

  return getFirstNonBlankValue_(row, [
    'Journey',
    'Journey Name',
    'Journey Title',
    'Workshop',
    'Workshop Name',
    'Course',
    'Course Name'
  ]) || runSettings.currentJourneyName || DEFAULT_JOURNEY_NAME;
}

function getFirstNonBlankValue_(row, keys) {
  for (var i = 0; i < keys.length; i++) {
    var value = getValue_(row, keys[i]);
    if (String(value || '').trim()) {
      return value;
    }
  }

  return '';
}

function writeMissingLookerPairings_(spreadsheet, representativeNames) {
  var sheet = getOrCreateSheet_(spreadsheet, MISSING_LOOKER_PAIRINGS_SHEET_NAME);
  var names = representativeNames
    .filter(function(name) {
      return String(name || '').trim();
    })
    .sort(function(a, b) {
      return String(a).localeCompare(String(b));
    });

  sheet.clear();
  sheet.getRange(1, 1, 1, MISSING_LOOKER_PAIRINGS_HEADERS.length).setValues([MISSING_LOOKER_PAIRINGS_HEADERS]);
  sheet.setFrozenRows(1);

  if (!names.length) {
    sheet.autoResizeColumns(1, MISSING_LOOKER_PAIRINGS_HEADERS.length);
    return;
  }

  sheet.getRange(2, 1, names.length, 1).setValues(names.map(function(name) {
    return [name];
  }));
  sheet.autoResizeColumns(1, MISSING_LOOKER_PAIRINGS_HEADERS.length);
}

function writeExceptions_(spreadsheet, exceptions) {
  var sheet = getOrCreateSheet_(spreadsheet, EXCEPTION_SHEET_NAME);
  var existingTracking = readExistingExceptionTracking_(sheet);

  sheet.clear();

  sheet.getRange(1, 1, 1, EXCEPTION_HEADERS.length).setValues([EXCEPTION_HEADERS]);
  sheet.setFrozenRows(1);

  if (!exceptions.length) {
    return;
  }

  var values = exceptions.map(function(row) {
    var tracking = existingTracking[exceptionTrackingKey_(row)] || {};

    return [
      row.reportRunAt,
      row.repName,
      row.repEmail,
      row.managerName,
      row.managerEmail,
      row.seniorName,
      row.seniorEmail,
      row.journeyName,
      row.simulationName,
      row.status,
      row.score,
      row.action,
      tracking.managerEmailSent || '',
      tracking.managerEmailSentAt || '',
      tracking.testSent || ''
    ];
  });

  sheet.getRange(2, 1, values.length, EXCEPTION_HEADERS.length).setValues(values);
  sheet.getRange(2, 11, values.length, 1).setNumberFormat('0%');
  sheet.getRange(2, 14, values.length, 1).setNumberFormat('m/d/yyyy h:mm AM/PM');
  sheet.autoResizeColumns(1, EXCEPTION_HEADERS.length);
}

function readExistingExceptionTracking_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(header) {
    return String(header).trim();
  });

  var col = buildColumnIndex_(headers);
  var required = [
    'Representative Email',
    'Journey',
    'Simulation',
    'Completion Status',
    'Required Follow-up Action'
  ];

  var hasRequired = required.every(function(header) {
    return col[header] !== undefined;
  });

  if (!hasRequired) {
    return {};
  }

  var tracking = {};

  values.slice(1).forEach(function(row) {
    var key = exceptionTrackingKey_({
      repEmail: row[col['Representative Email']],
      journeyName: row[col['Journey']],
      simulationName: row[col['Simulation']],
      status: row[col['Completion Status']],
      action: row[col['Required Follow-up Action']]
    });

    tracking[key] = {
      managerEmailSent: col['Manager Email Sent'] !== undefined ? row[col['Manager Email Sent']] : '',
      managerEmailSentAt: col['Manager Email Sent At'] !== undefined ? row[col['Manager Email Sent At']] : '',
      testSent: col['Test Sent'] !== undefined ? row[col['Test Sent']] : ''
    };
  });

  return tracking;
}

function exceptionTrackingKey_(row) {
  return [
    row.repEmail,
    row.journeyName,
    row.simulationName,
    row.status,
    row.action
  ].map(function(value) {
    return String(value || '').toLowerCase().trim();
  }).join('|');
}

function writeRunLog_(spreadsheet, rowsProcessed, exceptionsFound, emailsSent) {
  var sheet = getOrCreateSheet_(spreadsheet, RUN_LOG_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([[
      'Report Run At',
      'Rows Processed',
      'Exceptions Found',
      'Emails Sent'
    ]]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    rowsProcessed,
    exceptionsFound,
    emailsSent
  ]);
}

function buildManagerEmailBody_(managerName, rows, testMode, intendedManagerEmail, intendedCcRecipients) {
  var sections = buildEmailSections_(rows);
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - This batch would have gone to: ' + intendedManagerEmail);
    lines.push('CC would have included: ' + (intendedCcRecipients || 'None'));
    lines.push('');
  }

  lines.push('Hi' + (managerName ? ' ' + managerName : '') + ',');
  lines.push('');
  lines.push('Below is the latest ReflexAI status for simulations scheduled for your team through the current peak cadence week.');
  lines.push('Please use this video as a resource for navigating the ReflexAI Platform for further insights: ' + REFLEXAI_PLATFORM_RESOURCE_URL);
  lines.push('');

  appendPlainTextSection_(lines, COMPLETED_CLEARED_LABEL, sections.completedClearedGroups);
  appendPlainTextSection_(lines, COMPLETED_NOT_CLEARED_LABEL + ' - Priority', sections.lowScoreGroups);
  appendPlainTextSection_(lines, NOT_STARTED_LABEL, sections.incompleteGroups);

  lines.push('');
  lines.push('Thank you.');

  return lines.join('\n');
}

function buildManagerEmailHtml_(managerName, rows, testMode, intendedManagerEmail, managerMetrics, intendedCcRecipients, takeRate, takeRateChartCid) {
  var sections = buildEmailSections_(rows);
  var metricsBucket = managerMetrics && managerMetrics.bucket;
  var averageRows = managerMetrics && managerMetrics.rows ? managerMetrics.rows : rows;
  var completedClearedCount = metricsBucket ? metricsBucket.completedAbove : 0;
  var lowScoreCount = metricsBucket ? metricsBucket.completedBelow : countGroupedSimulations_(sections.lowScoreGroups);
  var incompleteCount = metricsBucket ? metricsBucket.notStarted : countGroupedSimulations_(sections.incompleteGroups);

  var bodyHtml = (testMode
      ? testBanner_(true, 'This batch would have gone to: ' + intendedManagerEmail + '. CC would have included: ' + (intendedCcRecipients || 'None'))
      : '') +
    introCard_(
      'Hi' + (managerName ? ' ' + managerName : '') + ',',
      'Below is the latest ReflexAI status for simulations that have been scheduled for your team through the current peak cadence week.'
    ) +
    '<div style="text-align:center;margin:18px 0 22px 0;">' +
      '<a href="' + REFLEXAI_PLATFORM_RESOURCE_URL + '" style="display:inline-block;background:#24205f;color:#ffffff;text-decoration:none;padding:12px 20px;border-radius:999px;font-weight:700;font-size:14px;">Watch ReflexAI Navigation Video</a>' +
      '<span style="display:inline-block;width:10px;"></span>' +
      '<a href="' + REFLEXAI_LOGIN_URL + '" style="display:inline-block;background:#ffffff;color:#24205f;text-decoration:none;padding:11px 20px;border-radius:999px;border:2px solid #24205f;font-weight:800;font-size:14px;">Open ReflexAI Login</a>' +
    '</div>' +
    sectionCard_(
      'TEAM SNAPSHOT',
      'Scheduled Simulations Included',
      metricTiles_([
        { label: COMPLETED_CLEARED_LABEL, value: completedClearedCount, color: COMPLETED_CLEARED_COLOR },
        { label: COMPLETED_NOT_CLEARED_LABEL, value: lowScoreCount, color: COMPLETED_NOT_CLEARED_COLOR },
        { label: NOT_STARTED_LABEL, value: incompleteCount, color: NOT_STARTED_COLOR }
      ]) +
      progressBar_(completedClearedCount, lowScoreCount, incompleteCount) +
      managerSimulationAveragesHtml_(averageRows)
    ) +
    buildTakeRateSectionHtml_(
      'OFFER TAKE RATE',
      'Current Offer Status',
      takeRate,
      takeRateChartCid,
      buildManagerOpenOfferDetailsHtml_(takeRate)
    ) +
    buildHtmlSection_(COMPLETED_NOT_CLEARED_LABEL + ' - Coaching Needed', sections.lowScoreGroups, true) +
    buildHtmlSection_(NOT_STARTED_LABEL + ' - Action Needed', sections.incompleteGroups, false) +
    buildHtmlSection_(COMPLETED_CLEARED_LABEL, sections.completedClearedGroups, false) +
    introCard_('Thank you.', 'Please use this report to prioritize coaching and completion follow-up.');

  return emailShell_(
    getCurrentMonthName_() + ' ReflexAI Simulation Follow-Up',
    EMAIL_SUBTITLE,
    bodyHtml
  );
}

function countGroupedSimulations_(groups) {
  return groups.reduce(function(total, group) {
    return total + group.simulations.length;
  }, 0);
}

function buildEmailSections_(rows) {
  var completedClearedRows = [];
  var lowScoreRows = [];
  var incompleteRows = [];

  rows.forEach(function(row) {
    if (isCompletedClearedEmailRow_(row)) {
      completedClearedRows.push(row);
    } else if (isBelowThresholdEmailRow_(row)) {
      lowScoreRows.push(row);
    } else {
      incompleteRows.push(row);
    }
  });

  return {
    completedClearedGroups: aggregateEmailRowsByRepJourney_(completedClearedRows),
    lowScoreGroups: aggregateEmailRowsByRepJourney_(lowScoreRows),
    incompleteGroups: aggregateEmailRowsByRepJourney_(incompleteRows)
  };
}

function aggregateEmailRowsByRepJourney_(rows) {
  var grouped = {};

  rows.forEach(function(row) {
    var key = [
      String(row.repEmail || row.repName || '').toLowerCase().trim(),
      String(row.journeyName || '').toLowerCase().trim()
    ].join('|');

    if (!grouped[key]) {
      grouped[key] = {
        repName: row.repName,
        journeyName: row.journeyName,
        simulations: [],
        minScorePercent: null
      };
    }

    var scorePercent = parseEmailScorePercent_(row.score);
    if (isFinite(scorePercent)) {
      grouped[key].minScorePercent = grouped[key].minScorePercent === null
        ? scorePercent
        : Math.min(grouped[key].minScorePercent, scorePercent);
    }

    grouped[key].simulations.push({
      simulationName: row.simulationName,
      assignedSimulationName: row.assignedSimulationName || row.simulationName,
      dueWeekLabel: row.dueWeekLabel || '',
      dateAssignedLabel: row.dateAssignedLabel || '',
      overdue: !!row.overdue,
      status: row.status,
      score: row.score,
      action: row.action
    });
  });

  return Object.keys(grouped)
    .map(function(key) {
      return grouped[key];
    })
    .sort(function(a, b) {
      return String(a.repName).localeCompare(String(b.repName));
    });
}

function appendPlainTextSection_(lines, title, groups) {
  if (!groups.length) {
    return;
  }

  lines.push(title);
  lines.push('');

  groups.forEach(function(group) {
    group.simulations.forEach(function(item, index) {
      lines.push([
        index === 0 ? group.repName : '',
        index === 0 ? group.journeyName : '',
        item.simulationName,
        item.dateAssignedLabel,
        item.status,
        formatScore_(item.score)
      ].join(' | '));
    });

    lines.push('');
  });
}

function buildHtmlSection_(title, groups, useScoreGradient) {
  if (!groups.length) {
    return '';
  }

  var tableRows = groups.map(function(group) {
    var rowStyle = '';

    var rowspan = group.simulations.length;
    return group.simulations.map(function(item, index) {
      var leadingCells = index === 0
        ? '<td rowspan="' + rowspan + '">' + escapeHtml_(group.repName) + '</td>' +
          '<td rowspan="' + rowspan + '">' + escapeHtml_(group.journeyName) + '</td>'
        : '';

      return '<tr' + rowStyle + '>' +
        leadingCells +
        '<td>' + escapeHtml_(item.assignedSimulationName || item.simulationName) + '</td>' +
        dateAssignedCell_(item) +
        '<td>' + statusBadge_(item.status, useScoreGradient ? 'warning' : '') + '</td>' +
        '<td>' + scoreBadge_(item.score) + '</td>' +
        '</tr>';
    }).join('');
  }).join('');

  return sectionCard_(
    title === COMPLETED_CLEARED_LABEL ? 'NO ACTION NEEDED' : (useScoreGradient ? 'COACHING NEEDED' : 'ACTION NEEDED'),
    title,
    styledTable_(['Representative', 'Journey', 'Assigned Sim', 'Date Assigned', 'Status', 'Score'], tableRows)
  );
}

function dateAssignedCell_(item) {
  var label = item.dateAssignedLabel || '';
  if (!item.overdue) {
    return '<td>' + escapeHtml_(label) + '</td>';
  }

  return '<td>' +
    '<span style="display:block;white-space:nowrap;">' + escapeHtml_(label) + '</span>' +
    '<span style="display:inline-block;margin-top:4px;border-radius:999px;background:#fff1f2;color:#b91c1c;border:1px solid #fecdd3;font-weight:900;font-size:11px;padding:3px 7px;white-space:nowrap;">Overdue</span>' +
    '</td>';
}

function statusBadge_(status, tone) {
  var value = String(status || '').trim() || 'No Status';
  var normalized = value.toLowerCase();
  var color = '#6a62d2';
  var background = '#f1efff';

  if (tone === 'warning') {
    if (normalized.indexOf('completed') !== -1) {
      value = 'Attempted';
    }
    color = '#9a6a00';
    background = '#fff8df';
  } else if (normalized.indexOf('completed') !== -1) {
    color = '#245bc5';
    background = '#eef5ff';
  } else if (normalized.indexOf('escalated') !== -1) {
    color = '#9a6a00';
    background = '#fff8df';
  } else if (normalized.indexOf('not') !== -1 || normalized.indexOf('progress') !== -1 || normalized.indexOf('started') !== -1) {
    color = '#b91c85';
    background = '#fff0fb';
  }

  return '<span style="display:inline-block;border-radius:999px;background:' + background + ';color:' + color + ';font-weight:800;font-size:12px;padding:5px 9px;white-space:nowrap;">' + escapeHtml_(value) + '</span>';
}

function scoreBadge_(score) {
  var scorePercent = parseEmailScorePercent_(score);
  var value = formatScore_(score);
  var color = '#6a62d2';
  var background = '#f1efff';

  if (isFinite(scorePercent)) {
    if (scorePercent >= PASSING_SCORE_PERCENT) {
      color = '#245bc5';
      background = '#eef5ff';
    } else {
      color = '#9a6a00';
      background = '#fff8df';
    }
  } else {
    color = '#b91c85';
    background = '#fff0fb';
  }

  return '<span style="display:inline-block;border-radius:999px;background:' + background + ';color:' + color + ';font-weight:900;font-size:12px;padding:5px 9px;white-space:nowrap;">' + escapeHtml_(value) + '</span>';
}

function summarizeActions_(simulations) {
  var allCleared = simulations.every(function(item) {
    return isCompletedClearedEmailRow_(item);
  });
  if (allCleared) return NO_FOLLOW_UP_ACTION;

  var hasLowScore = simulations.some(function(item) {
    return isBelowThresholdEmailRow_(item);
  });

  if (hasLowScore) {
    return RETAKE_SIMULATION_ACTION;
  }

  return COMPLETE_SIMULATION_ACTION;
}

function actionForOutcome_(category) {
  if (category === 'completedAbove') return NO_FOLLOW_UP_ACTION;
  if (category === 'completedBelow') return RETAKE_SIMULATION_ACTION;
  return COMPLETE_SIMULATION_ACTION;
}

function isCompletedClearedEmailRow_(row) {
  var scorePercent = parseEmailScorePercent_(row.score);
  return isFinite(scorePercent) && scorePercent >= PASSING_SCORE_PERCENT;
}

function isBelowThresholdEmailRow_(row) {
  var scorePercent = parseEmailScorePercent_(row.score);
  return isFinite(scorePercent) && scorePercent < PASSING_SCORE_PERCENT;
}

function parseEmailScorePercent_(score) {
  if (score === '' || score === null || score === undefined) {
    return NaN;
  }

  if (typeof score === 'number') {
    return score <= 1 ? score * 100 : score;
  }

  var parsed = Number(String(score).replace('%', '').trim());

  if (!isFinite(parsed)) {
    return NaN;
  }

  return parsed <= 1 ? parsed * 100 : parsed;
}

function scoreGradientColor_(scorePercent) {
  if (!isFinite(scorePercent)) {
    return '#ffffff';
  }

  var clamped = Math.max(0, Math.min(79, scorePercent));
  var ratio = clamped / 79;

  var start = { r: 244, g: 204, b: 204 };
  var end = { r: 255, g: 242, b: 204 };

  var r = Math.round(start.r + (end.r - start.r) * ratio);
  var g = Math.round(start.g + (end.g - start.g) * ratio);
  var b = Math.round(start.b + (end.b - start.b) * ratio);

  return 'rgb(' + r + ',' + g + ',' + b + ')';
}

function readSheetRows_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return [];
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  var headers = values[0].map(function(header) {
    return String(header).trim();
  });

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(value) {
        return value !== '';
      });
    })
    .map(function(row) {
      var record = {};

      headers.forEach(function(header, index) {
        record[header] = row[index];
      });

      return record;
    });
}

function loadReflexAiCsvRows_(spreadsheet, runSettings) {
  var legacyRows = [];
  var bulkRows = [];
  var defaultJourneyName = runSettings.currentJourneyName || DEFAULT_JOURNEY_NAME;

  spreadsheet.getSheets().forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName === CSV_DUMP_SHEET_NAME || sheetName.indexOf(CSV_DUMP_SHEET_PREFIX) !== 0) {
      return;
    }

    var journeyName = sheetName.slice(CSV_DUMP_SHEET_PREFIX.length).trim() || defaultJourneyName;
    bulkRows = bulkRows.concat(annotateRowsWithJourney_(readRowsFromSheet_(sheet), journeyName));
  });

  if (bulkRows.length) {
    return filterRemovedReps_(spreadsheet, bulkRows);
  }

  legacyRows = legacyRows.concat(annotateRowsWithJourney_(readSheetRows_(spreadsheet, CSV_DUMP_SHEET_NAME), defaultJourneyName));
  return filterRemovedReps_(spreadsheet, legacyRows);
}

function filterRemovedReps_(spreadsheet, rows) {
  var removalLookup = buildRemovedRepLookup_(spreadsheet);

  if (!removalLookup.hasEntries) {
    return rows;
  }

  return rows.filter(function(row) {
    var email = String(getValue_(row, 'User Email') || '').toLowerCase().trim();
    var name = normalizePersonKey_(getValue_(row, 'User Name'));

    if (email && removalLookup.emails[email]) {
      return false;
    }

    if (name && removalLookup.names[name]) {
      return false;
    }

    return true;
  });
}

function buildRemovedRepLookup_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, REMOVED_REPS_SHEET_NAME);
  var lookup = {
    emails: {},
    names: {},
    hasEntries: false
  };

  rows.forEach(function(row) {
    var email = String(getFirstNonBlankValue_(row, [
      'Representative Email',
      'Email',
      'User Email'
    ]) || '').toLowerCase().trim();
    var name = normalizePersonKey_(getFirstNonBlankValue_(row, [
      'Representative Name',
      'Name',
      'User Name'
    ]));

    if (email) {
      lookup.emails[email] = true;
      lookup.hasEntries = true;
    }

    if (name) {
      lookup.names[name] = true;
      lookup.hasEntries = true;
    }
  });

  return lookup;
}

function annotateRowsWithJourney_(rows, journeyName) {
  return rows.map(function(row) {
    row.__journeyName = getJourneyNameForRow_(row, { currentJourneyName: journeyName });
    return row;
  });
}

function readRowsFromSheet_(sheet) {
  if (!sheet) {
    return [];
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  var headers = values[0].map(function(header) {
    return String(header).trim();
  });

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(value) {
        return value !== '';
      });
    })
    .map(function(row) {
      var record = {};

      headers.forEach(function(header, index) {
        record[header] = row[index];
      });

      return record;
    });
}

function ensureExceptionTrackingHeaders_(sheet) {
  var headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), EXCEPTION_HEADERS.length))
    .getValues()[0];

  var names = headers.map(function(header) {
    return String(header).trim();
  });

  EXCEPTION_HEADERS.forEach(function(header) {
    if (names.indexOf(header) !== -1) {
      return;
    }

    var nextColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextColumn).setValue(header);
    names.push(header);
  });
}

function createSheetWithHeaders_(spreadsheet, sheetName, headers) {
  var sheet = getOrCreateSheet_(spreadsheet, sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function buildColumnIndex_(headers) {
  var index = {};

  headers.forEach(function(header, position) {
    if (header) {
      index[header] = position;
    }
  });

  return index;
}

function getValue_(row, key) {
  var value = row[key];

  if (value === null || value === undefined) {
    return '';
  }

  return value;
}

function parseScore_(value) {
  if (value === null || value === undefined || value === '') {
    return NaN;
  }

  var cleaned = String(value).replace('%', '').trim();
  return Number(cleaned);
}

function formatScore_(score) {
  if (score === '' || score === null || score === undefined || isNaN(score)) {
    return 'N/A';
  }

  if (typeof score === 'string' && score.indexOf('%') !== -1) {
    return score;
  }

  var number = Number(score);

  if (number <= 1) {
    return Math.round(number * 100) + '%';
  }

  return Math.round(number) + '%';
}

function formatPercent_(count, total) {
  if (!total) {
    return '0%';
  }

  return Math.round((count / total) * 100) + '%';
}

function formatDecimalPercent_(value) {
  if (!value || isNaN(value)) {
    return '0%';
  }

  return Math.round(value * 100) + '%';
}

function formatCountPercent_(count, percent) {
  return count + ' (' + formatDecimalPercent_(percent) + ')';
}

function formatOfferSentDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d/yyyy');
  }

  var raw = String(value || '').trim();
  if (!raw) return '';

  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'M/d/yyyy');
  }

  var match = raw.match(/^(\d{1,2}\/\d{1,2}\/\d{4})/);
  return match ? match[1] : raw;
}

function emailFromName_(name) {
  var override = EMAIL_NAME_OVERRIDES[normalizePersonKey_(name)];
  if (override) {
    return override;
  }

  var parts = String(name || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '')
    .split(/\s+/)
    .filter(Boolean);

  if (!parts.length) {
    return '';
  }

  return parts.join('.') + '@varsitytutors.com';
}

function getCurrentMonthName_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM');
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
