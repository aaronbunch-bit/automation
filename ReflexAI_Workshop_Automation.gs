var CSV_DUMP_SHEET_NAME = 'ReflexAI CSV Dump';
var CSV_DUMP_SHEET_PREFIX = 'ReflexAI CSV - ';
var LOOKER_MANAGER_LOOKUP_SHEET_NAME = 'Looker Manager Lookup';
var NAME_MATCH_OVERRIDES_SHEET_NAME = 'Name Match Overrides';
var MISSING_LOOKER_PAIRINGS_SHEET_NAME = 'Missing Looker Manager Pairing';
var EXCEPTION_SHEET_NAME = 'Simulation Exceptions';
var RUN_LOG_SHEET_NAME = 'Run Log';
var RUN_SETTINGS_SHEET_NAME = 'Run Settings';
var LOOKER_MANAGER_SOURCE_SPREADSHEET_ID = '1a6bE3cI-98tbAMyGizZwrsj3pQJdwXd1GMf5oxmFHUo';
var LOOKER_MANAGER_SOURCE_RANGE = "'Sales Roster Update.csv'!A:G";

var PASSING_SCORE_PERCENT = 80;
var DEFAULT_JOURNEY_NAME = 'High School Year Round Workshops - Jun';

var COMPLETE_SIMULATION_ACTION = 'Ask representative to complete this simulation and schedule time via Assembled for representative to complete the simulation adhering to capacity constraints.';
var RETAKE_SIMULATION_ACTION = 'Ask representative to retake this simulation and coach on missed skills.';
var REFLEXAI_PLATFORM_RESOURCE_URL = 'https://drive.google.com/file/d/18X5z6iRGRk-fKY4bAvIxswwys3ne2z3r/view';
var REFLEXAI_LOGIN_URL = 'https://varsitytutors.reflexai.com/home';
var COMPLETED_CLEARED_LABEL = 'Completed & Cleared 80% Threshold';
var COMPLETED_NOT_CLEARED_LABEL = 'Completed & Not Cleared 80% Threshold';
var NOT_STARTED_LABEL = 'Not Started';

var TEST_EMAIL_RECIPIENTS = [
  'aaron.bunch@varsitytutors.com'
];

var DIRECTOR_EMAIL_RECIPIENTS = [
  'joshua.langford@varsitytutors.com',
  'taylor.wisnasky@varsitytutors.com',
  'aaron.bunch@varsitytutors.com'
];

var SENIOR_LEADER_SUPERGROUP_RECIPIENTS = [
  {
    supergroupName: 'Prof Certs',
    leaderName: 'Yago Lupi',
    email: 'yago.lupi@varsitytutors.com'
  },
  {
    supergroupName: 'ELD',
    leaderName: 'Aftynn Peters',
    email: 'aftynn.peters@varsitytutors.com'
  },
  {
    supergroupName: 'High School',
    leaderName: 'Aftynn Peters',
    email: 'aftynn.peters@varsitytutors.com'
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
  'Manager ID',
  'Manager',
  'Regional Director',
  'Work Group',
  'Sub Group',
  'Group'
];

var MISSING_LOOKER_PAIRINGS_HEADERS = [
  'Representative'
];

var NAME_MATCH_OVERRIDES_HEADERS = [
  'ReflexAI Name',
  'Looker Name'
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
  SpreadsheetApp.getUi()
    .createMenu('ReflexAI Reporting')
    .addItem('Set Up Sheets', 'createSetupSheets')
    .addItem('Connect Looker Report', 'connectLookerManagerImport')
    .addItem('Set Current Journey Name', 'setCurrentJourneyName')
    .addItem('Run CSV Dump', 'runWeeklySimulationExceptionReport')
    .addItem('Test Manager Email', 'sendTestManagerExceptionEmails')
    .addItem('Test Senior Leader Email', 'sendTestSeniorLeadershipRecap')
    .addItem('Test Director Email', 'sendTestDirectorEmail')
    .addItem('Send Manager Email', 'sendManagerExceptionEmails')
    .addItem('Send Senior Leader Email', 'sendSeniorLeadershipRecap')
    .addItem('Send Director Email', 'sendDirectorEmail')
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
  sheet.autoResizeColumns(1, 7);

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
  var sheet = spreadsheet.getSheetByName(EXCEPTION_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No exception rows found. Run the report first.');
    return;
  }

  ensureExceptionTrackingHeaders_(sheet);

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(header) {
    return String(header).trim();
  });

  var col = buildColumnIndex_(headers);

  var grouped = {};
  var skippedAlreadyManagerSent = 0;
  var previouslyTestSent = 0;
  var skippedNoManagerEmail = 0;

  values.slice(1).forEach(function(row, offset) {
    var sheetRowNumber = offset + 2;

    var managerEmail = String(row[col['Manager Email']] || '').toLowerCase().trim();
    var sentStatus = String(row[col['Manager Email Sent']] || '').toLowerCase().trim();
    var testSentStatus = String(row[col['Test Sent']] || '').toLowerCase().trim();
    var action = String(row[col['Required Follow-up Action']] || '').trim();

    if (!action) {
      return;
    }

    if (sentStatus === 'yes') {
      skippedAlreadyManagerSent++;
      return;
    }

    if (testMode && testSentStatus === 'y') {
      previouslyTestSent++;
    }

    if (!managerEmail) {
      skippedNoManagerEmail++;
      return;
    }

    if (!grouped[managerEmail]) {
      grouped[managerEmail] = {
        managerName: row[col['Manager']] || '',
        recipientEmails: {},
        rows: []
      };
    }

    grouped[managerEmail].recipientEmails[managerEmail] = true;
    var seniorEmail = String(row[col['Senior / Team Lead Email']] || '').toLowerCase().trim();
    if (seniorEmail) {
      grouped[managerEmail].recipientEmails[seniorEmail] = true;
    }

    grouped[managerEmail].rows.push({
      sheetRowNumber: sheetRowNumber,
      repName: row[col['Representative']] || '',
      repEmail: row[col['Representative Email']] || '',
      journeyName: row[col['Journey']] || '',
      simulationName: row[col['Simulation']] || '',
      status: row[col['Completion Status']] || '',
      score: row[col['Score']],
      action: action
    });
  });

  var managerEmails = Object.keys(grouped);

  if (!managerEmails.length) {
    SpreadsheetApp.getUi().alert(
      'No unsent notifications found.\n\n' +
      'Already manager-sent rows skipped: ' +
      skippedAlreadyManagerSent +
      '\nPreviously test-sent rows included for retest: ' +
      previouslyTestSent +
      '\nRows missing manager email: ' +
      skippedNoManagerEmail
    );
    return;
  }

  var sentAt = new Date();
  var batchesSent = 0;
  var rowsMarked = 0;

  managerEmails.forEach(function(managerEmail) {
    var batch = grouped[managerEmail];

    var recipients = testMode
      ? TEST_EMAIL_RECIPIENTS.join(',')
      : Object.keys(batch.recipientEmails).join(',');

    MailApp.sendEmail({
      to: recipients,
      subject: (testMode ? '[TEST] ' : '') + 'ReflexAI Weekly Simulation Follow-Up',
      body: buildManagerEmailBody_(batch.managerName, batch.rows, testMode, managerEmail),
      htmlBody: buildManagerEmailHtml_(batch.managerName, batch.rows, testMode, managerEmail)
    });

    batchesSent++;

    batch.rows.forEach(function(item) {
      if (testMode) {
        sheet.getRange(item.sheetRowNumber, col['Test Sent'] + 1).setValue('Y');
      } else {
        sheet.getRange(item.sheetRowNumber, col['Manager Email Sent'] + 1).setValue('Yes');
        sheet.getRange(item.sheetRowNumber, col['Manager Email Sent At'] + 1).setValue(sentAt);
      }

      rowsMarked++;
    });
  });

  if (!testMode) {
    sheet
      .getRange(2, col['Manager Email Sent At'] + 1, Math.max(sheet.getLastRow() - 1, 1), 1)
      .setNumberFormat('m/d/yyyy h:mm AM/PM');
  }

  SpreadsheetApp.getUi().alert(
    (testMode ? 'Test emails sent.' : 'Manager emails sent.') +
    '\n\nBatches sent: ' +
    batchesSent +
    '\nRows marked: ' +
    rowsMarked +
    '\nAlready manager-sent rows skipped: ' +
    skippedAlreadyManagerSent +
    '\nPreviously test-sent rows included for retest: ' +
    previouslyTestSent +
    '\nRows missing manager email: ' +
    skippedNoManagerEmail
  );
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
  var sentCount = 0;

  SENIOR_LEADER_SUPERGROUP_RECIPIENTS.forEach(function(config) {
    var filteredRows = filterRowsForSupergroup_(csvRows, config.supergroupName, runSettings);
    var recap = buildSeniorLeadershipRecap_(filteredRows, managerRoster, {
      currentJourneyName: config.supergroupName
    });
    recap.supergroupName = config.supergroupName;
    var recipients = testMode ? TEST_EMAIL_RECIPIENTS : [config.email];
    var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Supergroup Recap';

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: buildSeniorLeadershipRecapText_(recap, testMode, config),
      htmlBody: buildSeniorLeadershipRecapHtml_(recap, testMode, config)
    });

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
  var recap = buildDirectorRecap_(csvRows, managerRoster, runSettings);
  var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Director Recap';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: buildDirectorEmailText_(recap, testMode),
    htmlBody: buildDirectorEmailHtml_(recap, testMode)
  });

  SpreadsheetApp.getUi().alert(
    'Director recap sent.\n\nRecipients: ' +
    recipients.join(', ') +
    '\nRows included: ' +
    recap.totalRows
  );
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
    byManager: {}
  };

  csvRows.forEach(function(row) {
    var userName = getValue_(row, 'User Name');
    var userEmail = getValue_(row, 'User Email');
    var managerInfo = getManagerInfoForRep_(managerRoster, userEmail, userName);

    if (!managerInfo.managerName) {
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
  return recap;
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

function emailShell_(title, subtitle, bodyHtml) {
  return '<div style="margin:0;padding:0;background:linear-gradient(135deg,#fbf2ff 0%,#edf4ff 52%,#f7f5ff 100%);font-family:Arial,Helvetica,sans-serif;color:#24205f;">' +
    '<div style="max-width:960px;margin:0 auto;padding:24px;">' +
      '<div style="background:#ffffff;border-radius:22px;overflow:hidden;border:1px solid #dedaf8;box-shadow:0 14px 34px rgba(36,32,95,0.16);">' +
        '<div style="height:12px;background:linear-gradient(90deg,#ffcc33 0%,#ff6b9c 25%,#ff3ec8 48%,#7957ff 72%,#20d5d2 100%);"></div>' +
        '<div style="background:linear-gradient(135deg,#24205f 0%,#353082 68%,#5c4be8 100%);color:#ffffff;padding:30px 34px;">' +
          '<div style="display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.28);border-radius:999px;padding:6px 12px;font-size:12px;letter-spacing:1.2px;text-transform:uppercase;color:#ffffff;font-weight:800;">Varsity Tutors</div>' +
          '<div style="display:inline-block;margin-left:8px;background:rgba(32,213,210,0.14);border:1px solid rgba(32,213,210,0.45);border-radius:999px;padding:6px 12px;font-size:12px;color:#dffefe;font-weight:800;">ReflexAI Signals</div>' +
          '<div style="font-size:28px;line-height:34px;font-weight:800;margin-top:8px;">' + escapeHtml_(title) + '</div>' +
          '<div style="font-size:14px;line-height:20px;color:#e8e6ff;margin-top:8px;">' + escapeHtml_(subtitle || '') + '</div>' +
          '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;border-collapse:collapse;"><tr>' +
            '<td style="width:33%;padding-right:8px;"><div style="height:5px;border-radius:999px;background:#ffcc33;"></div></td>' +
            '<td style="width:34%;padding:0 8px;"><div style="height:5px;border-radius:999px;background:#ff3ec8;"></div></td>' +
            '<td style="width:33%;padding-left:8px;"><div style="height:5px;border-radius:999px;background:#20d5d2;"></div></td>' +
          '</tr></table>' +
        '</div>' +
        '<div style="padding:28px 32px;background:#f7f5ff;">' + bodyHtml + '</div>' +
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

  return '<div style="background:#ffffff;border:1px solid #dedaf8;border-radius:18px;padding:0;margin:0 0 22px 0;overflow:hidden;box-shadow:0 10px 26px rgba(36,32,95,0.10);">' +
    '<div style="height:5px;background:' + accent + ';"></div>' +
    '<div style="padding:16px 20px;border-bottom:1px solid #ebe8ff;background:linear-gradient(90deg,#ffffff 0%,#faf8ff 100%);">' +
      '<div style="font-size:11px;letter-spacing:1.1px;text-transform:uppercase;color:' + accent + ';font-weight:900;">' + escapeHtml_(eyebrow || '') + '</div>' +
      '<div style="font-size:20px;line-height:26px;color:#24205f;font-weight:800;margin-top:4px;">' + escapeHtml_(title) + '</div>' +
    '</div>' +
    '<div style="padding:18px 20px;">' + contentHtml + '</div>' +
    '</div>';
}

function metricTiles_(tiles) {
  return '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>' +
    tiles.map(function(tile) {
      return '<td style="width:33.33%;padding:6px;vertical-align:top;">' +
        '<div style="border:1px solid #ebe8ff;border-radius:16px;padding:18px 12px;background:linear-gradient(180deg,#ffffff 0%,#fbfaff 100%);text-align:center;box-shadow:0 5px 14px rgba(36,32,95,0.06);">' +
          '<div style="height:4px;border-radius:999px;background:' + tile.color + ';margin:0 auto 12px auto;width:52px;"></div>' +
          '<div style="font-size:28px;font-weight:900;color:' + tile.color + ';line-height:32px;text-align:center;">' + escapeHtml_(String(tile.value)) + '</div>' +
          '<div style="font-size:12px;line-height:17px;color:#4d49a3;font-weight:700;margin-top:6px;">' + escapeHtml_(tile.label) + '</div>' +
        '</div>' +
      '</td>';
    }).join('') +
    '</tr></table>';
}

function styledTable_(headers, bodyRowsHtml) {
  var styledBody = bodyRowsHtml
    .replace(/<tr>/g, '<tr style="background:#ffffff;">')
    .replace(/<td([^>]*)>/g, '<td$1 style="padding:11px 10px;border-bottom:1px solid #ebe8ff;vertical-align:middle;">');

  return '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;color:#24205f;">' +
    '<thead><tr>' +
      headers.map(function(header) {
        return '<th style="text-align:left;background:linear-gradient(180deg,#f3f0ff 0%,#ebe8ff 100%);color:#24205f;padding:12px 10px;border-bottom:1px solid #dedaf8;font-weight:900;">' + escapeHtml_(header) + '</th>';
      }).join('') +
    '</tr></thead>' +
    '<tbody>' + styledBody + '</tbody>' +
    '</table>';
}

function accentColorForSection_(eyebrow, title) {
  var text = String(eyebrow || '') + ' ' + String(title || '');
  text = text.toLowerCase();

  if (text.indexOf('priority') !== -1 || text.indexOf('not cleared') !== -1) return '#f59e0b';
  if (text.indexOf('action') !== -1 || text.indexOf('not started') !== -1) return '#ef4444';
  if (text.indexOf('snapshot') !== -1 || text.indexOf('cleared') !== -1) return '#1f9d55';
  if (text.indexOf('manager') !== -1) return '#7957ff';
  return '#6a62d2';
}

function buildDirectorEmailHtml_(recap, testMode) {
  var bodyHtml = testBanner_(testMode, 'Director recap preview.') +
    introCard_(
      'Hi Directors,',
      'Below is the ' + getCurrentMonthName_() + ' ReflexAI director recap thus far for ' + recap.supergroupName + '.'
    ) +
    sectionCard_(
      'CONSUMER SALES SNAPSHOT',
      'Overall Simulation Counts',
      metricTiles_([
        { label: COMPLETED_CLEARED_LABEL, value: formatCountPercent_(recap.company.completedAbove, recap.company.completedAbovePercent), color: '#1f9d55' },
        { label: COMPLETED_NOT_CLEARED_LABEL, value: formatCountPercent_(recap.company.completedBelow, recap.company.completedBelowPercent), color: '#f59e0b' },
        { label: NOT_STARTED_LABEL, value: formatCountPercent_(recap.company.notStarted, recap.company.notStartedPercent), color: '#ef4444' }
      ])
    ) +
    buildDirectorTable_('Supergroup Breakdown', recap.supergroupRows.concat([companySummaryRow_(recap.company)]), 'Supergroup') +
    buildDirectorTable_('Manager Breakdown', recap.managerRows, 'Manager');

  return emailShell_(
    getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Director Recap',
    'Varsity Tutors Coaching Enablement',
    bodyHtml
  );
}

function buildDirectorTable_(title, rows, firstColumnLabel) {
  var tableRows = rows.map(function(row) {
    var rowStyle = row.isCompanySummary ? ' style="background-color:#e7f7ec;font-weight:bold;"' : '';

    return '<tr' + rowStyle + '>' +
      '<td>' + escapeHtml_(row.name || 'Company') + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedAbove, row.completedAbovePercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedBelow, row.completedBelowPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.notStarted, row.notStartedPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatScore_(row.completedAverageScore)) + '</td>' +
      '<td>' + row.total + '</td>' +
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
      'Total'
    ], tableRows)
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

    if (!managerInfo.managerName) {
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

  return lines.join('\n');
}

function buildSeniorLeadershipRecapHtml_(recap, testMode, recipientConfig) {
  var intended = recipientConfig && recipientConfig.email ? 'This email would have gone to: ' + recipientConfig.email + '.' : '';
  var bodyHtml = testBanner_(testMode, 'Senior leadership recap preview. ' + intended) +
    introCard_(
      'Hi ' + ((recipientConfig && recipientConfig.leaderName) || 'Senior Leaders') + ',',
      'Below is the ' + getCurrentMonthName_() + ' ReflexAI recap thus far for ' + recap.supergroupName + '.'
    ) +
    buildLeadershipCountsTable_(recap) +
    buildSimulationAverageTable_(recap.simulationAverages) +
    buildManagerRecapTable_(recap.managerRows);

  return emailShell_(
    getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Supergroup Recap',
    'Varsity Tutors Coaching Enablement',
    bodyHtml
  );
}

function buildLeadershipCountsTable_(recap) {
  return sectionCard_(
    'SUPERGROUP SNAPSHOT',
    'Overall Simulation Counts',
    metricTiles_([
      { label: COMPLETED_CLEARED_LABEL, value: recap.counts.completedAbove, color: '#1f9d55' },
      { label: COMPLETED_NOT_CLEARED_LABEL, value: recap.counts.completedBelow, color: '#f59e0b' },
      { label: NOT_STARTED_LABEL, value: recap.counts.notStarted, color: '#ef4444' }
    ])
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

function buildLookerManagerRoster_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, LOOKER_MANAGER_LOOKUP_SHEET_NAME);
  var roster = {};
  roster.__candidates = [];
  var exactByName = {};
  var firstLastCandidates = {};
  var firstInitialLastCandidates = {};

  rows.forEach(function(row) {
    var repName = getValue_(row, 'Manager');
    var repEmail = getFirstNonBlankValue_(row, [
      'Email',
      'Employee Email',
      'Manager Email'
    ]);
    var managerName = getFirstNonBlankValue_(row, [
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
      seniorName: managerName,
      seniorEmail: managerEmail
    };

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

function buildManagerEmailBody_(managerName, rows, testMode, intendedManagerEmail) {
  var sections = buildEmailSections_(rows);
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - This batch would have gone to: ' + intendedManagerEmail);
    lines.push('');
  }

  lines.push('Hi' + (managerName ? ' ' + managerName : '') + ',');
  lines.push('');
  lines.push('Below are ReflexAI simulation follow-up items for your team.');
  lines.push('Only ' + NOT_STARTED_LABEL + ' or ' + COMPLETED_NOT_CLEARED_LABEL + ' simulations are included.');
  lines.push('Please use this video as a resource for navigating the ReflexAI Platform for further insights: ' + REFLEXAI_PLATFORM_RESOURCE_URL);
  lines.push('');

  appendPlainTextSection_(lines, COMPLETED_NOT_CLEARED_LABEL + ' - Priority', sections.lowScoreGroups);
  appendPlainTextSection_(lines, NOT_STARTED_LABEL, sections.incompleteGroups);

  lines.push('');
  lines.push('Thank you.');

  return lines.join('\n');
}

function buildManagerEmailHtml_(managerName, rows, testMode, intendedManagerEmail) {
  var sections = buildEmailSections_(rows);
  var lowScoreCount = countGroupedSimulations_(sections.lowScoreGroups);
  var incompleteCount = countGroupedSimulations_(sections.incompleteGroups);

  var bodyHtml = (testMode
      ? testBanner_(true, 'This batch would have gone to: ' + intendedManagerEmail)
      : '') +
    introCard_(
      'Hi' + (managerName ? ' ' + managerName : '') + ',',
      'Below are ReflexAI simulation follow-up items for your team. Only ' + NOT_STARTED_LABEL + ' or ' + COMPLETED_NOT_CLEARED_LABEL + ' simulations are included.'
    ) +
    sectionCard_(
      'TEAM SNAPSHOT',
      'Follow-Up Items Included',
      metricTiles_([
        { label: COMPLETED_NOT_CLEARED_LABEL, value: lowScoreCount, color: '#f59e0b' },
        { label: NOT_STARTED_LABEL, value: incompleteCount, color: '#ef4444' },
        { label: 'Total Follow-Up Items', value: lowScoreCount + incompleteCount, color: '#6a62d2' }
      ])
    ) +
    '<div style="text-align:center;margin:18px 0 22px 0;">' +
      '<a href="' + REFLEXAI_PLATFORM_RESOURCE_URL + '" style="display:inline-block;background:#24205f;color:#ffffff;text-decoration:none;padding:12px 20px;border-radius:999px;font-weight:700;font-size:14px;">Watch ReflexAI Navigation Video</a>' +
      '<span style="display:inline-block;width:10px;"></span>' +
      '<a href="' + REFLEXAI_LOGIN_URL + '" style="display:inline-block;background:#ff3ec8;color:#ffffff;text-decoration:none;padding:12px 20px;border-radius:999px;font-weight:700;font-size:14px;">Open ReflexAI Login</a>' +
    '</div>' +
    buildHtmlSection_(COMPLETED_NOT_CLEARED_LABEL + ' - Priority', sections.lowScoreGroups, true) +
    buildHtmlSection_(NOT_STARTED_LABEL, sections.incompleteGroups, false) +
    introCard_('Thank you.', 'Please use this report to prioritize coaching and completion follow-up.');

  return emailShell_(
    'ReflexAI Weekly Simulation Follow-Up',
    'Varsity Tutors Coaching Enablement',
    bodyHtml
  );
}

function countGroupedSimulations_(groups) {
  return groups.reduce(function(total, group) {
    return total + group.simulations.length;
  }, 0);
}

function buildEmailSections_(rows) {
  var lowScoreRows = [];
  var incompleteRows = [];

  rows.forEach(function(row) {
    if (isBelowThresholdEmailRow_(row)) {
      lowScoreRows.push(row);
    } else {
      incompleteRows.push(row);
    }
  });

  return {
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
        item.status,
        formatScore_(item.score),
        index === 0 ? summarizeActions_(group.simulations) : ''
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
    var action = summarizeActions_(group.simulations);

    return group.simulations.map(function(item, index) {
      var leadingCells = index === 0
        ? '<td rowspan="' + rowspan + '">' + escapeHtml_(group.repName) + '</td>' +
          '<td rowspan="' + rowspan + '">' + escapeHtml_(group.journeyName) + '</td>'
        : '';

      var actionCell = index === 0
        ? '<td rowspan="' + rowspan + '">' + escapeHtml_(action) + '</td>'
        : '';

      return '<tr' + rowStyle + '>' +
        leadingCells +
        '<td>' + escapeHtml_(item.simulationName) + '</td>' +
        '<td>' + statusBadge_(item.status) + '</td>' +
        '<td>' + scoreBadge_(item.score) + '</td>' +
        actionCell +
        '</tr>';
    }).join('');
  }).join('');

  return sectionCard_(
    useScoreGradient ? 'PRIORITY FOLLOW-UP' : 'ACTION NEEDED',
    title,
    styledTable_(['Representative', 'Journey', 'Simulation', 'Status', 'Score', 'Follow-up Action'], tableRows)
  );
}

function statusBadge_(status) {
  var value = String(status || '').trim() || 'No Status';
  var normalized = value.toLowerCase();
  var color = '#6a62d2';
  var background = '#f1efff';

  if (normalized.indexOf('completed') !== -1) {
    color = '#1f9d55';
    background = '#e7f7ec';
  } else if (normalized.indexOf('not') !== -1 || normalized.indexOf('progress') !== -1 || normalized.indexOf('started') !== -1) {
    color = '#ef4444';
    background = '#fff1f1';
  }

  return '<span style="display:inline-block;border-radius:999px;background:' + background + ';color:' + color + ';font-weight:800;font-size:12px;padding:5px 9px;">' + escapeHtml_(value) + '</span>';
}

function scoreBadge_(score) {
  var scorePercent = parseEmailScorePercent_(score);
  var value = formatScore_(score);
  var color = '#6a62d2';
  var background = '#f1efff';

  if (isFinite(scorePercent)) {
    if (scorePercent >= PASSING_SCORE_PERCENT) {
      color = '#1f9d55';
      background = '#e7f7ec';
    } else {
      color = '#f59e0b';
      background = '#fff7e6';
    }
  }

  return '<span style="display:inline-block;border-radius:999px;background:' + background + ';color:' + color + ';font-weight:900;font-size:12px;padding:5px 9px;">' + escapeHtml_(value) + '</span>';
}

function summarizeActions_(simulations) {
  var hasLowScore = simulations.some(function(item) {
    return isBelowThresholdEmailRow_(item);
  });

  if (hasLowScore) {
    return RETAKE_SIMULATION_ACTION;
  }

  return COMPLETE_SIMULATION_ACTION;
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
    return bulkRows;
  }

  legacyRows = legacyRows.concat(annotateRowsWithJourney_(readSheetRows_(spreadsheet, CSV_DUMP_SHEET_NAME), defaultJourneyName));
  return legacyRows;
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
    return 'No score';
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

function emailFromName_(name) {
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
