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
  sendSeniorLeadershipRecap_(TEST_EMAIL_RECIPIENTS, true);
}

function sendSeniorLeadershipRecap() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var recipients = getSeniorLeaderRecipients_(spreadsheet);

  if (!recipients.length) {
    SpreadsheetApp.getUi().alert('No senior leader emails found in column G of "Simulation Exceptions". Run the report first and confirm senior leader emails are populated.');
    return;
  }

  sendSeniorLeadershipRecap_(recipients, false);
}

function sendSeniorLeadershipRecap_(recipients, testMode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runSettings = getRunSettings_(spreadsheet);
  var csvRows = loadReflexAiCsvRows_(spreadsheet, runSettings);

  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import ReflexAI CSV data into "' + CSV_DUMP_SHEET_NAME + '" or tabs named "' + CSV_DUMP_SHEET_PREFIX + '[Journey Name]".');
    return;
  }

  var managerRoster = buildLookerManagerRoster_(spreadsheet);
  var recap = buildSeniorLeadershipRecap_(csvRows, managerRoster, runSettings);
  var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ' + recap.supergroupName + ' ReflexAI Supergroup Recap';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: buildSeniorLeadershipRecapText_(recap, testMode),
    htmlBody: buildSeniorLeadershipRecapHtml_(recap, testMode)
  });

  SpreadsheetApp.getUi().alert(
    'Senior leadership recap sent.\n\nRecipients: ' +
    recipients.join(', ') +
    '\nRows included: ' +
    recap.totalRows
  );
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

function buildDirectorEmailHtml_(recap, testMode) {
  return (testMode ? '<p><strong>TEST MODE</strong> - Director recap preview.</p>' : '') +
    '<p>Hi Directors,</p>' +
    '<p>Below is the ' + escapeHtml_(getCurrentMonthName_()) + ' ReflexAI director recap thus far for ' + escapeHtml_(recap.supergroupName) + '.</p>' +
    buildDirectorTable_('Supergroup Breakdown', recap.supergroupRows.concat([companySummaryRow_(recap.company)])) +
    buildDirectorTable_('Manager Breakdown', recap.managerRows);
}

function buildDirectorTable_(title, rows) {
  var tableRows = rows.map(function(row) {
    var rowStyle = row.isCompanySummary ? ' style="background-color:#d9ead3;font-weight:bold;"' : '';

    return '<tr' + rowStyle + '>' +
      '<td>' + escapeHtml_(row.name || 'Company') + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedAbove, row.completedAbovePercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.completedBelow, row.completedBelowPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatCountPercent_(row.notStarted, row.notStartedPercent)) + '</td>' +
      '<td>' + escapeHtml_(formatScore_(row.completedAverageScore)) + '</td>' +
      '<td>' + row.total + '</td>' +
      '</tr>';
  }).join('');

  return '<h3>' + escapeHtml_(title) + '</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>' + escapeHtml_(title === 'Manager Breakdown' ? 'Manager' : 'Segment') + '</th>' +
    '<th>' + COMPLETED_CLEARED_LABEL + '</th>' +
    '<th>' + COMPLETED_NOT_CLEARED_LABEL + '</th>' +
    '<th>' + NOT_STARTED_LABEL + '</th>' +
    '<th>Average Completed Score</th>' +
    '<th>Total</th></tr></thead>' +
    '<tbody>' + tableRows + '</tbody></table>';
}

function companySummaryRow_(company) {
  var row = {};
  Object.keys(company).forEach(function(key) {
    row[key] = company[key];
  });
  row.name = 'Company';
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

function buildSeniorLeadershipRecapText_(recap, testMode) {
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - Senior leadership recap preview.');
    lines.push('');
  }

  lines.push('Hi Senior Leaders,');
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

function buildSeniorLeadershipRecapHtml_(recap, testMode) {
  return (testMode ? '<p><strong>TEST MODE</strong> - Senior leadership recap preview.</p>' : '') +
    '<p>Hi Senior Leaders,</p>' +
    '<p>Below is the ' + escapeHtml_(getCurrentMonthName_()) + ' ReflexAI recap thus far for ' + escapeHtml_(recap.supergroupName) + '.</p>' +
    buildLeadershipCountsTable_(recap) +
    buildSimulationAverageTable_(recap.simulationAverages) +
    buildManagerRecapTable_(recap.managerRows);
}

function buildLeadershipCountsTable_(recap) {
  return '<h3>Overall Simulation Counts</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Metric</th><th>Count</th></tr></thead>' +
    '<tbody>' +
    '<tr><td>' + NOT_STARTED_LABEL + '</td><td>' + recap.counts.notStarted + '</td></tr>' +
    '<tr><td>' + COMPLETED_NOT_CLEARED_LABEL + '</td><td>' + recap.counts.completedBelow + '</td></tr>' +
    '<tr><td>' + COMPLETED_CLEARED_LABEL + '</td><td>' + recap.counts.completedAbove + '</td></tr>' +
    '</tbody></table>';
}

function buildSimulationAverageTable_(simulationAverages) {
  var rows = simulationAverages.map(function(item) {
    return '<tr>' +
      '<td>' + escapeHtml_(item.simulationName) + '</td>' +
      '<td>' + item.completedCount + '</td>' +
      '<td>' + escapeHtml_(formatScore_(item.averageScore)) + '</td>' +
      '</tr>';
  }).join('');

  return '<h3>Average Score on Completed Simulations by Simulation</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Simulation</th><th>Completed Count</th><th>Average Score</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>';
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

  return '<h3>Manager Breakdown</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Manager</th><th>% ' + COMPLETED_CLEARED_LABEL + '</th><th>% ' + COMPLETED_NOT_CLEARED_LABEL + '</th><th>% ' + NOT_STARTED_LABEL + '</th><th>Total Simulations</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>';
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

  return (testMode
      ? '<p><strong>TEST MODE</strong> - This batch would have gone to: ' +
        escapeHtml_(intendedManagerEmail) +
        '</p>'
      : '') +
    '<p>Hi' + (managerName ? ' ' + escapeHtml_(managerName) : '') + ',</p>' +
    '<p>Below are ReflexAI simulation follow-up items for your team. Only ' + escapeHtml_(NOT_STARTED_LABEL) + ' or ' + escapeHtml_(COMPLETED_NOT_CLEARED_LABEL) + ' simulations are included.</p>' +
    '<p>Please use <a href="' + REFLEXAI_PLATFORM_RESOURCE_URL + '">this video</a> as a resource for navigating the ReflexAI Platform for further insights.</p>' +
    buildHtmlSection_(COMPLETED_NOT_CLEARED_LABEL + ' - Priority', sections.lowScoreGroups, true) +
    buildHtmlSection_(NOT_STARTED_LABEL, sections.incompleteGroups, false) +
    '<p>Thank you.</p>';
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
    var rowStyle = useScoreGradient
      ? ' style="background-color:' + scoreGradientColor_(group.minScorePercent) + ';"'
      : '';

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
        '<td>' + escapeHtml_(item.status) + '</td>' +
        '<td>' + escapeHtml_(formatScore_(item.score)) + '</td>' +
        actionCell +
        '</tr>';
    }).join('');
  }).join('');

  return '<h3>' + escapeHtml_(title) + '</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Representative</th><th>Journey</th><th>Simulation</th><th>Status</th><th>Score</th><th>Follow-up Action</th></tr></thead>' +
    '<tbody>' + tableRows + '</tbody>' +
    '</table>';
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
