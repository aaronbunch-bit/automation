var CSV_DUMP_SHEET_NAME = 'ReflexAI CSV Dump';
var MANAGER_ROSTER_SHEET_NAME = 'Manager Roster';
var EXCEPTION_SHEET_NAME = 'Simulation Exceptions';
var RUN_LOG_SHEET_NAME = 'Run Log';
var RUN_SETTINGS_SHEET_NAME = 'Run Settings';

var PASSING_SCORE_PERCENT = 80;
var DEFAULT_JOURNEY_NAME = 'High School Year Round Workshops - Jun';

var COMPLETE_SIMULATION_ACTION = 'Ask representative to complete this simulation and schedule time via Assembled for representative to complete the simulation adhering to capacity constraints.';
var RETAKE_SIMULATION_ACTION = 'Ask representative to retake this simulation and coach on missed skills.';

var TEST_EMAIL_RECIPIENTS = [
  'aaron.bunch@varsitytutors.com',
  'robert.sorrell@varsitytutors.com'
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
    .addItem('Create setup sheets', 'createSetupSheets')
    .addItem('Set current journey name', 'setCurrentJourneyName')
    .addItem('Run report from CSV dump now', 'runWeeklySimulationExceptionReport')
    .addItem('Test senior leadership recap to Aaron/Bobby', 'sendTestSeniorLeadershipRecap')
    .addItem('Test email batches to Aaron/Robert', 'sendTestManagerExceptionEmails')
    .addItem('Email managers current exceptions', 'sendManagerExceptionEmails')
    .addToUi();
}

function createSetupSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  createSheetWithHeaders_(spreadsheet, CSV_DUMP_SHEET_NAME, CSV_DUMP_HEADERS);
  getOrCreateSheet_(spreadsheet, MANAGER_ROSTER_SHEET_NAME);
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

  var csvRows = readSheetRows_(spreadsheet, CSV_DUMP_SHEET_NAME);
  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import the ReflexAI CSV first.');
    return;
  }

  var managerRoster = buildManagerRoster_(spreadsheet);
  var runSettings = getRunSettings_(spreadsheet);
  var exceptions = [];

  csvRows.forEach(function(row) {
    var userName = getValue_(row, 'User Name');
    var userEmail = getValue_(row, 'User Email');
    var simulationName = getValue_(row, 'Simulation Name');
    var status = getValue_(row, 'Status');
    var score = getValue_(row, 'Best Score (%)');

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

    var managerInfo =
      managerRoster[String(userEmail).toLowerCase().trim()] ||
      managerRoster[String(userName).toLowerCase().trim()] ||
      {};

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

  writeExceptions_(spreadsheet, exceptions);
  writeRunLog_(spreadsheet, csvRows.length, exceptions.length, false);

  SpreadsheetApp.getUi().alert(
    'Report complete.\n\nRows checked: ' +
    csvRows.length +
    '\nExceptions found: ' +
    exceptions.length
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
        rows: []
      };
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
      : managerEmail;

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

function sendSeniorLeadershipRecap_(recipients, testMode) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var csvRows = readSheetRows_(spreadsheet, CSV_DUMP_SHEET_NAME);

  if (!csvRows.length) {
    SpreadsheetApp.getUi().alert('No CSV data found. Import the ReflexAI CSV first.');
    return;
  }

  var managerRoster = buildManagerRoster_(spreadsheet);
  var recap = buildSeniorLeadershipRecap_(csvRows, managerRoster);
  var subject = (testMode ? '[TEST] ' : '') + getCurrentMonthName_() + ' ReflexAI Supergroup Recap';

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

function buildSeniorLeadershipRecap_(csvRows, managerRoster) {
  var recap = {
    generatedAt: new Date(),
    totalRows: csvRows.length,
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
    var managerInfo =
      managerRoster[String(userEmail).toLowerCase().trim()] ||
      managerRoster[String(userName).toLowerCase().trim()] ||
      {};
    var managerKey = String(managerInfo.managerEmail || managerInfo.managerName || 'Unassigned').toLowerCase().trim();

    if (!managerKey) managerKey = 'unassigned';

    recap.counts[category]++;

    if (!recap.byManager[managerKey]) {
      recap.byManager[managerKey] = {
        managerName: managerInfo.managerName || 'Unassigned',
        managerEmail: managerInfo.managerEmail || '',
        notStarted: 0,
        completedBelow: 0,
        completedAbove: 0,
        total: 0
      };
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

function buildSeniorLeadershipRecapText_(recap, testMode) {
  var lines = [];

  if (testMode) {
    lines.push('TEST MODE - Senior leadership recap preview.');
    lines.push('');
  }

  lines.push('ReflexAI senior leadership recap');
  lines.push('Rows included: ' + recap.totalRows);
  lines.push('');
  lines.push('Overall simulation counts');
  lines.push('Not Started / In Progress: ' + recap.counts.notStarted);
  lines.push('Completed Below Threshold: ' + recap.counts.completedBelow);
  lines.push('Completed At/Above Threshold: ' + recap.counts.completedAbove);
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
      ' | Completed At/Above: ' +
      formatPercent_(manager.completedAbove, manager.total) +
      ' | Completed Below: ' +
      formatPercent_(manager.completedBelow, manager.total) +
      ' | Not Started / In Progress: ' +
      formatPercent_(manager.notStarted, manager.total)
    );
  });

  return lines.join('\n');
}

function buildSeniorLeadershipRecapHtml_(recap, testMode) {
  return (testMode ? '<p><strong>TEST MODE</strong> - Senior leadership recap preview.</p>' : '') +
    '<h2>ReflexAI Senior Leadership Recap</h2>' +
    '<p><strong>Rows included:</strong> ' + recap.totalRows + '</p>' +
    buildLeadershipCountsTable_(recap) +
    buildSimulationAverageTable_(recap.simulationAverages) +
    buildManagerRecapTable_(recap.managerRows);
}

function buildLeadershipCountsTable_(recap) {
  return '<h3>Overall Simulation Counts</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Metric</th><th>Count</th></tr></thead>' +
    '<tbody>' +
    '<tr><td>Not Started / In Progress</td><td>' + recap.counts.notStarted + '</td></tr>' +
    '<tr><td>Completed Below Threshold</td><td>' + recap.counts.completedBelow + '</td></tr>' +
    '<tr><td>Completed At/Above Threshold</td><td>' + recap.counts.completedAbove + '</td></tr>' +
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
      '<td>' + escapeHtml_(manager.managerEmail) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.completedAbove, manager.total)) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.completedBelow, manager.total)) + '</td>' +
      '<td>' + escapeHtml_(formatPercent_(manager.notStarted, manager.total)) + '</td>' +
      '<td>' + manager.total + '</td>' +
      '</tr>';
  }).join('');

  return '<h3>Manager Breakdown</h3>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">' +
    '<thead><tr><th>Manager</th><th>Manager Email</th><th>% Completed At/Above Threshold</th><th>% Completed Below Threshold</th><th>% Not Started / In Progress</th><th>Total Simulations</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table>';
}

function buildManagerRoster_(spreadsheet) {
  var rows = readSheetRows_(spreadsheet, MANAGER_ROSTER_SHEET_NAME);
  var roster = {};
  var nameToEmail = {};

  rows.forEach(function(row) {
    var name = String(getValue_(row, 'Name')).toLowerCase().trim();
    var email = String(getValue_(row, 'Email')).toLowerCase().trim();

    if (name && email) {
      nameToEmail[name] = email;
    }
  });

  rows.forEach(function(row) {
    var repEmail = String(getValue_(row, 'Email')).toLowerCase().trim();
    var repName = String(getValue_(row, 'Name')).toLowerCase().trim();

    var managerName = getValue_(row, 'Manager');
    var seniorName = getValue_(row, 'Senior');

    var managerEmail = nameToEmail[String(managerName).toLowerCase().trim()] || '';
    var seniorEmail = nameToEmail[String(seniorName).toLowerCase().trim()] || '';

    var managerInfo = {
      managerName: managerName,
      managerEmail: managerEmail,
      seniorName: seniorName,
      seniorEmail: seniorEmail
    };

    if (repEmail) {
      roster[repEmail] = managerInfo;
    }

    if (repName) {
      roster[repName] = managerInfo;
    }
  });

  return roster;
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
  lines.push('Only not-started, in-progress, or below-80% simulations are included.');
  lines.push('');

  appendPlainTextSection_(lines, 'BELOW 80% COMPLETED SIMULATIONS - PRIORITY', sections.lowScoreGroups);
  appendPlainTextSection_(lines, 'NOT STARTED / IN PROGRESS SIMULATIONS', sections.incompleteGroups);

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
    '<p>Below are ReflexAI simulation follow-up items for your team. Only not-started, in-progress, or below-80% simulations are included.</p>' +
    buildHtmlSection_('Below 80% Completed Simulations - Priority', sections.lowScoreGroups, true) +
    buildHtmlSection_('Not Started / In Progress Simulations', sections.incompleteGroups, false) +
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
