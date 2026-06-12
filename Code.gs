/**
 * ReflexAI weekly exception reporting for Google Sheets.
 *
 * Configure Script Properties before running:
 * - SPREADSHEET_ID
 * - REFLEXAI_BASE_URL
 * - REFLEXAI_REPORT_PATH
 * - One authentication option:
 *   - REFLEXAI_BEARER_TOKEN
 *   - REFLEXAI_API_KEY plus optional REFLEXAI_API_KEY_HEADER
 *   - REFLEXAI_USERNAME, REFLEXAI_PASSWORD, REFLEXAI_LOGIN_PATH, and REFLEXAI_LOGIN_TOKEN_JSON_PATH
 */

var CONFIG_KEYS = {
  spreadsheetId: 'SPREADSHEET_ID',
  reflexBaseUrl: 'REFLEXAI_BASE_URL',
  reportPath: 'REFLEXAI_REPORT_PATH',
  reportRowsJsonPath: 'REFLEXAI_REPORT_ROWS_JSON_PATH',
  bearerToken: 'REFLEXAI_BEARER_TOKEN',
  apiKey: 'REFLEXAI_API_KEY',
  apiKeyHeader: 'REFLEXAI_API_KEY_HEADER',
  username: 'REFLEXAI_USERNAME',
  password: 'REFLEXAI_PASSWORD',
  loginPath: 'REFLEXAI_LOGIN_PATH',
  loginTokenJsonPath: 'REFLEXAI_LOGIN_TOKEN_JSON_PATH',
  fieldMapJson: 'FIELD_MAP_JSON',
  monthlyJourneyIds: 'MONTHLY_JOURNEY_IDS',
  monthlyJourneyNamePattern: 'MONTHLY_JOURNEY_NAME_PATTERN',
  passingScorePercent: 'PASSING_SCORE_PERCENT',
  senderName: 'EMAIL_SENDER_NAME',
  emailSubjectPrefix: 'EMAIL_SUBJECT_PREFIX',
  dryRun: 'DRY_RUN',
  sendEmails: 'SEND_EMAILS'
};

var DEFAULT_FIELD_MAP = {
  repName: ['repName', 'representativeName', 'learnerName', 'agentName', 'userName', 'name'],
  repEmail: ['repEmail', 'representativeEmail', 'learnerEmail', 'agentEmail', 'userEmail', 'email'],
  managerName: ['managerName', 'supervisorName'],
  managerEmail: ['managerEmail', 'supervisorEmail'],
  teamLeadName: ['teamLeadName', 'team_lead_name', 'leadName'],
  teamLeadEmail: ['teamLeadEmail', 'team_lead_email', 'leadEmail'],
  journeyId: ['journeyId', 'journey_id', 'courseId'],
  journeyName: ['journeyName', 'journey', 'courseName', 'assignmentName'],
  simulationId: ['simulationId', 'simulation_id', 'scenarioId'],
  simulationName: ['simulationName', 'simulation', 'scenarioName', 'moduleName'],
  status: ['status', 'completionStatus', 'attemptStatus', 'state'],
  score: ['score', 'scorePercent', 'overallScore', 'overallJourneyScore', 'percentageScore'],
  completedAt: ['completedAt', 'completionDate', 'completed_at', 'lastAttemptAt'],
  attemptedAt: ['attemptedAt', 'startedAt', 'lastStartedAt']
};

var EXCEPTION_SHEET_NAME = 'Simulation Exceptions';
var RUN_LOG_SHEET_NAME = 'Run Log';
var CONFIG_SHEET_NAME = 'Setup Checklist';

var EXCEPTION_HEADERS = [
  'Report Run At',
  'Representative',
  'Representative Email',
  'Manager',
  'Manager Email',
  'Team Lead',
  'Team Lead Email',
  'Journey',
  'Simulation',
  'Completion Status',
  'Score',
  'Required Follow-up Action'
];

/**
 * Adds a small menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ReflexAI Reporting')
    .addItem('Run weekly exception report now', 'runWeeklySimulationExceptionReport')
    .addItem('Validate configuration', 'validateConfiguration')
    .addItem('Install weekly trigger', 'installWeeklyTrigger')
    .addItem('Create setup sheets', 'createSetupSheets')
    .addToUi();
}

/**
 * Main scheduled function. Pulls ReflexAI simulation data, writes exception rows,
 * and emails Managers/Team Leads when there is action to take.
 */
function runWeeklySimulationExceptionReport() {
  var config = getConfig_();
  validateConfigOrThrow_(config);

  var reportRunAt = new Date();
  var records = fetchReflexAiReportRows_(config);
  var exceptions = buildExceptionRows_(records, config, reportRunAt);

  writeExceptionRows_(config.spreadsheetId, exceptions);
  writeRunLog_(config.spreadsheetId, reportRunAt, records.length, exceptions.length);

  if (config.sendEmails) {
    sendExceptionEmails_(exceptions, config, reportRunAt);
  }

  return {
    recordsProcessed: records.length,
    exceptionsFound: exceptions.length,
    emailsEnabled: config.sendEmails,
    dryRun: config.dryRun
  };
}

/**
 * Checks required configuration and verifies ReflexAI connectivity.
 */
function validateConfiguration() {
  var config = getConfig_();
  validateConfigOrThrow_(config);

  var records = fetchReflexAiReportRows_(config);
  if (!records.length) {
    throw new Error('ReflexAI request succeeded, but no report rows were returned. Confirm REFLEXAI_REPORT_ROWS_JSON_PATH and the report filters.');
  }

  var normalized = normalizeRecord_(records[0], config.fieldMap);
  var missing = ['repName', 'journeyName', 'simulationName', 'status'].filter(function(field) {
    return !normalized[field];
  });

  if (missing.length) {
    throw new Error('ReflexAI returned data, but these fields could not be mapped from the first row: ' + missing.join(', ') + '. Update FIELD_MAP_JSON.');
  }

  createSetupSheets();
  Browser.msgBox('Configuration is valid. First ReflexAI row was read and mapped successfully.');
}

/**
 * Installs a weekly trigger for Monday morning.
 */
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runWeeklySimulationExceptionReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runWeeklySimulationExceptionReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
}

/**
 * Creates starter tabs for output and configuration guidance.
 */
function createSetupSheets() {
  var config = getConfig_({ skipValidation: true });
  if (!config.spreadsheetId) {
    throw new Error('Set SPREADSHEET_ID in Script Properties before creating sheets.');
  }

  var spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  var exceptionSheet = getOrCreateSheet_(spreadsheet, EXCEPTION_SHEET_NAME);
  exceptionSheet.clear();
  exceptionSheet.getRange(1, 1, 1, EXCEPTION_HEADERS.length).setValues([EXCEPTION_HEADERS]);
  exceptionSheet.setFrozenRows(1);

  var runLogSheet = getOrCreateSheet_(spreadsheet, RUN_LOG_SHEET_NAME);
  runLogSheet.clear();
  runLogSheet.getRange(1, 1, 1, 5).setValues([[
    'Report Run At',
    'Records Processed',
    'Exceptions Found',
    'Emails Enabled',
    'Dry Run'
  ]]);
  runLogSheet.setFrozenRows(1);

  var checklistSheet = getOrCreateSheet_(spreadsheet, CONFIG_SHEET_NAME);
  checklistSheet.clear();
  checklistSheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Required?', 'Notes']]);
  checklistSheet.getRange(2, 1, 17, 3).setValues([
    ['SPREADSHEET_ID', 'Yes', 'Google Sheet ID that receives exception rows.'],
    ['REFLEXAI_BASE_URL', 'Yes', 'Base ReflexAI tenant URL, for example https://your-tenant.example.com.'],
    ['REFLEXAI_REPORT_PATH', 'Yes', 'Path for the ReflexAI report/export endpoint that returns simulation attempts.'],
    ['REFLEXAI_REPORT_ROWS_JSON_PATH', 'If JSON response is nested', 'Dot path to the array of records, for example data.rows.'],
    ['REFLEXAI_BEARER_TOKEN', 'One auth option', 'Use if ReflexAI provides a bearer/API token.'],
    ['REFLEXAI_API_KEY', 'One auth option', 'Use with REFLEXAI_API_KEY_HEADER if ReflexAI uses an API-key header.'],
    ['REFLEXAI_API_KEY_HEADER', 'No', 'Defaults to x-api-key.'],
    ['REFLEXAI_USERNAME', 'One auth option', 'Use for login-token flow only.'],
    ['REFLEXAI_PASSWORD', 'One auth option', 'Use for login-token flow only.'],
    ['REFLEXAI_LOGIN_PATH', 'Login-token auth only', 'Path that exchanges username/password for a token.'],
    ['REFLEXAI_LOGIN_TOKEN_JSON_PATH', 'Login-token auth only', 'Dot path to token in login response, for example access_token.'],
    ['FIELD_MAP_JSON', 'Only if ReflexAI names differ', 'JSON object mapping normalized names to ReflexAI field names.'],
    ['MONTHLY_JOURNEY_IDS', 'No', 'Comma-separated journey IDs to include.'],
    ['MONTHLY_JOURNEY_NAME_PATTERN', 'No', 'Regex used to include monthly workshop journeys by name.'],
    ['PASSING_SCORE_PERCENT', 'No', 'Defaults to 80.'],
    ['SEND_EMAILS', 'No', 'Set true to email managers/team leads. Defaults false until tested.'],
    ['DRY_RUN', 'No', 'Set true to write the sheet without sending emails.']
  ]);
  checklistSheet.setFrozenRows(1);
}

function getConfig_(options) {
  var props = PropertiesService.getScriptProperties();
  var rawFieldMap = props.getProperty(CONFIG_KEYS.fieldMapJson);
  var fieldMap = mergeFieldMap_(DEFAULT_FIELD_MAP, rawFieldMap ? JSON.parse(rawFieldMap) : {});

  var passingScore = Number(props.getProperty(CONFIG_KEYS.passingScorePercent) || '80');
  var monthlyJourneyIds = splitCsv_(props.getProperty(CONFIG_KEYS.monthlyJourneyIds));
  var monthlyJourneyNamePattern = props.getProperty(CONFIG_KEYS.monthlyJourneyNamePattern);

  return {
    spreadsheetId: props.getProperty(CONFIG_KEYS.spreadsheetId),
    reflexBaseUrl: trimTrailingSlash_(props.getProperty(CONFIG_KEYS.reflexBaseUrl) || ''),
    reportPath: props.getProperty(CONFIG_KEYS.reportPath),
    reportRowsJsonPath: props.getProperty(CONFIG_KEYS.reportRowsJsonPath),
    bearerToken: props.getProperty(CONFIG_KEYS.bearerToken),
    apiKey: props.getProperty(CONFIG_KEYS.apiKey),
    apiKeyHeader: props.getProperty(CONFIG_KEYS.apiKeyHeader) || 'x-api-key',
    username: props.getProperty(CONFIG_KEYS.username),
    password: props.getProperty(CONFIG_KEYS.password),
    loginPath: props.getProperty(CONFIG_KEYS.loginPath),
    loginTokenJsonPath: props.getProperty(CONFIG_KEYS.loginTokenJsonPath) || 'access_token',
    fieldMap: fieldMap,
    monthlyJourneyIds: monthlyJourneyIds,
    monthlyJourneyNameRegex: monthlyJourneyNamePattern ? new RegExp(monthlyJourneyNamePattern, 'i') : null,
    passingScorePercent: passingScore,
    senderName: props.getProperty(CONFIG_KEYS.senderName) || 'ReflexAI Simulation Reporting',
    emailSubjectPrefix: props.getProperty(CONFIG_KEYS.emailSubjectPrefix) || 'ReflexAI weekly simulation follow-up',
    dryRun: toBoolean_(props.getProperty(CONFIG_KEYS.dryRun), false),
    sendEmails: toBoolean_(props.getProperty(CONFIG_KEYS.sendEmails), false),
    skipValidation: options && options.skipValidation
  };
}

function validateConfigOrThrow_(config) {
  var missing = [];
  if (!config.spreadsheetId) missing.push(CONFIG_KEYS.spreadsheetId);
  if (!config.reflexBaseUrl) missing.push(CONFIG_KEYS.reflexBaseUrl);
  if (!config.reportPath) missing.push(CONFIG_KEYS.reportPath);
  if (!isFinite(config.passingScorePercent)) missing.push(CONFIG_KEYS.passingScorePercent);

  var hasBearer = Boolean(config.bearerToken);
  var hasApiKey = Boolean(config.apiKey);
  var hasLogin = Boolean(config.username && config.password && config.loginPath);
  if (!hasBearer && !hasApiKey && !hasLogin) {
    missing.push('one auth method: REFLEXAI_BEARER_TOKEN, REFLEXAI_API_KEY, or username/password/login settings');
  }

  if (missing.length) {
    throw new Error('Missing required Script Properties: ' + missing.join(', '));
  }
}

function fetchReflexAiReportRows_(config) {
  var response = reflexFetch_(config, config.reportPath, { method: 'get' });
  var contentText = response.getContentText();
  var contentType = String(response.getHeaders()['Content-Type'] || response.getHeaders()['content-type'] || '');

  if (contentType.indexOf('text/csv') !== -1 || looksLikeCsv_(contentText)) {
    return parseCsvRecords_(contentText);
  }

  var parsed = JSON.parse(contentText);
  var rows = config.reportRowsJsonPath ? getByPath_(parsed, config.reportRowsJsonPath) : parsed;
  if (rows && rows.data && Array.isArray(rows.data)) rows = rows.data;
  if (rows && rows.rows && Array.isArray(rows.rows)) rows = rows.rows;
  if (!Array.isArray(rows)) {
    throw new Error('ReflexAI report response did not contain an array of rows. Set REFLEXAI_REPORT_ROWS_JSON_PATH to the array location.');
  }
  return rows;
}

function reflexFetch_(config, path, options) {
  var requestOptions = options || {};
  var headers = requestOptions.headers || {};
  headers.Accept = headers.Accept || 'application/json';

  if (config.bearerToken) {
    headers.Authorization = 'Bearer ' + config.bearerToken;
  } else if (config.apiKey) {
    headers[config.apiKeyHeader] = config.apiKey;
  } else {
    headers.Authorization = 'Bearer ' + loginAndGetToken_(config);
  }

  var fetchOptions = {
    method: requestOptions.method || 'get',
    headers: headers,
    muteHttpExceptions: true
  };

  if (requestOptions.payload) {
    fetchOptions.contentType = requestOptions.contentType || 'application/json';
    fetchOptions.payload = typeof requestOptions.payload === 'string'
      ? requestOptions.payload
      : JSON.stringify(requestOptions.payload);
  }

  var response = UrlFetchApp.fetch(buildUrl_(config.reflexBaseUrl, path), fetchOptions);
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('ReflexAI request failed with HTTP ' + code + ': ' + response.getContentText().slice(0, 500));
  }

  return response;
}

function loginAndGetToken_(config) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('REFLEXAI_LOGIN_TOKEN');
  if (cached) return cached;

  var response = UrlFetchApp.fetch(buildUrl_(config.reflexBaseUrl, config.loginPath), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      username: config.username,
      email: config.username,
      password: config.password
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('ReflexAI login failed with HTTP ' + code + ': ' + response.getContentText().slice(0, 500));
  }

  var body = JSON.parse(response.getContentText());
  var token = getByPath_(body, config.loginTokenJsonPath);
  if (!token) {
    throw new Error('ReflexAI login succeeded, but no token was found at ' + config.loginTokenJsonPath + '.');
  }

  cache.put('REFLEXAI_LOGIN_TOKEN', token, 3300);
  return token;
}

function buildExceptionRows_(records, config, reportRunAt) {
  return records
    .map(function(record) {
      return normalizeRecord_(record, config.fieldMap);
    })
    .filter(function(record) {
      return isMonthlyJourney_(record, config);
    })
    .map(function(record) {
      return toExceptionRow_(record, config, reportRunAt);
    })
    .filter(function(row) {
      return row !== null;
    });
}

function normalizeRecord_(record, fieldMap) {
  var normalized = {};
  Object.keys(fieldMap).forEach(function(targetField) {
    normalized[targetField] = firstValue_(record, fieldMap[targetField]);
  });
  return normalized;
}

function toExceptionRow_(record, config, reportRunAt) {
  var scorePercent = parseScorePercent_(record.score);
  var status = String(record.status || '').trim();
  var statusLower = status.toLowerCase();
  var hasAttempt = Boolean(record.attemptedAt || record.completedAt || status || isFinite(scorePercent));
  var completed = ['completed', 'complete', 'passed', 'failed', 'scored'].indexOf(statusLower) !== -1 || Boolean(record.completedAt);

  var action = null;
  if (!hasAttempt || ['not started', 'not attempted', 'not_started', 'assigned', 'pending'].indexOf(statusLower) !== -1) {
    status = status || 'Not attempted';
    action = 'Ask representative to complete this simulation.';
  } else if (completed && isFinite(scorePercent) && scorePercent < config.passingScorePercent) {
    action = 'Ask representative to retake this simulation and coach on missed skills.';
  } else if (completed && !isFinite(scorePercent)) {
    action = 'Review manually: completed simulation is missing a score.';
  }

  if (!action) return null;

  return {
    reportRunAt: reportRunAt,
    repName: record.repName || '',
    repEmail: record.repEmail || '',
    managerName: record.managerName || '',
    managerEmail: record.managerEmail || '',
    teamLeadName: record.teamLeadName || '',
    teamLeadEmail: record.teamLeadEmail || '',
    journeyName: record.journeyName || record.journeyId || '',
    simulationName: record.simulationName || record.simulationId || '',
    status: status,
    score: isFinite(scorePercent) ? scorePercent / 100 : '',
    action: action
  };
}

function isMonthlyJourney_(record, config) {
  if (config.monthlyJourneyIds.length && config.monthlyJourneyIds.indexOf(String(record.journeyId)) === -1) {
    return false;
  }

  if (config.monthlyJourneyNameRegex && !config.monthlyJourneyNameRegex.test(String(record.journeyName || ''))) {
    return false;
  }

  return true;
}

function writeExceptionRows_(spreadsheetId, exceptions) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = getOrCreateSheet_(spreadsheet, EXCEPTION_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, EXCEPTION_HEADERS.length).setValues([EXCEPTION_HEADERS]);
  sheet.setFrozenRows(1);

  if (!exceptions.length) return;

  var values = exceptions.map(function(row) {
    return [
      row.reportRunAt,
      row.repName,
      row.repEmail,
      row.managerName,
      row.managerEmail,
      row.teamLeadName,
      row.teamLeadEmail,
      row.journeyName,
      row.simulationName,
      row.status,
      row.score,
      row.action
    ];
  });

  sheet.getRange(2, 1, values.length, EXCEPTION_HEADERS.length).setValues(values);
  sheet.getRange(2, 11, values.length, 1).setNumberFormat('0%');
  sheet.autoResizeColumns(1, EXCEPTION_HEADERS.length);
}

function writeRunLog_(spreadsheetId, reportRunAt, recordsProcessed, exceptionsFound) {
  var config = getConfig_({ skipValidation: true });
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = getOrCreateSheet_(spreadsheet, RUN_LOG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([[
      'Report Run At',
      'Records Processed',
      'Exceptions Found',
      'Emails Enabled',
      'Dry Run'
    ]]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([reportRunAt, recordsProcessed, exceptionsFound, config.sendEmails, config.dryRun]);
}

function sendExceptionEmails_(exceptions, config, reportRunAt) {
  var grouped = groupExceptionsByRecipient_(exceptions);
  Object.keys(grouped).forEach(function(email) {
    var rows = grouped[email];
    var subject = config.emailSubjectPrefix + ' - ' + Utilities.formatDate(reportRunAt, Session.getScriptTimeZone(), 'MMM d, yyyy');
    var htmlBody = buildEmailHtml_(rows, reportRunAt);
    var plainBody = buildEmailText_(rows, reportRunAt);

    if (config.dryRun) {
      Logger.log('DRY_RUN email to %s: %s\n%s', email, subject, plainBody);
      return;
    }

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: config.senderName
    });
  });
}

function groupExceptionsByRecipient_(exceptions) {
  var grouped = {};
  exceptions.forEach(function(row) {
    [row.managerEmail, row.teamLeadEmail].forEach(function(email) {
      if (!email) return;
      var normalizedEmail = String(email).trim().toLowerCase();
      if (!normalizedEmail) return;
      if (!grouped[normalizedEmail]) grouped[normalizedEmail] = [];
      grouped[normalizedEmail].push(row);
    });
  });
  return grouped;
}

function buildEmailHtml_(rows, reportRunAt) {
  var tableRows = rows.map(function(row) {
    return '<tr>' +
      '<td>' + escapeHtml_(row.repName) + '</td>' +
      '<td>' + escapeHtml_(row.journeyName) + '</td>' +
      '<td>' + escapeHtml_(row.simulationName) + '</td>' +
      '<td>' + escapeHtml_(row.status) + '</td>' +
      '<td>' + escapeHtml_(formatScore_(row.score)) + '</td>' +
      '<td>' + escapeHtml_(row.action) + '</td>' +
      '</tr>';
  }).join('');

  return '<p>Below are ReflexAI monthly workshop simulation exceptions as of ' +
    escapeHtml_(Utilities.formatDate(reportRunAt, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')) +
    '.</p>' +
    '<p>Only simulations that are not attempted or scored below the passing threshold are included.</p>' +
    '<table border="1" cellpadding="6" cellspacing="0">' +
    '<thead><tr><th>Representative</th><th>Journey</th><th>Simulation</th><th>Status</th><th>Score</th><th>Follow-up Action</th></tr></thead>' +
    '<tbody>' + tableRows + '</tbody>' +
    '</table>';
}

function buildEmailText_(rows, reportRunAt) {
  var lines = [
    'ReflexAI monthly workshop simulation exceptions as of ' +
      Utilities.formatDate(reportRunAt, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a') + '.',
    '',
    'Only simulations that are not attempted or scored below the passing threshold are included.',
    ''
  ];

  rows.forEach(function(row) {
    lines.push([
      row.repName,
      row.journeyName,
      row.simulationName,
      row.status,
      formatScore_(row.score),
      row.action
    ].join(' | '));
  });

  return lines.join('\n');
}

function parseCsvRecords_(contentText) {
  var rows = Utilities.parseCsv(contentText);
  if (!rows.length) return [];
  var headers = rows.shift().map(function(header) {
    return String(header).trim();
  });

  return rows
    .filter(function(row) {
      return row.some(function(value) { return value !== ''; });
    })
    .map(function(row) {
      var record = {};
      headers.forEach(function(header, index) {
        record[header] = row[index];
      });
      return record;
    });
}

function firstValue_(record, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var value = getByPath_(record, candidates[i]);
    if (value !== undefined && value !== null && value !== '') return value;
  }
  return '';
}

function getByPath_(object, path) {
  if (!path) return undefined;
  return String(path).split('.').reduce(function(current, key) {
    if (current === undefined || current === null) return undefined;
    return current[key];
  }, object);
}

function parseScorePercent_(value) {
  if (value === undefined || value === null || value === '') return NaN;
  if (typeof value === 'number') return value <= 1 ? value * 100 : value;

  var normalized = String(value).trim().replace('%', '');
  var parsed = Number(normalized);
  if (!isFinite(parsed)) return NaN;
  return parsed <= 1 ? parsed * 100 : parsed;
}

function formatScore_(score) {
  return typeof score === 'number' && isFinite(score)
    ? Math.round(score * 100) + '%'
    : 'No score';
}

function mergeFieldMap_(defaults, custom) {
  var merged = {};
  Object.keys(defaults).forEach(function(key) {
    merged[key] = defaults[key].slice();
  });

  Object.keys(custom || {}).forEach(function(key) {
    merged[key] = Array.isArray(custom[key]) ? custom[key] : [custom[key]];
  });

  return merged;
}

function splitCsv_(value) {
  if (!value) return [];
  return String(value)
    .split(',')
    .map(function(item) { return item.trim(); })
    .filter(Boolean);
}

function toBoolean_(value, defaultValue) {
  if (value === undefined || value === null || value === '') return defaultValue;
  return ['true', '1', 'yes', 'y'].indexOf(String(value).toLowerCase()) !== -1;
}

function buildUrl_(baseUrl, path) {
  if (/^https?:\/\//i.test(path)) return path;
  return trimTrailingSlash_(baseUrl) + '/' + String(path || '').replace(/^\/+/, '');
}

function trimTrailingSlash_(value) {
  return String(value || '').replace(/\/+$/, '');
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function looksLikeCsv_(contentText) {
  var firstLine = String(contentText || '').split(/\r?\n/)[0] || '';
  return firstLine.indexOf(',') !== -1 && firstLine.indexOf('{') === -1 && firstLine.indexOf('[') === -1;
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
