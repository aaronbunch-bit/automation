/**
 * Google Sheets reader/writer for coaching events and score writeback.
 */

import { google } from 'googleapis';
import { formatUtcDate, toUtcDate } from './score.js';

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function findColumnIndex(headers, wanted) {
  const target = normalizeHeader(wanted);
  return headers.findIndex((h) => normalizeHeader(h) === target);
}

function ensureColumn(headers, wanted) {
  let idx = findColumnIndex(headers, wanted);
  if (idx >= 0) return { headers, index: idx, added: false };
  const next = [...headers, wanted];
  return { headers: next, index: next.length - 1, added: true };
}

export function createSheetsClient(sheetConfig) {
  const auth = new google.auth.JWT({
    email: sheetConfig.serviceAccountEmail,
    key: sheetConfig.serviceAccountPrivateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

function parseCellDate(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    // Sheets serial date
    const utc = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    return formatUtcDate(utc);
  }
  return formatUtcDate(toUtcDate(value));
}

/**
 * Read coaching events from the configured tab.
 * @returns {Promise<{ events: object[], headers: string[], sheetIdHint: string }>}
 */
export async function readCoachingEvents(sheets, sheetConfig) {
  const range = `${sheetConfig.tab}!A:Z`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetConfig.spreadsheetId,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });

  const values = res.data.values || [];
  if (!values.length) return { events: [], headers: [] };

  let headers = values[0].map((h) => String(h ?? ''));
  const cols = sheetConfig.columns;
  const coacheeIdx = findColumnIndex(headers, cols.coachee);
  const emailIdx = findColumnIndex(headers, cols.coacheeEmail);
  const coachIdx = findColumnIndex(headers, cols.coach);
  const dateIdx = findColumnIndex(headers, cols.coachingDate);

  if (dateIdx < 0) {
    throw new Error(
      `Coaching sheet tab "${sheetConfig.tab}" is missing column "${cols.coachingDate}". ` +
        `Found headers: ${headers.join(', ') || '(none)'}`
    );
  }
  if (coacheeIdx < 0 && emailIdx < 0) {
    throw new Error(
      `Coaching sheet needs "${cols.coachee}" or "${cols.coacheeEmail}". ` +
        `Found headers: ${headers.join(', ')}`
    );
  }

  const events = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const coachingDate = parseCellDate(row[dateIdx]);
    if (!coachingDate) continue;
    const coachee = coacheeIdx >= 0 ? String(row[coacheeIdx] ?? '').trim() : '';
    const coacheeEmail = emailIdx >= 0 ? String(row[emailIdx] ?? '').trim() : '';
    const coach = coachIdx >= 0 ? String(row[coachIdx] ?? '').trim() : '';
    if (!coachee && !coacheeEmail) continue;
    events.push({
      id: `${coacheeEmail || coachee}|${coachingDate}|${i + 1}`,
      coachee,
      coacheeEmail,
      coach,
      coachingDate,
      sheetRow: i + 1,
    });
  }

  return { events, headers };
}

/**
 * Ensure writeback columns exist and patch scored values onto matching rows.
 */
export async function writeScoresToSheet(sheets, sheetConfig, scored) {
  if (!sheetConfig.writeback) return { wrote: 0, headersAdded: [] };

  const range = `${sheetConfig.tab}!A1:ZZ`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetConfig.spreadsheetId,
    range,
  });
  const values = res.data.values || [];
  if (!values.length) return { wrote: 0, headersAdded: [] };

  let headers = values[0].map((h) => String(h ?? ''));
  const cols = sheetConfig.columns;
  const headersAdded = [];

  for (const key of ['l7', 'n7', 'delta', 'verdict', 'scoredAt']) {
    const ensured = ensureColumn(headers, cols[key]);
    headers = ensured.headers;
    if (ensured.added) headersAdded.push(cols[key]);
  }

  const colIndex = {
    l7: findColumnIndex(headers, cols.l7),
    n7: findColumnIndex(headers, cols.n7),
    delta: findColumnIndex(headers, cols.delta),
    verdict: findColumnIndex(headers, cols.verdict),
    scoredAt: findColumnIndex(headers, cols.scoredAt),
  };

  const width = headers.length;
  const grid = values.map((row) => {
    const next = [...row];
    while (next.length < width) next.push('');
    return next;
  });
  grid[0] = headers;

  const byRow = new Map(scored.filter((s) => s.sheetRow).map((s) => [s.sheetRow, s]));
  let wrote = 0;
  for (const [rowNum, score] of byRow) {
    if (score.verdict === 'pending') continue;
    const r = rowNum - 1;
    if (r < 1 || r >= grid.length) continue;
    while (grid[r].length < width) grid[r].push('');
    grid[r][colIndex.l7] = score.l7 ?? '';
    grid[r][colIndex.n7] = score.n7 ?? '';
    grid[r][colIndex.delta] = score.delta ?? '';
    grid[r][colIndex.verdict] = score.verdict ?? '';
    grid[r][colIndex.scoredAt] = score.scoredAt ?? '';
    wrote += 1;
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetConfig.spreadsheetId,
    range: `${sheetConfig.tab}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: grid },
  });

  return { wrote, headersAdded };
}
