/**
 * Orchestrates one full sync of the pGC coaching bank.
 */

import { assertLiveCredentials, loadConfig } from './config.js';
import { loadBank, saveBank } from './bank.js';
import { loadDemoCoachingEvents, loadDemoPgcSeries, DEMO_AS_OF } from './demo.js';
import { fetchPgcSeries, indexSeriesByPerson } from './looker.js';
import {
  addDays,
  formatUtcDate,
  scoreCoachingEvent,
  summarizeByCoach,
  summarizeBySupergroup,
  toUtcDate,
} from './score.js';
import { createSheetsClient, readCoachingEvents, writeScoresToSheet } from './sheets.js';

function personKeyForEvent(event) {
  return (event.coacheeEmail || event.coachee || '').trim();
}

function dateBounds(events) {
  if (!events.length) {
    const today = formatUtcDate(new Date());
    return { startDate: today, endDate: today };
  }
  let min = events[0].coachingDate;
  let max = events[0].coachingDate;
  for (const e of events) {
    if (e.coachingDate < min) min = e.coachingDate;
    if (e.coachingDate > max) max = e.coachingDate;
  }
  return {
    startDate: formatUtcDate(addDays(toUtcDate(min), -7)),
    endDate: formatUtcDate(addDays(toUtcDate(max), 7)),
  };
}

export async function runSync(options = {}) {
  const config = options.config || loadConfig();
  const asOf = options.asOf || (config.mode === 'demo' ? DEMO_AS_OF : new Date());
  const writeback = options.writeback ?? config.sheet.writeback;

  let events;
  let series;
  let writebackResult = { wrote: 0, headersAdded: [], skipped: true };

  if (config.mode === 'demo') {
    events = loadDemoCoachingEvents();
    series = loadDemoPgcSeries();
  } else {
    assertLiveCredentials(config);
    const sheets = createSheetsClient(config.sheet);
    const sheetData = await readCoachingEvents(sheets, config.sheet);
    events = sheetData.events;

    const keys = [
      ...new Set(events.map(personKeyForEvent).filter(Boolean)),
    ];
    const bounds = dateBounds(events);
    series = await fetchPgcSeries(config.looker, {
      personKeys: keys,
      startDate: bounds.startDate,
      endDate: bounds.endDate,
    });

    if (writeback) {
      // score first, then write — computed below; placeholder filled after scoring
      writebackResult.skipped = false;
    }
  }

  const byPerson = indexSeriesByPerson(series);
  const scored = events.map((event) => {
    const key = personKeyForEvent(event).toLowerCase();
    const personSeries = byPerson.get(key) || [];
    return scoreCoachingEvent(event, personSeries, asOf);
  });

  if (config.mode !== 'demo' && writeback) {
    const sheets = createSheetsClient(config.sheet);
    writebackResult = await writeScoresToSheet(sheets, config.sheet, scored);
  }

  const coaches = summarizeByCoach(scored);
  const supergroups = summarizeBySupergroup(scored);
  const bank = {
    lastSync: new Date().toISOString(),
    asOf: formatUtcDate(asOf),
    mode: config.mode,
    sessions: scored,
    coaches,
    supergroups,
    meta: {
      eventCount: events.length,
      seriesPoints: series.length,
      writeback: writebackResult,
      spreadsheetId: config.sheet.spreadsheetId,
      tab: config.sheet.tab,
    },
  };

  const saved = await saveBank(config.blobStore, bank);
  return { bank: saved.payload, backend: saved.backend };
}
