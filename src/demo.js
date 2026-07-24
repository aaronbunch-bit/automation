/**
 * Demo fixtures so the bank can run without Looker / Sheets credentials.
 * Simulates three coaching sessions with known L7→N7 outcomes.
 */

import { addDays, formatUtcDate, toUtcDate } from './score.js';

function dailySeries(personKey, startDate, days, base, slope) {
  const out = [];
  const start = toUtcDate(startDate);
  for (let i = 0; i < days; i++) {
    out.push({
      personKey,
      date: formatUtcDate(addDays(start, i)),
      pgc: Number((base + slope * i).toFixed(4)),
    });
  }
  return out;
}

/** Fixed "today" for deterministic demo scoring. */
export const DEMO_AS_OF = '2026-07-24';

export function loadDemoCoachingEvents() {
  return [
    {
      id: 'alex@example.com|2026-07-10|2',
      coachee: 'Alex Rivera',
      coacheeEmail: 'alex@example.com',
      coach: 'Jordan Lee',
      coachingDate: '2026-07-10',
      sheetRow: 2,
    },
    {
      id: 'sam@example.com|2026-07-08|3',
      coachee: 'Sam Patel',
      coacheeEmail: 'sam@example.com',
      coach: 'Jordan Lee',
      coachingDate: '2026-07-08',
      sheetRow: 3,
    },
    {
      id: 'casey@example.com|2026-07-20|4',
      coachee: 'Casey Ng',
      coacheeEmail: 'casey@example.com',
      coach: 'Morgan Diaz',
      coachingDate: '2026-07-20',
      sheetRow: 4,
    },
  ];
}

/**
 * Alex: clear lift after coaching → effective
 * Sam: flat/down after coaching → poor
 * Casey: N7 not yet complete → pending
 */
export function loadDemoPgcSeries() {
  return [
    // Alex: L7 ~0.40, N7 climbs to ~0.55
    ...dailySeries('alex@example.com', '2026-07-01', 7, 0.4, 0),
    ...dailySeries('alex@example.com', '2026-07-11', 7, 0.52, 0.01),
    // Sam: L7 ~0.55, N7 ~0.48
    ...dailySeries('sam@example.com', '2026-07-01', 7, 0.55, 0),
    ...dailySeries('sam@example.com', '2026-07-09', 7, 0.5, -0.005),
    // Casey: only pre-window so far
    ...dailySeries('casey@example.com', '2026-07-13', 7, 0.45, 0.002),
  ];
}
