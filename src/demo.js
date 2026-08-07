/**
 * Demo fixtures so the bank can run without Looker / Sheets credentials.
 * Simulates the Consumer Sales org: six supergroups, each with a couple of
 * reps whose L7→N7 pGC outcomes are deterministic.
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

/**
 * Consumer Sales rep catalog grouped by supergroup.
 * `l7`/`n7` are the flat pGC means for the before/after windows.
 * `pending: true` uses a recent coaching date so N7 has not completed yet.
 */
const DEMO_REPS = [
  // Adult Learning — both lift → effective
  { supergroup: 'Adult Learning', coachee: 'Alex Rivera', email: 'alex@example.com', coach: 'Jordan Lee', date: '2026-07-10', l7: 0.40, n7: 0.55 },
  { supergroup: 'Adult Learning', coachee: 'Dana Kim', email: 'dana@example.com', coach: 'Jordan Lee', date: '2026-07-09', l7: 0.50, n7: 0.62 },
  // College — both drop → poor
  { supergroup: 'College', coachee: 'Sam Patel', email: 'sam@example.com', coach: 'Morgan Diaz', date: '2026-07-08', l7: 0.55, n7: 0.485 },
  { supergroup: 'College', coachee: 'Priya Shah', email: 'priya@example.com', coach: 'Morgan Diaz', date: '2026-07-10', l7: 0.60, n7: 0.50 },
  // ELD — one up, one down
  { supergroup: 'ELD', coachee: 'Marco Ruiz', email: 'marco@example.com', coach: 'Riley Fox', date: '2026-07-08', l7: 0.30, n7: 0.42 },
  { supergroup: 'ELD', coachee: 'Lena Ortiz', email: 'lena@example.com', coach: 'Riley Fox', date: '2026-07-09', l7: 0.45, n7: 0.40 },
  // High School — one effective, one still pending
  { supergroup: 'High School', coachee: 'Noah Park', email: 'noah@example.com', coach: 'Taylor Reed', date: '2026-07-10', l7: 0.48, n7: 0.60 },
  { supergroup: 'High School', coachee: 'Casey Ng', email: 'casey@example.com', coach: 'Taylor Reed', date: '2026-07-20', l7: 0.45, pending: true },
  // Prof Certs — one up, one down
  { supergroup: 'Prof Certs', coachee: 'Ivy Chen', email: 'ivy@example.com', coach: 'Jamie Wu', date: '2026-07-08', l7: 0.52, n7: 0.58 },
  { supergroup: 'Prof Certs', coachee: 'Omar Ali', email: 'omar@example.com', coach: 'Jamie Wu', date: '2026-07-09', l7: 0.50, n7: 0.47 },
];

/**
 * Simulated coaching-simulation history for a rep: a few scheduled sims with
 * scores, spanning the L7 window through the N7 window (pending reps only get
 * pre-coaching sims). Deterministic so the demo is stable.
 */
function demoSimulations(r) {
  const day = toUtcDate(r.date);
  const after = typeof r.n7 === 'number' ? r.n7 : r.l7;
  const planned = [
    { offset: -6, score: r.l7 },
    { offset: -2, score: (r.l7 + after) / 2 },
  ];
  if (!r.pending && typeof r.n7 === 'number') {
    planned.push({ offset: 3, score: r.n7 });
    planned.push({ offset: 6, score: r.n7 });
  } else {
    planned.push({ offset: 2, score: r.l7 });
  }
  return planned.map((p, i) => ({
    id: `${r.email}-sim-${i + 1}`,
    name: `Simulation ${i + 1}`,
    scheduledDate: formatUtcDate(addDays(day, p.offset)),
    score: Number(Number(p.score).toFixed(3)),
  }));
}

export function loadDemoCoachingEvents() {
  return DEMO_REPS.map((r, i) => ({
    id: `${r.email}|${r.date}|${i + 2}`,
    coachee: r.coachee,
    coacheeEmail: r.email,
    coach: r.coach,
    supergroup: r.supergroup,
    coachingDate: r.date,
    sheetRow: i + 2,
    simulations: demoSimulations(r),
  }));
}

/**
 * Build the per-rep daily pGC series. Each rep gets a flat L7 window and,
 * unless pending, a flat N7 window — so scored L7/N7 equal the catalog values.
 */
export function loadDemoPgcSeries() {
  const out = [];
  for (const r of DEMO_REPS) {
    const day = toUtcDate(r.date);
    const l7Start = formatUtcDate(addDays(day, -7));
    out.push(...dailySeries(r.email, l7Start, 7, r.l7, 0));
    if (!r.pending && typeof r.n7 === 'number') {
      const n7Start = formatUtcDate(addDays(day, 1));
      out.push(...dailySeries(r.email, n7Start, 7, r.n7, 0));
    }
  }
  return out;
}
