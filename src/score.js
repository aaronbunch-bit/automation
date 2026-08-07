/**
 * L7 / N7 windowing and coaching effectivity scoring.
 *
 * L7 = mean pGC over the 7 calendar days before coaching day (excluded)
 * N7 = mean pGC over the 7 calendar days after coaching day (excluded)
 * Effective when delta (N7 − L7) > 0; otherwise poor/neutral.
 */

const MS_PER_DAY = 24 * 60 * 60 * 1000;

/** Parse YYYY-MM-DD or Date into a UTC calendar date at midnight. */
export function toUtcDate(value) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) throw new Error('Invalid Date');
    return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  }
  const raw = String(value).trim();
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Cannot parse date: ${value}`);
  }
  return new Date(
    Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate())
  );
}

export function formatUtcDate(date) {
  return toUtcDate(date).toISOString().slice(0, 10);
}

export function addDays(date, days) {
  const d = toUtcDate(date);
  return new Date(d.getTime() + days * MS_PER_DAY);
}

/** Inclusive [start, end] UTC date strings for the L7 window. */
export function l7Window(coachingDate) {
  const day = toUtcDate(coachingDate);
  return {
    start: formatUtcDate(addDays(day, -7)),
    end: formatUtcDate(addDays(day, -1)),
  };
}

/** Inclusive [start, end] UTC date strings for the N7 window. */
export function n7Window(coachingDate) {
  const day = toUtcDate(coachingDate);
  return {
    start: formatUtcDate(addDays(day, 1)),
    end: formatUtcDate(addDays(day, 7)),
  };
}

/** N7 is complete once the calendar day after N7 end has begun (asOf >= N7.end + 1 day). */
export function isN7Complete(coachingDate, asOf = new Date()) {
  const readyOn = addDays(toUtcDate(coachingDate), 8);
  return toUtcDate(asOf).getTime() >= readyOn.getTime();
}

export function mean(values) {
  const nums = values.filter((v) => typeof v === 'number' && Number.isFinite(v));
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

/**
 * Filter a person-day pGC series into a window and average.
 * @param {Array<{ date: string, pgc: number }>} series
 */
export function windowMean(series, window) {
  const start = toUtcDate(window.start).getTime();
  const end = toUtcDate(window.end).getTime();
  const values = [];
  for (const row of series || []) {
    const t = toUtcDate(row.date).getTime();
    if (t >= start && t <= end) values.push(Number(row.pgc));
  }
  return mean(values);
}

export function verdictFromDelta(delta) {
  if (delta == null || !Number.isFinite(delta)) return 'insufficient_data';
  if (delta > 0) return 'effective';
  return 'poor';
}

/**
 * Score one coaching event against a pGC day series for that coachee.
 * @returns {object} scored row
 */
export function scoreCoachingEvent(event, series, asOf = new Date()) {
  const coachingDate = formatUtcDate(event.coachingDate);
  const l7 = l7Window(coachingDate);
  const n7 = n7Window(coachingDate);
  const ready = isN7Complete(coachingDate, asOf);

  const base = {
    id: event.id || `${event.coacheeEmail || event.coachee}|${coachingDate}`,
    coachee: event.coachee || '',
    coacheeEmail: event.coacheeEmail || '',
    coach: event.coach || '',
    supergroup: event.supergroup || '',
    coachingDate,
    sheetRow: event.sheetRow ?? null,
    l7Window: l7,
    n7Window: n7,
    n7Complete: ready,
    l7: null,
    n7: null,
    delta: null,
    verdict: ready ? 'insufficient_data' : 'pending',
    scoredAt: null,
  };

  if (!ready) return base;

  const l7Mean = windowMean(series, l7);
  const n7Mean = windowMean(series, n7);
  const delta =
    l7Mean == null || n7Mean == null ? null : Number((n7Mean - l7Mean).toFixed(6));

  return {
    ...base,
    l7: l7Mean == null ? null : Number(l7Mean.toFixed(6)),
    n7: n7Mean == null ? null : Number(n7Mean.toFixed(6)),
    delta,
    verdict: verdictFromDelta(delta),
    scoredAt: new Date().toISOString(),
  };
}

/**
 * Aggregate coach-level effectivity rates from scored sessions.
 */
export function summarizeByCoach(scored) {
  const byCoach = new Map();
  for (const row of scored) {
    if (row.verdict === 'pending') continue;
    const key = row.coach || '(unknown)';
    if (!byCoach.has(key)) {
      byCoach.set(key, {
        coach: key,
        sessions: 0,
        effective: 0,
        poor: 0,
        insufficient_data: 0,
        avgDelta: null,
        deltas: [],
      });
    }
    const bucket = byCoach.get(key);
    bucket.sessions += 1;
    if (row.verdict === 'effective') bucket.effective += 1;
    else if (row.verdict === 'poor') bucket.poor += 1;
    else bucket.insufficient_data += 1;
    if (typeof row.delta === 'number') bucket.deltas.push(row.delta);
  }

  return [...byCoach.values()]
    .map((b) => {
      const scoredCount = b.effective + b.poor;
      return {
        coach: b.coach,
        sessions: b.sessions,
        effective: b.effective,
        poor: b.poor,
        insufficient_data: b.insufficient_data,
        effectivityRate: scoredCount ? b.effective / scoredCount : null,
        avgDelta: b.deltas.length ? Number(mean(b.deltas).toFixed(6)) : null,
      };
    })
    .sort((a, b) => (b.effectivityRate ?? -1) - (a.effectivityRate ?? -1));
}

/**
 * Aggregate supergroup-level pGC for the Consumer Sales hierarchy.
 * Overall L7 / N7 are the mean of scored reps' L7 / N7; delta = overall N7 − L7.
 * Each rollup carries its reps so the UI can drill into individual sessions.
 */
export function summarizeBySupergroup(scored) {
  const groups = new Map();
  for (const row of scored) {
    const key = row.supergroup || '(Unassigned)';
    if (!groups.has(key)) {
      groups.set(key, {
        supergroup: key,
        sessions: 0,
        effective: 0,
        poor: 0,
        pending: 0,
        insufficient_data: 0,
        l7s: [],
        n7s: [],
        deltas: [],
        reps: [],
      });
    }
    const bucket = groups.get(key);
    bucket.sessions += 1;
    if (row.verdict === 'effective') bucket.effective += 1;
    else if (row.verdict === 'poor') bucket.poor += 1;
    else if (row.verdict === 'pending') bucket.pending += 1;
    else bucket.insufficient_data += 1;
    if (typeof row.l7 === 'number') bucket.l7s.push(row.l7);
    if (typeof row.n7 === 'number') bucket.n7s.push(row.n7);
    if (typeof row.delta === 'number') bucket.deltas.push(row.delta);
    bucket.reps.push(row);
  }

  return [...groups.values()]
    .map((b) => {
      const scoredCount = b.effective + b.poor;
      const l7 = b.l7s.length ? Number(mean(b.l7s).toFixed(6)) : null;
      const n7 = b.n7s.length ? Number(mean(b.n7s).toFixed(6)) : null;
      const delta =
        l7 == null || n7 == null ? null : Number((n7 - l7).toFixed(6));
      return {
        supergroup: b.supergroup,
        sessions: b.sessions,
        effective: b.effective,
        poor: b.poor,
        pending: b.pending,
        insufficient_data: b.insufficient_data,
        l7,
        n7,
        delta,
        effectivityRate: scoredCount ? b.effective / scoredCount : null,
        reps: b.reps,
      };
    })
    .sort((a, b) => a.supergroup.localeCompare(b.supergroup));
}
