import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  addDays,
  formatUtcDate,
  isN7Complete,
  l7Window,
  n7Window,
  scoreCoachingEvent,
  summarizeByCoach,
  summarizeBySupergroup,
  verdictFromDelta,
  windowMean,
} from '../src/score.js';
import { loadDemoCoachingEvents, loadDemoPgcSeries, DEMO_AS_OF } from '../src/demo.js';
import { indexSeriesByPerson } from '../src/looker.js';
import { runSync } from '../src/sync.js';

describe('windows', () => {
  it('excludes coaching day from L7 and N7', () => {
    assert.deepEqual(l7Window('2026-07-10'), { start: '2026-07-03', end: '2026-07-09' });
    assert.deepEqual(n7Window('2026-07-10'), { start: '2026-07-11', end: '2026-07-17' });
  });

  it('marks N7 complete on coaching_date + 8 calendar days', () => {
    assert.equal(isN7Complete('2026-07-10', '2026-07-17'), false);
    assert.equal(isN7Complete('2026-07-10', '2026-07-18'), true);
  });
});

describe('scoring', () => {
  it('labels lift as effective and drop as poor', () => {
    assert.equal(verdictFromDelta(0.01), 'effective');
    assert.equal(verdictFromDelta(0), 'poor');
    assert.equal(verdictFromDelta(-0.02), 'poor');
  });

  it('averages only days inside the window', () => {
    const series = [
      { date: '2026-07-02', pgc: 1 },
      { date: '2026-07-03', pgc: 0.4 },
      { date: '2026-07-09', pgc: 0.6 },
      { date: '2026-07-10', pgc: 9 },
    ];
    assert.equal(windowMean(series, l7Window('2026-07-10')), 0.5);
  });

  it('keeps sessions pending until N7 completes', () => {
    const scored = scoreCoachingEvent(
      {
        coacheeEmail: 'a@x.com',
        coach: 'C',
        coachingDate: '2026-07-20',
      },
      [{ date: '2026-07-15', pgc: 0.4 }],
      '2026-07-24'
    );
    assert.equal(scored.verdict, 'pending');
    assert.equal(scored.n7Complete, false);
  });
});

describe('demo bank', () => {
  it('scores alex effective, sam poor, casey pending', async () => {
    const result = await runSync({
      config: {
        mode: 'demo',
        sheet: { writeback: false, spreadsheetId: 'demo', tab: 'Coaching' },
        looker: {},
        blobStore: 'pgc-test-bank',
      },
      asOf: DEMO_AS_OF,
      writeback: false,
    });

    const byEmail = Object.fromEntries(
      result.bank.sessions.map((s) => [s.coacheeEmail, s])
    );
    assert.equal(byEmail['alex@example.com'].verdict, 'effective');
    assert.ok(byEmail['alex@example.com'].delta > 0);
    assert.equal(byEmail['alex@example.com'].supergroup, 'Adult Learning');
    assert.equal(byEmail['sam@example.com'].verdict, 'poor');
    assert.ok(byEmail['sam@example.com'].delta <= 0);
    assert.equal(byEmail['casey@example.com'].verdict, 'pending');

    const jordan = result.bank.coaches.find((c) => c.coach === 'Jordan Lee');
    assert.ok(jordan);
    assert.equal(jordan.effective, 2);
    assert.equal(jordan.poor, 0);
    assert.equal(jordan.effectivityRate, 1);
  });

  it('rolls up the five Consumer Sales supergroups', async () => {
    const result = await runSync({
      config: {
        mode: 'demo',
        sheet: { writeback: false, spreadsheetId: 'demo', tab: 'Coaching' },
        looker: {},
        blobStore: 'pgc-test-bank',
      },
      asOf: DEMO_AS_OF,
      writeback: false,
    });

    const groups = result.bank.supergroups;
    assert.equal(groups.length, 5);
    assert.deepEqual(
      groups.map((g) => g.supergroup),
      ['Adult Learning', 'College', 'ELD', 'High School', 'Prof Certs']
    );

    const adult = groups.find((g) => g.supergroup === 'Adult Learning');
    // L7 mean(0.40, 0.50) = 0.45; N7 mean(0.55, 0.62) = 0.585; delta = 0.135
    assert.equal(adult.l7, 0.45);
    assert.equal(adult.n7, 0.585);
    assert.ok(adult.delta > 0);
    assert.equal(adult.effective, 2);
    assert.equal(adult.reps.length, 2);

    const college = groups.find((g) => g.supergroup === 'College');
    assert.ok(college.delta < 0);
    assert.equal(college.poor, 2);

    const highSchool = groups.find((g) => g.supergroup === 'High School');
    assert.equal(highSchool.pending, 1);
    // Pending rep is excluded from the L7/N7 means.
    assert.equal(highSchool.reps.length, 2);
    assert.equal(highSchool.n7, 0.6);
  });

  it('indexes demo series by person', () => {
    const map = indexSeriesByPerson(loadDemoPgcSeries());
    assert.ok(map.get('alex@example.com').length >= 14);
    assert.equal(loadDemoCoachingEvents().length, 10);
  });

  it('attaches scheduled simulations to each rep', async () => {
    const result = await runSync({
      config: {
        mode: 'demo',
        sheet: { writeback: false, spreadsheetId: 'demo', tab: 'Coaching' },
        looker: {},
        blobStore: 'pgc-test-bank',
      },
      asOf: DEMO_AS_OF,
      writeback: false,
    });
    const alex = result.bank.sessions.find((s) => s.coacheeEmail === 'alex@example.com');
    assert.ok(Array.isArray(alex.simulations));
    assert.ok(alex.simulations.length >= 1);
    assert.ok(alex.simulations[0].scheduledDate);
    assert.equal(typeof alex.simulations[0].score, 'number');
  });
});

describe('helpers', () => {
  it('formats utc dates stably', () => {
    assert.equal(formatUtcDate(addDays('2026-07-10', 1)), '2026-07-11');
  });

  it('summarizes coaches', () => {
    const summary = summarizeByCoach([
      { coach: 'A', verdict: 'effective', delta: 0.1 },
      { coach: 'A', verdict: 'poor', delta: -0.05 },
      { coach: 'B', verdict: 'pending', delta: null },
    ]);
    assert.equal(summary[0].coach, 'A');
    assert.equal(summary[0].effectivityRate, 0.5);
  });

  it('summarizes supergroups with overall L7/N7/delta', () => {
    const summary = summarizeBySupergroup([
      { supergroup: 'College', verdict: 'effective', l7: 0.4, n7: 0.6, delta: 0.2 },
      { supergroup: 'College', verdict: 'poor', l7: 0.6, n7: 0.5, delta: -0.1 },
      { supergroup: 'College', verdict: 'pending', l7: null, n7: null, delta: null },
    ]);
    assert.equal(summary.length, 1);
    const college = summary[0];
    assert.equal(college.supergroup, 'College');
    assert.equal(college.sessions, 3);
    assert.equal(college.pending, 1);
    assert.equal(college.l7, 0.5); // mean(0.4, 0.6)
    assert.equal(college.n7, 0.55); // mean(0.6, 0.5)
    assert.equal(college.effectivityRate, 0.5); // 1 effective of 2 scored
    assert.equal(college.reps.length, 3);
  });
});
