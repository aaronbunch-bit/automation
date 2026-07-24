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
    assert.equal(byEmail['sam@example.com'].verdict, 'poor');
    assert.ok(byEmail['sam@example.com'].delta <= 0);
    assert.equal(byEmail['casey@example.com'].verdict, 'pending');

    const jordan = result.bank.coaches.find((c) => c.coach === 'Jordan Lee');
    assert.ok(jordan);
    assert.equal(jordan.effective, 1);
    assert.equal(jordan.poor, 1);
    assert.equal(jordan.effectivityRate, 0.5);
  });

  it('indexes demo series by person', () => {
    const map = indexSeriesByPerson(loadDemoPgcSeries());
    assert.ok(map.get('alex@example.com').length >= 14);
    assert.equal(loadDemoCoachingEvents().length, 3);
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
});
