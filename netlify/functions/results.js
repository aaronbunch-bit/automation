import { loadBank } from '../../src/bank.js';
import { loadConfig } from '../../src/config.js';

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*',
    },
    body: JSON.stringify(body),
  };
}

/**
 * GET /api/results — scored sessions + coach rollups from the bank.
 * Query: ?coach=Name&verdict=effective|poor|pending
 */
export async function handler(event) {
  try {
    const config = loadConfig();
    const bank = await loadBank(config.blobStore);
    const qs = event.queryStringParameters || {};

    let sessions = bank.sessions || [];
    if (qs.coach) {
      const needle = qs.coach.toLowerCase();
      sessions = sessions.filter((s) => (s.coach || '').toLowerCase().includes(needle));
    }
    if (qs.verdict) {
      sessions = sessions.filter((s) => s.verdict === qs.verdict);
    }

    return json(200, {
      updatedAt: bank.updatedAt,
      lastSync: bank.lastSync,
      asOf: bank.asOf,
      mode: bank.mode,
      meta: bank.meta,
      coaches: bank.coaches || [],
      supergroups: bank.supergroups || [],
      sessions,
    });
  } catch (err) {
    console.error(err);
    return json(500, { error: err.message || String(err) });
  }
}
