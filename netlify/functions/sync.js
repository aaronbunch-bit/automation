import { runSync } from '../../src/sync.js';
import { loadConfig } from '../../src/config.js';

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*',
    },
    body: JSON.stringify(body, null, 2),
  };
}

/**
 * Scheduled (@daily) and on-demand sync of coaching dates → L7/N7 scores.
 * Manual trigger: POST /.netlify/functions/sync or /api/sync
 */
export async function handler(event) {
  try {
    if (event.httpMethod && event.httpMethod !== 'POST' && event.httpMethod !== 'GET') {
      return json(405, { error: 'Method not allowed' });
    }

    const config = loadConfig();
    // Allow ?mode=demo for a dry run without credentials
    const qs = event.queryStringParameters || {};
    if (qs.mode === 'demo') config.mode = 'demo';

    const result = await runSync({ config });
    return json(200, {
      ok: true,
      backend: result.backend,
      lastSync: result.bank.lastSync,
      asOf: result.bank.asOf,
      mode: result.bank.mode,
      sessions: result.bank.sessions.length,
      coaches: result.bank.coaches.length,
      writeback: result.bank.meta?.writeback,
      summary: result.bank.coaches,
    });
  } catch (err) {
    console.error(err);
    return json(500, { ok: false, error: err.message || String(err) });
  }
}
