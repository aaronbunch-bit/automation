#!/usr/bin/env node
/**
 * Local CLI: PGC_MODE=demo node scripts/run-sync.js
 */
import { runSync } from '../src/sync.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig();
if (process.argv.includes('--demo')) config.mode = 'demo';

const result = await runSync({ config });
console.log(JSON.stringify({
  ok: true,
  backend: result.backend,
  lastSync: result.bank.lastSync,
  mode: result.bank.mode,
  sessions: result.bank.sessions,
  coaches: result.bank.coaches,
}, null, 2));
