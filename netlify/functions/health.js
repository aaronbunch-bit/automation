import { loadConfig } from '../../src/config.js';
import { loadBank } from '../../src/bank.js';

export async function handler() {
  const config = loadConfig();
  let bankUpdatedAt = null;
  try {
    const bank = await loadBank(config.blobStore);
    bankUpdatedAt = bank.updatedAt;
  } catch {
    // ignore
  }

  return {
    statusCode: 200,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body: JSON.stringify({
      ok: true,
      service: 'pgc-coaching-bank',
      mode: config.mode,
      spreadsheetId: config.sheet.spreadsheetId,
      tab: config.sheet.tab,
      bankUpdatedAt,
      time: new Date().toISOString(),
    }),
  };
}
