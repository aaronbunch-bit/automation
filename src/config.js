/**
 * Runtime configuration for the pGC coaching effectivity bank.
 * Sheet column headers and Looker field names are env-driven so the
 * private coaching sheet schema can be wired without code changes.
 */

function env(name, fallback = '') {
  const value = process.env[name];
  return value == null || value === '' ? fallback : value;
}

function envBool(name, fallback = false) {
  const value = env(name, fallback ? 'true' : 'false').toLowerCase();
  return value === '1' || value === 'true' || value === 'yes';
}

export function loadConfig() {
  return {
    // Defaults to demo so the bank boots without secrets; set PGC_MODE=live in Netlify.
    mode: env('PGC_MODE', 'demo').toLowerCase(),
    sheet: {
      spreadsheetId: env(
        'GOOGLE_SHEETS_SPREADSHEET_ID',
        '1ClH8uYgfiKMjPJ1-ooFj3bsTel1r6vHSZ3focyjLe0A'
      ),
      tab: env('GOOGLE_SHEETS_TAB', 'Coaching'),
      serviceAccountEmail: env('GOOGLE_SERVICE_ACCOUNT_EMAIL'),
      serviceAccountPrivateKey: env('GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY').replace(
        /\\n/g,
        '\n'
      ),
      writeback: envBool('SHEET_WRITEBACK', true),
      columns: {
        coachee: env('SHEET_COL_COACHEE', 'Coachee'),
        coacheeEmail: env('SHEET_COL_COACHEE_EMAIL', 'Coachee Email'),
        coach: env('SHEET_COL_COACH', 'Coach'),
        coachingDate: env('SHEET_COL_COACHING_DATE', 'Coaching Date'),
        l7: env('SHEET_COL_L7', 'pGC L7'),
        n7: env('SHEET_COL_N7', 'pGC N7'),
        delta: env('SHEET_COL_DELTA', 'pGC Delta'),
        verdict: env('SHEET_COL_VERDICT', 'Verdict'),
        scoredAt: env('SHEET_COL_SCORED_AT', 'Scored At'),
      },
    },
    looker: {
      baseUrl: env('LOOKER_BASE_URL').replace(/\/$/, ''),
      clientId: env('LOOKER_CLIENT_ID'),
      clientSecret: env('LOOKER_CLIENT_SECRET'),
      model: env('LOOKER_MODEL'),
      explore: env('LOOKER_EXPLORE'),
      personField: env('LOOKER_PERSON_FIELD', 'user.email'),
      dateField: env('LOOKER_DATE_FIELD', 'performance.date'),
      pgcField: env('LOOKER_PGC_FIELD', 'performance.pgc'),
      lookId: env('LOOKER_LOOK_ID'),
    },
    blobStore: env('PGC_BLOB_STORE', 'pgc-coaching-bank'),
  };
}

export function assertLiveCredentials(config) {
  const missing = [];
  if (!config.sheet.serviceAccountEmail) missing.push('GOOGLE_SERVICE_ACCOUNT_EMAIL');
  if (!config.sheet.serviceAccountPrivateKey) {
    missing.push('GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY');
  }
  if (!config.looker.baseUrl) missing.push('LOOKER_BASE_URL');
  if (!config.looker.clientId) missing.push('LOOKER_CLIENT_ID');
  if (!config.looker.clientSecret) missing.push('LOOKER_CLIENT_SECRET');
  if (!config.looker.lookId && (!config.looker.model || !config.looker.explore)) {
    missing.push('LOOKER_LOOK_ID or LOOKER_MODEL+LOOKER_EXPLORE');
  }
  if (missing.length) {
    throw new Error(
      `Missing required credentials for live mode: ${missing.join(', ')}. ` +
        'Set PGC_MODE=demo to run with fixtures, or provide the secrets.'
    );
  }
}
