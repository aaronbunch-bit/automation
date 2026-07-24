/**
 * Looker API 4.0 client for daily pGC series.
 * Auth: API3 client_id / client_secret → access token.
 */

async function lookerFetch(baseUrl, path, { method = 'GET', token, body } = {}) {
  const url = `${baseUrl}/api/4.0${path}`;
  const headers = { Accept: 'application/json' };
  if (token) headers.Authorization = `token ${token}`;
  if (body != null) headers['Content-Type'] = 'application/json';

  const res = await fetch(url, {
    method,
    headers,
    body: body == null ? undefined : JSON.stringify(body),
  });
  const text = await res.text();
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = text;
  }
  if (!res.ok) {
    const detail = typeof data === 'string' ? data : JSON.stringify(data);
    throw new Error(`Looker ${method} ${path} failed (${res.status}): ${detail}`);
  }
  return data;
}

export async function lookerLogin(looker) {
  const url = `${looker.baseUrl}/api/4.0/login`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json',
    },
    body: new URLSearchParams({
      client_id: looker.clientId,
      client_secret: looker.clientSecret,
    }),
  });
  const payload = await res.json().catch(() => ({}));
  if (!res.ok || !payload.access_token) {
    throw new Error(`Looker login failed (${res.status}): ${JSON.stringify(payload)}`);
  }
  return payload.access_token;
}

function buildInlineQuery(looker, personKeys, startDate, endDate) {
  const filters = {
    [looker.dateField]: `${startDate} to ${endDate}`,
  };
  if (personKeys.length === 1) {
    filters[looker.personField] = personKeys[0];
  } else if (personKeys.length > 1) {
    filters[looker.personField] = personKeys.join(',');
  }

  return {
    model: looker.model,
    view: looker.explore,
    fields: [looker.personField, looker.dateField, looker.pgcField],
    filters,
    sorts: [`${looker.dateField} asc`],
    limit: '5000',
  };
}

/**
 * Fetch person-day pGC rows from Looker.
 * @returns {Promise<Array<{ personKey: string, date: string, pgc: number }>>}
 */
export async function fetchPgcSeries(looker, { personKeys, startDate, endDate }) {
  if (!personKeys.length) return [];

  const token = await lookerLogin(looker);
  let rows;

  if (looker.lookId) {
    rows = await lookerFetch(
      looker.baseUrl,
      `/looks/${encodeURIComponent(looker.lookId)}/run/json`,
      { method: 'GET', token }
    );
  } else {
    const query = buildInlineQuery(looker, personKeys, startDate, endDate);
    rows = await lookerFetch(looker.baseUrl, '/queries/run/json', {
      method: 'POST',
      token,
      body: query,
    });
  }

  const personSet = new Set(personKeys.map((k) => String(k).toLowerCase()));
  const series = [];
  for (const row of rows || []) {
    const personKey = String(
      row[looker.personField] ?? row.person_key ?? row.email ?? ''
    ).trim();
    const dateRaw = row[looker.dateField] ?? row.date;
    const pgcRaw = row[looker.pgcField] ?? row.pgc;
    if (!personKey || dateRaw == null || pgcRaw == null || pgcRaw === '') continue;
    if (personSet.size && !personSet.has(personKey.toLowerCase())) continue;
    const date = String(dateRaw).slice(0, 10);
    if (date < startDate || date > endDate) continue;
    const pgc = Number(pgcRaw);
    if (!Number.isFinite(pgc)) continue;
    series.push({ personKey, date, pgc });
  }
  return series;
}

/** Group flat series by lowercase person key. */
export function indexSeriesByPerson(series) {
  const map = new Map();
  for (const row of series) {
    const key = String(row.personKey).toLowerCase();
    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ date: row.date, pgc: row.pgc });
  }
  return map;
}
