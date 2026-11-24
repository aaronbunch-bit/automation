/** Shared final booking gate. Library callers use this project's lock and ledger. */
function commitReflexActivity(headers, payload, audience, requestedCap) {
  var lock = LockService.getScriptLock();
  var ownsLock = !lock.hasLock();
  if (ownsLock) lock.waitLock(30000);
  try {
    var start = Number(payload.start_time), end = Number(payload.end_time);
    if (!isFinite(start) || !isFinite(end) || end <= start || start <= Date.now() / 1000) throw Error('Invalid ReflexAI booking interval');
    var group = rcAudience_(audience);
    if (!group) throw Error('Cannot verify ReflexAI audience');
    var cap = Number(requestedCap);
    cap = isFinite(cap) && cap >= 1 ? Math.min(3, Math.floor(cap)) : 3;
    var context = rcContext_(headers);
    var candidate = context.people[String(payload.agent_id)];
    if (!candidate || candidate.groups.indexOf(group) < 0) throw Error('Booking audience does not match Assembled queues');
    if (String(payload.type_id) !== context.trainingId) throw Error('ReflexAI booking must use Training activity type');
    var activities = rcActivities_(headers, start, end);
    var spans = rcSpans_(activities, context, group);
    var props = PropertiesService.getScriptProperties();
    var ledger = props.getProperties();
    Object.keys(ledger).filter(function(key) { return key.indexOf('REFLEX_CAP_') === 0; }).forEach(function(key) {
      var entry = JSON.parse(ledger[key]);
      if (entry.end <= Date.now() / 1000) { props.deleteProperty(key); return; }
      // Retain uncertain and not-yet-visible writes until their scheduled end.
      // A successful API write cannot disappear from capacity due to a later
      // spreadsheet error or a delayed Assembled read.
      var visible = activities.some(function(a) {
        return (entry.id && String(a.id) === entry.id) || String(a.description || '').indexOf(entry.marker) >= 0;
      });
      if (!visible && entry.group === group) spans.push(entry);
    });
    if (rcPeak_(spans, start, end) >= cap) throw Error('Audience overlap limit reached; no booking was written');
    var marker = '[RCAP:' + Utilities.getUuid() + ']';
    var key = 'REFLEX_CAP_' + marker.slice(6, -1);
    var reservation = { group: group, start: start, end: end, marker: marker, id: '' };
    props.setProperty(key, JSON.stringify(reservation));
    var body = Object.assign({}, payload, {
      allow_conflicts: false,
      description: String(payload.description || '') + ' ' + marker
    });
    // Never retry a write. Keep the reservation if its outcome is uncertain.
    var response = UrlFetchApp.fetch('https://api.assembledhq.com/v0/activities', {
      method: 'post', headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw Error('Assembled write not confirmed; reservation retained for review');
    var result = JSON.parse(response.getContentText() || '{}');
    var id = String(result.id || (result.activity && result.activity.id) || '');
    if (!id) throw Error('Assembled returned no activity ID; reservation retained for review');
    reservation.id = id;
    props.setProperty(key, JSON.stringify(reservation));
    return result;
  } finally {
    if (ownsLock) lock.releaseLock();
  }
}

function rcAudience_(value) {
  var key = String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (key.indexOf('highschool') >= 0) return 'high school';
  if (key.indexOf('adultlearn') >= 0) return 'adult learning';
  if (key.indexOf('college') >= 0) return 'college';
  if (key.indexOf('elementary') >= 0 || key === 'eld') return 'eld';
  if (key.indexOf('profcert') >= 0 || key.indexOf('professionalcert') >= 0) return 'prof certs';
  return '';
}
function rcList_(raw) {
  return Array.isArray(raw) ? raw : Object.keys(raw || {}).map(function(id) { return Object.assign({id: id}, raw[id]); });
}
function rcGet_(headers, path, params) {
  var query = Object.keys(params || {}).map(function(k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
  var response = UrlFetchApp.fetch('https://api.assembledhq.com/v0' + path + (query ? '?' + query : ''), {method: 'get', headers: headers, muteHttpExceptions: true});
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw Error('Could not verify live capacity: ' + path + ' HTTP ' + response.getResponseCode());
  return JSON.parse(response.getContentText() || '{}');
}
function rcContext_(headers) {
  var queues = rcList_(rcGet_(headers, '/queues', {}).queues);
  var queueGroups = {};
  queues.forEach(function(q) { var group = rcAudience_(q.name); if (group) queueGroups[String(q.id)] = group; });
  var types = rcList_(rcGet_(headers, '/activity_types', {}).activity_types);
  var training = types.filter(function(t) { return String(t.name || '').trim().toLowerCase() === 'training'; });
  if (training.length !== 1) throw Error('Cannot uniquely resolve Training activity type');
  var people = {};
  for (var offset = 0; offset < 10000;) {
    var response = rcGet_(headers, '/people', {limit: 100, offset: offset});
    var page = rcList_(response.people);
    page.forEach(function(p) {
      var raw = p.queues || p.queue_ids || [];
      var ids = Array.isArray(raw) ? raw.map(function(q) { return typeof q === 'object' ? q.id || q.queue_id : q; }) : Object.keys(raw);
      var groups = [];
      ids.forEach(function(id) { var g = queueGroups[String(id)]; if (g && groups.indexOf(g) < 0) groups.push(g); });
      people[String(p.agent_id || p.id)] = {groups: groups};
    });
    offset += page.length;
    if (page.length < 100 || (Number(response.total) > 0 && offset >= Number(response.total))) return {people: people, trainingId: String(training[0].id)};
  }
  throw Error('People pagination incomplete; no booking written');
}
function rcActivities_(headers, start, end) {
  var response = rcGet_(headers, '/activities', {start_time: Math.floor(start), end_time: Math.ceil(end), include_agents: true, include_activity_types: true});
  var rows = rcList_(response.activities);
  if (response.has_more || response.next || response.next_page || response.next_cursor || Number(response.total) > rows.length) throw Error('Activity response is incomplete; no booking written');
  return rows;
}
function rcSpans_(activities, context, group) {
  var spans = [];
  activities.forEach(function(a) {
    if (String(a.type_id) !== context.trainingId) return;
    var person = context.people[String(a.agent_id)];
    var groups = person ? person.groups : [];
    if (!groups.length) {
      if (/reflex[\s-]*ai|\[TSIM:|\[RCAP:/i.test(String(a.description || ''))) throw Error('Existing ReflexAI block has no verified audience');
      return;
    }
    if (groups.indexOf(group) < 0) return;
    var start = Number(a.start_time), end = Number(a.end_time);
    if (!isFinite(start) || !isFinite(end) || end <= start) throw Error('Existing training block has invalid timing');
    // Conservatively include all Training blocks, matching the targeted scheduler.
    spans.push({start: start, end: end});
  });
  return spans;
}
function rcPeak_(spans, start, end) {
  var events = [];
  spans.forEach(function(s) {
    if (s.start < end && s.end > start) {
      events.push([Math.max(start, s.start), 1]);
      events.push([Math.min(end, s.end), -1]);
    }
  });
  events.sort(function(a,b) { return a[0]-b[0] || a[1]-b[1]; });
  var active=0, peak=0;
  events.forEach(function(e) { active+=e[1]; peak=Math.max(peak,active); });
  return peak;
}
/** Read-only live connection check. Does not write bookings or send messages. */
function verifyReflexCapacityReadOnly() {
  var headers = tsAuthHeaders_(tsGetApiKey_());
  var context = rcContext_(headers);
  var start = Date.now()/1000, end = start + 7*86400;
  var rows = rcActivities_(headers, start, end);
  var counts = {};
  ['high school','adult learning','college','eld','prof certs'].forEach(function(g) { counts[g] = rcSpans_(rows,context,g).length; });
  Logger.log(JSON.stringify({verified: true, people: Object.keys(context.people).length, upcomingTrainingBlocksByAudience: counts, bookingsWritten: 0}));
}
