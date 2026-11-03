// Pulls every classroom slot from the 360Learning API v2 and writes:
//   docs/events.json   – read by the calendar page
//   docs/calendar.ics  – subscribable feed for Outlook / Google Calendar
//
// Runs in GitHub Actions (Node 20+, no dependencies). Credentials come from
// repository secrets and are never written to the output files.

import { writeFile, mkdir } from 'node:fs/promises';

const API_BASE      = (process.env.API_BASE || 'https://app.360learning.com/api/v2').replace(/\/$/, '');
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const PLATFORM_URL  = (process.env.PLATFORM_URL || '').replace(/\/$/, '');   // e.g. https://yourcompany.360learning.com
const SESSION_LINK  = process.env.SESSION_LINK || '';                        // optional fallback link if a session's IDs can't be found
const DAYS_BACK     = Number(process.env.DAYS_BACK || 30);                   // how far back to include past sessions
const DEBUG         = process.env.DEBUG === 'true';

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('Missing CLIENT_ID or CLIENT_SECRET. Add them as repository secrets.');
  process.exit(1);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── Auth ─────────────────────────────────────────────────────────────
async function getToken() {
  const res = await fetch(`${API_BASE}/oauth2/token`, {
    method: 'POST',
    headers: { accept: 'application/json', 'content-type': 'application/json' },
    body: JSON.stringify({ grant_type: 'client_credentials', client_id: CLIENT_ID, client_secret: CLIENT_SECRET }),
  });
  if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
  return (await res.json()).access_token;
}

// ── Requests with retry on rate limits ──────────────────────────────
async function apiGet(url, token, attempt = 1) {
  const res = await fetch(url, {
    headers: { accept: 'application/json', authorization: `Bearer ${token}`, '360-api-version': 'v2.0' },
  });
  if (res.status === 429 && attempt <= 5) {
    const wait = Number(res.headers.get('retry-after')) * 1000 || 2000 * attempt;
    console.log(`Rate limited, waiting ${wait}ms…`);
    await sleep(wait);
    return apiGet(url, token, attempt + 1);
  }
  if (!res.ok) throw new Error(`GET ${url} failed: ${res.status} ${await res.text()}`);
  return { body: await res.json(), link: res.headers.get('link') };
}

// Follows the Link: <url>; rel="next" header until the last page.
async function getAll(path, token) {
  let url = `${API_BASE}${path}`;
  const items = [];
  while (url) {
    const { body, link } = await apiGet(url, token);
    const page = Array.isArray(body) ? body : body.data || body.items || body.results || [];
    items.push(...page);
    const next = link && link.match(/<([^>]+)>\s*;\s*rel="?next"?/);
    url = next ? next[1] : null;
  }
  return items;
}

// Runs async jobs with a small concurrency limit (kind to rate limits).
async function mapLimit(list, limit, fn) {
  const out = new Array(list.length);
  let i = 0;
  await Promise.all(Array.from({ length: Math.min(limit, list.length) }, async () => {
    while (i < list.length) { const idx = i++; out[idx] = await fn(list[idx], idx); }
  }));
  return out;
}

const idOf = (o) => o && (o._id || o.id);
const personName = (u) => {
  if (!u) return null;
  const full = [u.firstName, u.lastName].filter(Boolean).join(' ').trim();
  return full || u.name || u.fullName || u.mail || u.email || null;
};

// Finds the path ID and session ID anywhere inside a slot (e.g. in its "context"),
// so each event can link to: PLATFORM_URL/paths/{pathId}/home?sessionId={sessionId}
const HEX24 = /^[a-f0-9]{24}$/i;
function findIds(obj) {
  const found = { pathId: null, sessionId: null };
  (function walk(o, keyPath) {
    if (!o || typeof o !== 'object') return;
    for (const [k, v] of Object.entries(o)) {
      const p = keyPath ? `${keyPath}.${k}` : k;
      if (typeof v === 'string' && HEX24.test(v)) {
        if (/session/i.test(p) && !found.sessionId) found.sessionId = v;
        else if (/path/i.test(p) && !/session/i.test(p) && !found.pathId) found.pathId = v;
      } else if (v && typeof v === 'object') walk(v, p);
    }
  })(obj && obj.context ? { context: obj.context } : obj, '');
  return found;
}

// ── Build the event list ────────────────────────────────────────────
async function main() {
  const token = await getToken();
  console.log('Authenticated.');

  const classrooms = await getAll('/classrooms', token);
  console.log(`Found ${classrooms.length} classrooms.`);

  const cutoff = Date.now() - DAYS_BACK * 864e5;
  const trainerIds = new Set();
  const debugInfo = { classroomSample: null, slotSamples: [], userLookup: null, pathLookup: null };
  const strip = (o) => JSON.parse(JSON.stringify(o, (k, v) => (/location|url|link/i.test(k) ? '[hidden]' : v)));
  if (DEBUG && classrooms[0]) debugInfo.classroomSample = strip(classrooms[0]);

  const perClassroom = await mapLimit(classrooms, 4, async (c) => {
    const slots = await getAll(`/classrooms/${idOf(c)}/slots`, token);
    if (DEBUG && slots[0] && debugInfo.slotSamples.length < 2) debugInfo.slotSamples.push(strip(slots[0]));
    return slots
      .filter((s) => s.startDate && new Date(s.endDate || s.startDate).getTime() >= cutoff)
      .map((s) => {
        (s.trainerIds || s.trainers || []).forEach((t) => trainerIds.add(typeof t === 'string' ? t : idOf(t)));
        return { classroom: c, slot: s };
      });
  });
  const rows = perClassroom.flat();
  console.log(`Found ${rows.length} upcoming or recent slots.`);

  // Resolve trainer names (needs the users:read scope; skipped quietly if not allowed)
  const names = {};
  await mapLimit([...trainerIds], 4, async (uid) => {
    try {
      const u = (await apiGet(`${API_BASE}/users/${uid}`, token)).body;
      names[uid] = personName(u);
      if (DEBUG && !debugInfo.userLookup) debugInfo.userLookup = { ok: true, fields: Object.keys(u) };
    } catch (err) {
      names[uid] = null;
      if (DEBUG && !debugInfo.userLookup) debugInfo.userLookup = { ok: false, error: String(err.message).slice(0, 200) };
    }
  });

  // Map each slot's context ID to its path, so links can open the exact session:
  //   PLATFORM_URL/paths/{pathId}/home?sessionId={sessionId}
  // Needs the paths:read permission; falls back to the platform home page without it.
  const sessionToPath = {}, pathIds = new Set();
  try {
    const paths = await getAll('/paths', token);
    await mapLimit(paths, 4, async (p) => {
      const pid = idOf(p); pathIds.add(pid);
      try { (await getAll(`/paths/${pid}/sessions`, token)).forEach((ses) => { sessionToPath[idOf(ses)] = ses.pathId || pid; }); }
      catch { /* path without sessions access */ }
    });
    if (DEBUG) debugInfo.pathLookup = { ok: true, paths: pathIds.size, sessions: Object.keys(sessionToPath).length };
  } catch (err) {
    console.log('Could not read paths (add the "Paths: read" permission for direct session links).');
    if (DEBUG) debugInfo.pathLookup = { ok: false, error: String(err.message).slice(0, 200) };
  }
  const linkFor = (slot) => {
    if (!PLATFORM_URL) return SESSION_LINK || '';
    const ctx = slot.context && idOf(slot.context);
    if (ctx && sessionToPath[ctx]) return `${PLATFORM_URL}/paths/${sessionToPath[ctx]}/home?sessionId=${ctx}`;
    if (ctx && pathIds.has(ctx)) return `${PLATFORM_URL}/paths/${ctx}/home`;
    const { pathId, sessionId } = findIds(slot);
    if (pathId) return `${PLATFORM_URL}/paths/${pathId}/home${sessionId ? `?sessionId=${sessionId}` : ''}`;
    return SESSION_LINK || PLATFORM_URL;
  };

  const events = rows.map(({ classroom, slot }) => {
    const trainers = (slot.trainerIds || slot.trainers || [])
      .map((t) => (typeof t === 'string' ? names[t] : personName(t)))
      .filter(Boolean);
    const location = slot.location || '';
    const isUrl = /^https?:\/\//i.test(location);
    const link = linkFor(slot);
    return {
      id: idOf(slot),
      title: slot.name || classroom.name || classroom.title || 'Classroom session',
      classroom: classroom.name || classroom.title || '',
      start: slot.startDate,
      end: slot.endDate,
      timezone: slot.timezone || null,
      durationHours: slot.reportedDuration ?? null,
      virtual: Boolean(slot.virtual),
      // Virtual meeting links are kept off the public page; learners join from the platform.
      location: slot.virtual ? 'Virtual' : (isUrl ? '' : location),
      trainers,
      capacity: slot.maxCapacity ?? null,
      link,
    };
  }).sort((a, b) => a.start.localeCompare(b.start));

  await mkdir('docs', { recursive: true });
  await writeFile('docs/events.json', JSON.stringify({ updatedAt: new Date().toISOString(), sample: false, events, ...(DEBUG ? { debug: { trainerIdsFound: trainerIds.size, ...debugInfo } } : {}) }, null, 2));
  await writeFile('docs/calendar.ics', toICS(events));
  console.log(`Wrote ${events.length} events to docs/events.json and docs/calendar.ics.`);
}

// ── ICS feed ────────────────────────────────────────────────────────
function toICS(events) {
  const fmt = (d) => new Date(d).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  const esc = (s = '') => String(s).replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/,/g, '\\,').replace(/;/g, '\\;');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//360Learning Live Calendar//EN', 'CALSCALE:GREGORIAN', 'X-WR-CALNAME:Upcoming classrooms'];
  for (const e of events) {
    lines.push(
      'BEGIN:VEVENT',
      `UID:${e.id}@360learning-live-calendar`,
      `DTSTAMP:${fmt(Date.now())}`,
      `DTSTART:${fmt(e.start)}`,
      `DTEND:${fmt(e.end || e.start)}`,
      `SUMMARY:${esc(e.title)}`,
      `LOCATION:${esc(e.location)}`,
      `DESCRIPTION:${esc([e.trainers.length ? 'Trainer: ' + e.trainers.join(', ') : '', e.link ? 'Register: ' + e.link : ''].filter(Boolean).join('\n'))}`,
      e.link ? `URL:${e.link}` : '',
      'END:VEVENT',
    );
  }
  lines.push('END:VCALENDAR');
  return lines.filter(Boolean).join('\r\n');
}

main().catch((err) => { console.error(err.message); process.exit(1); });
