#!/usr/bin/env python3
"""Build the public aliyah page from a snapshot of the full Menashe board database.

Usage: python3 scripts/build.py <db_dir>

<db_dir> holds the board's documents as JSON, one folder per collection:
  gantt/main.json, welcomes/main.json, visa/main.json, ticketing/{meta,assign,p1..p6}.json

Only aliyah material from 11.11.2026 is published. Interviews, budget and documents
are never read. Olim appear only as counts: no names, birth dates, passports,
family files or free-text notes leave this script.
The one exception is roster.json (from scripts/import_flights.py): full name, the last 4
passport characters and the flight number per person.
"""
import collections, datetime, hashlib, json, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
START = '2026-11-11'
SOURCES = {'gantt', 'rabbis', 'kh', 'donors'}


def load(db, path, default=None):
    p = os.path.join(db, path)
    if not os.path.exists(p):
        return default
    with open(p, encoding='utf8') as f:
        return json.load(f)


def scrub(t):
    """Remove booking references and similar codes from free text."""
    if not t:
        return t
    t = re.sub(r'\s*[^.]*כרטיס\s+[A-Z0-9]{6}[^.]*\.?', '', t)
    t = re.sub(r'\bPNR\s*[A-Z0-9]{5,7}\b', '', t)
    return t.strip()


def age(dob, iso):
    if not dob or not iso:
        return None
    a = [int(x) for x in dob.split('-')]
    b = [int(x) for x in iso.split('-')]
    return b[0] - a[0] - (1 if (b[1] < a[1] or (b[1] == a[1] and b[2] < a[2])) else 0)


BANDS = [(0, 2, 'מתחת לגיל 2'), (2, 5, '2 עד 4'), (5, 12, '5 עד 11'), (12, 18, '12 עד 17'),
         (18, 30, '18 עד 29'), (30, 50, '30 עד 49'), (50, 65, '50 עד 64'), (65, 200, '65 ומעלה')]


def band(a):
    if a is None:
        return 'לא ידוע'
    for lo, hi, l in BANDS:
        if lo <= a < hi:
            return l
    return 'לא ידוע'


def fare(a):
    if a is None:
        return 'לא ידוע'
    return 'infant' if a < 2 else 'child' if a < 12 else 'adult'


def add_days(iso, n):
    d = datetime.date.fromisoformat(iso) + datetime.timedelta(days=n)
    return d.isoformat()


def build_schedule(g):
    seen, events = set(), []
    for e in g.get('events', []):
        if e.get('date', '') < START or e.get('source') not in SOURCES:
            continue
        key = (e['date'], e.get('text', ''))
        if key in seen:
            continue
        seen.add(key)
        ev = {k: e[k] for k in ('date', 'time', 'end', 'to', 'kind', 'source', 'text') if e.get(k)}
        if e.get('note'):
            n = scrub(e['note'])
            if n:
                ev['note'] = n
        events.append(ev)

    people = []
    for p in g.get('people', []):
        segs = []
        for s in p.get('segments', []):
            if s.get('to', '') <= START and s.get('from', '') < START:
                continue
            segs.append({'from': max(s['from'], START), 'to': s['to'], 'loc': s.get('loc'), 'place': s.get('place')})
        if segs:
            people.append({'name': p['name'], 'role': p.get('role', ''), 'count': p.get('count', 1), 'segs': segs})

    groups = [{k: o.get(k) for k in ('id', 'count', 'manipur', 'origin', 'from', 'to', 'place')} for o in g.get('olim', [])]
    used = {s['place'] for p in people for s in p['segs']} | {o['place'] for o in groups}
    places = [{k: pl.get(k) for k in ('id', 'name', 'short', 'city', 'note', 'query')}
              for pl in g.get('places', []) if pl.get('id') in used]
    for pl in places:
        pl['note'] = scrub(pl.get('note'))
    return events, people, groups, places


def build_olim(db):
    meta = load(db, 'ticketing/meta.json')
    if not meta:
        return None
    amap = (load(db, 'ticketing/assign.json', {}) or {}).get('map', {})
    flights = {f['f']: f for f in meta.get('flights', [])}
    sgs = {s['id']: s for s in meta.get('subgroups', [])}
    people = []
    for i in range(1, 7):
        part = load(db, f'ticketing/p{i}.json')
        if part is None:
            sys.exit(f'ticketing/p{i}.json missing: export all of ticketing (meta, assign, p1..p6) together')
        people += part.get('people', [])
    if load(db, 'ticketing/assign.json') is None:
        sys.exit('ticketing/assign.json missing: export all of ticketing (meta, assign, p1..p6) together')
    if meta.get('total') is not None and len(people) != meta['total']:
        sys.exit(f"ticketing check failed: {len(people)} people in p1..p6, meta total is {meta['total']}")
    for p in people:
        miss = [k for k in ('f', 'hh', 'file', 'pop') if not p.get(k)]
        if miss:
            sys.exit(f"ticketing check failed: record {p.get('id')} lacks {miss}")

    hh_size = collections.Counter((p['f'], p['hh']) for p in people)
    cube = collections.Counter()
    issues = collections.defaultdict(collections.Counter)
    hh_cube = collections.Counter()
    comm = collections.Counter()
    seen_hh = set()
    for p in people:
        fl = flights.get(p['f'], {})
        fdate = fl.get('date')
        sg = amap.get(p['file'], str(p['f']))
        a = age(p.get('dob'), fdate)
        size = hh_size[(p['f'], p['hh'])]
        hhs = '7+' if size >= 7 else str(size)
        center = 'קריית ים' if p.get('pop') == 'manipur' else 'נוף הגליל'
        origin = 'דימאפור' if 'dmu' in (p.get('fl') or []) else 'אייזול'
        exp = p.get('exp')
        pp = 'none' if not exp else 'expired' if exp < fdate else 'short' if exp < add_days(fdate, 182) else 'ok'
        sex = p.get('sex') or ''
        key = (p['f'], sg, p.get('pop', ''), center, sex, band(a), fare(a), p.get('st') or 'לא צוין', hhs, origin, pp)
        cube[key] += 1
        iss = list(p.get('fl') or [])
        if pp in ('expired', 'short'):
            iss.append('pp_' + pp)
        if not p.get('fe'):
            iss.append('noFamEn')
        for x in iss:
            issues[key][x] += 1
        hk = (p['f'], p['hh'])
        if hk not in seen_hh:
            seen_hh.add(hk)
            hh_cube[(p['f'], sg, p.get('pop', ''), center, hhs)] += 1
        comm[(p['f'], sg, p.get('comm') or 'לא צוין')] += 1

    rows = [list(k) + [n, dict(issues[k])] if issues[k] else list(k) + [n] for k, n in sorted(cube.items(), key=lambda x: str(x[0]))]
    hold = collections.Counter(h.get('f') for h in meta.get('onHold', []))
    return {
        'source': meta.get('source', ''), 'imported': meta.get('imported', ''), 'total': len(people),
        'flights': [{k: f.get(k) for k in ('f', 'date', 'carrier', 'route', 'escort')} for f in meta.get('flights', [])],
        'subgroups': [{k: s.get(k) for k in ('id', 'f', 'dom', 'escort', 'target')} for s in meta.get('subgroups', [])],
        'onHold': {str(k): v for k, v in hold.items()},
        'dims': ['f', 'sg', 'pop', 'center', 'sex', 'band', 'fare', 'st', 'hhs', 'origin', 'pp'],
        'cube': rows,
        'hh': [list(k) + [n] for k, n in sorted(hh_cube.items(), key=lambda x: str(x[0]))],
        'comm': [list(k) + [n] for k, n in sorted(comm.items(), key=lambda x: str(x[0]))],
    }


def main():
    db = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'db')
    old = os.path.join(ROOT, 'data.json')
    prev = {}
    if os.path.exists(old):
        with open(old, encoding='utf8') as f:
            prev = json.load(f)

    # A source file that was not exported this run keeps its section from the previous data.json.
    data = {}
    g = load(db, 'gantt/main.json')
    if g:
        events, people, groups, places = build_schedule(g)
        data.update({'start': START, 'end': (g.get('range') or {}).get('to', '2026-11-26'), 'ganttUpdated': g.get('updated', ''),
                     'events': events, 'people': people, 'groups': groups, 'places': places})
    else:
        keys = ('start', 'end', 'ganttUpdated', 'events', 'people', 'groups', 'places')
        if not all(k in prev for k in keys):
            sys.exit('gantt/main.json not found in ' + db + ' and no previous data.json to keep')
        data.update({k: prev[k] for k in keys})

    wdoc = load(db, 'welcomes/main.json')
    if wdoc is not None:
        welcomes = [w for w in (wdoc or {}).get('items', []) if w.get('date', '') >= START]
        for w in welcomes:
            for f in w.get('fields', []):
                f['v'] = scrub(f.get('v', ''))
    else:
        welcomes = prev.get('welcomes', [])
    data['welcomes'] = welcomes

    vdoc = load(db, 'visa/main.json')
    if vdoc is not None:
        visa = {}
        for k, v in ((vdoc or {}).get('flights', {}) or {}).items():
            visa[k] = {x: v.get(x) for x in ('sent', 'ok', 'miss', 'at')}
    else:
        visa = prev.get('visa', {})
    data['visa'] = visa

    data['olim'] = build_olim(db) if os.path.exists(os.path.join(db, 'ticketing/meta.json')) else prev.get('olim')

    # Roster from the flights workbook: name, last 4 passport characters and flight only.
    rdoc = load(db, 'roster.json')
    if rdoc is not None:
        people = sorted(rdoc.get('people', []), key=lambda p: (p['f'], p['he']))
        for p in people:
            if set(p) - {'f', 'he', 'en', 'pp4'} or len(p.get('pp4', '')) > 4:
                sys.exit('roster.json holds fields beyond name, pp4 and flight')
        roster = {'source': rdoc.get('source', ''), 'imported': datetime.date.today().isoformat(), 'people': people}
        cube = collections.Counter()
        for r in (data.get('olim') or {}).get('cube', []):
            cube[r[0]] += r[11]
        rc = collections.Counter(p['f'] for p in people)
        if cube and dict(rc) != dict(cube):
            print('WARNING: roster per flight', dict(sorted(rc.items())), 'differs from board counts', dict(sorted(cube.items())))
    else:
        roster = prev.get('roster')
    if roster:
        data['roster'] = roster

    # Board document versions seen so far (integers only), used to skip unchanged sources.
    src = dict(prev.get('src') or {})
    src.update({k: int(v) for k, v in (load(db, 'versions.json', {}) or {}).items()})
    data['src'] = src

    body = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(',', ':'))
    digest = hashlib.sha256(body.encode()).hexdigest()[:12]

    data['digest'] = digest
    if prev.get('digest') == digest and prev.get('built'):
        data['built'] = prev['built']
        print('data unchanged')
    else:
        data['built'] = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    with open(old, 'w', encoding='utf8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    with open(os.path.join(ROOT, 'scripts', 'template.html'), encoding='utf8') as f:
        tpl = f.read()
    with open(os.path.join(ROOT, 'scripts', 'assets.html'), encoding='utf8') as f:
        sym, ja = f.read().split('<!--JA-->')
    page = (tpl.replace('/*__DATA__*/null', json.dumps(data, ensure_ascii=False).replace('</', '<\\/'))
               .replace('<!--__SYMBOLS__-->', sym).replace('__JA__', ja.strip()))
    with open(os.path.join(ROOT, 'index.html'), 'w', encoding='utf8') as f:
        f.write(page)
    print('built', digest)


if __name__ == '__main__':
    main()
