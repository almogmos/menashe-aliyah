#!/usr/bin/env python3
"""Turn the flights workbook (the per-flight sheets) into a minimal roster for the public page.

Usage: python3 scripts/import_flights.py <workbook.xlsx> <db_dir> [label]

Writes <db_dir>/roster.json. Only these fields leave the workbook, per person:
  f   flight number to Israel (matched by the date in the sheet name, e.g. "17.11")
  he  full name in Hebrew (family name + given name)
  en  full name in English (given name + family name)
  pp4 last 4 characters of the passport number
  sg  domestic group (from the "הערות הסעים" column, e.g. "6ב"), when the workbook has it
  df  domestic flight id from <db_dir>/ticketing/domestic.json, or "delhi" for olim already in Delhi
Birth dates, file numbers, parents, addresses, phones, notes and documents are never written out.
Birth dates and file numbers are read only to count infants per domestic flight and to keep families together.
Sheets without a date in their name (the summary, additions, "not arriving") are skipped.

The domestic flight of each person follows the same rule as the full board (tkDomApply):
already in Delhi by file -> "delhi"; otherwise the flight on the group's domestic day (Dimapur flight for
olim whose flight notes say Dimapur); where a group has two flights that day and the board's manual map
has no choice for the file, families are proposed in file order, whole files together, up to each flight's plan.
"""
import collections, datetime, json, os, re, sys
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def clean(v):
    return re.sub(r'\s+', ' ', str(v or '')).strip()


def col(h, *names):
    for n in names:
        for i, x in enumerate(h):
            if clean(x) == n:
                return i
    return None


def parse_date(v):
    """Workbook dates come as datetime or as text (d/m/Y with leading zero, m/d/Y without)."""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.date() if isinstance(v, datetime.datetime) else v
    t = clean(v)
    m = re.match(r'(\d{1,2})/+(\d{1,2})/+(\d{4})', t)
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    dm = t[0] == '0' or a > 12
    try:
        return datetime.date(y, b, a) if dm else datetime.date(y, a, b)
    except ValueError:
        try:
            return datetime.date(y, b, a)
        except ValueError:
            return None


def age_on(dob, iso):
    if not dob:
        return None
    d = datetime.date.fromisoformat(iso)
    return d.year - dob.year - ((d.month, d.day) < (dob.month, dob.day))


def assign_domestic(people, plan, sg_order):
    """people: dicts with f, sg, file, n, dom, dmu. Sets p['df'] in place. Mirrors tkDomApply in the board."""
    flights = sorted(plan.get('flights', []), key=lambda f: f['date'] + (f.get('dep') or '99:99') + ('1' if f.get('dmu') else '0'))
    by_id = {f['id']: f for f in flights}
    delhi, mp = plan.get('delhi', {}), plan.get('map', {})
    for p in people:
        p['df'], p['auto'], p['manual'], p['cand'] = '', False, False, None
        if p['file'] in delhi:
            p['df'] = 'delhi'
            continue
        cand = [f['id'] for f in flights if f['date'] == p['dom'] and p['sg'] in (f.get('sgs') or []) and bool(f.get('dmu')) == p['dmu']]
        if not cand:
            continue
        if len(cand) == 1:
            p['df'] = cand[0]
            continue
        p['cand'] = cand
        if mp.get(p['file']) in cand:
            p['df'], p['manual'] = mp[p['file']], True
    sets = collections.defaultdict(list)
    for p in people:
        if p['cand'] and not p['manual']:
            sets[','.join(p['cand'])].append(p)
    for k, L in sets.items():
        ids = k.split(',')
        used = {i: sum(1 for p in people if p['df'] == i) for i in ids}
        size = collections.Counter(p['sg'] for p in L)
        L = sorted(L, key=lambda p: (size[p['sg']], sg_order.get(p['sg'], 0), p['f'], p['n']))
        files, by_file = [], collections.OrderedDict()
        for p in L:
            if p['file'] not in by_file:
                by_file[p['file']] = []
                files.append(p['file'])
            by_file[p['file']].append(p)
        for fi in files:
            ms = by_file[fi]
            fid = next((i for i in ids if used[i] + len(ms) <= int(by_id[i].get('plan') or 0)), None)
            if fid is None:
                fid = sorted(ids, key=lambda i: -(int(by_id[i].get('plan') or 0) - used[i]))[0]
            for p in ms:
                p['df'], p['auto'] = fid, True
            used[fid] += len(ms)
        if len(ids) == 2:
            a, b = ids
            target = int(by_id[a].get('plan') or 0) - sum(1 for p in people if p['df'] == a and p['manual'])
            sz = lambda fi: len(by_file[fi])
            on = lambda i: [fi for fi in files if by_file[fi][0]['df'] == i]

            def move(fi, i):
                for p in by_file[fi]:
                    p['df'] = i
            for _ in range(30):
                diff = target - sum(sz(fi) for fi in on(a))
                if not diff:
                    break
                src = on(b) if diff > 0 else on(a)
                fi = next((x for x in reversed(src) if sz(x) == abs(diff)), None)
                if fi:
                    move(fi, a if diff > 0 else b)
                    continue
                sw = next(((x, y) for x in reversed(on(a)) for y in on(b) if sz(y) - sz(x) == diff), None)
                if sw:
                    move(sw[0], b)
                    move(sw[1], a)
                    continue
                break


def main():
    src, db = sys.argv[1], sys.argv[2]
    label = sys.argv[3] if len(sys.argv) > 3 else os.path.basename(src)
    with open(os.path.join(ROOT, 'data.json'), encoding='utf8') as f:
        olim = json.load(f).get('olim') or {}
    flights = olim.get('flights') or []
    subgroups = olim.get('subgroups') or []
    by_date = {fl['date'][5:]: fl['f'] for fl in flights}  # 'MM-DD' -> flight number
    sg_dom = {s['id']: s['dom'] for s in subgroups}
    sg_order = {s['id']: i for i, s in enumerate(subgroups)}
    plan_path = os.path.join(db, 'ticketing', 'domestic.json')
    plan = None
    if os.path.exists(plan_path):
        with open(plan_path, encoding='utf8') as f:
            plan = json.load(f)
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    out, seen, work = [], set(), []
    for ws in wb.worksheets:
        m = re.search(r'(\d{1,2})\.(\d{1,2})', ws.title)
        if not m:
            continue
        key = '%02d-%02d' % (int(m.group(2)), int(m.group(1)))
        if key not in by_date:
            sys.exit(f'sheet "{ws.title}": no flight on {key} in data.json')
        fno = by_date[key]
        rows = ws.iter_rows(values_only=True)
        h = list(next(rows))
        ip, ifam, igiv = col(h, 'Passport'), col(h, 'Family Name'), col(h, 'Given Name')
        ihf, ihg = col(h, 'שם משפחה'), col(h, 'שם בעברית')
        isg = col(h, 'הערות הסעים', 'הערת הסעים')
        ifile, inote, idob = col(h, 'תיק מקורי', 'תיק'), col(h, 'הערות לטיסה'), col(h, 'תאריך לידה')
        if None in (ip, ifam, igiv, ihf, ihg):
            sys.exit(f'sheet "{ws.title}": missing one of the name or passport columns')
        n = 0
        for r in rows:
            pp = clean(r[ip]).replace(' ', '').upper()
            if not pp:
                continue
            if pp in seen:
                sys.exit(f'sheet "{ws.title}": passport ending {pp[-4:]} appears twice')
            seen.add(pp)
            n += 1
            rec = {'f': fno, 'he': ' '.join(x for x in (clean(r[ihf]), clean(r[ihg])) if x),
                   'en': ' '.join(x for x in (clean(r[igiv]), clean(r[ifam])) if x), 'pp4': pp[-4:]}
            sg = clean(r[isg]).split('-')[0] if isg is not None else ''
            if sg:
                if sg not in sg_dom:
                    sys.exit(f'sheet "{ws.title}": group "{sg}" is not one of the board subgroups')
                rec['sg'] = sg
            out.append(rec)
            work.append({'rec': rec, 'f': fno, 'sg': sg, 'n': n, 'file': clean(r[ifile]).replace(' ', '') if ifile is not None else '',
                         'dom': sg_dom.get(sg, ''), 'dmu': 'דימאפור' in clean(r[inote]) if inote is not None else False,
                         'dob': parse_date(r[idob]) if idob is not None else None})
    if not out:
        sys.exit('no passengers found in ' + src)
    doc = {'source': label, 'people': out}
    if plan and all(w['sg'] for w in work):
        assign_domestic(work, plan, sg_order)
        by_id = {f['id']: f for f in plan.get('flights', [])}
        agg = collections.Counter()
        inf = collections.Counter()
        for w in work:
            w['rec']['df'] = w['df']
            k = (w['f'], w['sg'], w['df'], bool(w['auto']))
            agg[k] += 1
            fl = by_id.get(w['df'])
            a = age_on(w['dob'], fl['date']) if fl else None
            if a is not None and a < 2:
                inf[k] += 1
        doc['domestic'] = {
            'source': plan.get('source', ''),
            'flights': [{x: f.get(x) for x in ('id', 'date', 'no', 'dep', 'arr', 'from', 'sgs', 'plan', 'escort', 'dmu') if f.get(x) not in (None, '')}
                        for f in plan.get('flights', [])],
            # rows: flight to Israel, group, domestic flight, proposed (not yet assigned by name), people, infants
            'rows': [[k[0], k[1], k[2], k[3], n, inf[k]] for k, n in sorted(agg.items(), key=lambda x: str(x[0]))],
        }
    os.makedirs(db, exist_ok=True)
    with open(os.path.join(db, 'roster.json'), 'w', encoding='utf8') as f:
        json.dump(doc, f, ensure_ascii=False)
    print(len(out), 'people', {fl: sum(1 for p in out if p['f'] == fl) for fl in sorted({p['f'] for p in out})})
    if 'domestic' in doc:
        c = collections.Counter()
        for r in doc['domestic']['rows']:
            c[r[2]] += r[4]
        print('domestic', dict(sorted(c.items())))


if __name__ == '__main__':
    main()
