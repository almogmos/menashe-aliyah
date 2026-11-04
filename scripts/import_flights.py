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
People listed under "removed" in the domestic plan (file + given name) are left out. Files under "fixed" go to that
domestic flight (e.g. Kangpokpi olim flying from Imphal).

The domestic flight of each person follows the same rule as the full board (tkDomApply):
already in Delhi by file -> "delhi"; otherwise the flight on the group's domestic day (Dimapur flight for
olim whose flight notes say Dimapur); where a group has two flights that day and the board's manual map
has no choice for the file, the person waits on a "pend:<date>" entry for that day (no automatic split).
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
    flights = sorted(plan.get('flights', []), key=lambda f: (f.get('date') or '9999') + (f.get('dep') or '99:99') + ('1' if f.get('dmu') else '0'))
    by_id = {f['id']: f for f in flights}
    delhi, mp, fixed = plan.get('delhi', {}), plan.get('map', {}), plan.get('fixed', {})
    for p in people:
        p['df'], p['auto'], p['manual'], p['cand'] = '', False, False, None
        if p['file'] in delhi:
            p['df'] = 'delhi'
            continue
        if fixed.get(p['file']) in by_id:
            p['df'] = fixed[p['file']]
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
    # No automatic split: a group with two flights on its day and no manual choice waits on a "pending" entry for that day.
    pend = {}
    for p in people:
        if not p['cand'] or p['manual']:
            continue
        c0 = by_id[p['cand'][0]]
        pid = 'pend:' + c0['date'] + (':dmu' if c0.get('dmu') else '')
        if pid not in pend:
            cs = [by_id[x] for x in p['cand']]
            pend[pid] = {'id': pid, 'pend': True, 'date': c0['date'], 'from': c0.get('from', ''), 'cands': list(p['cand']),
                         'no': ' ו-'.join(f['no'] for f in cs if f.get('no')), 'deps': [f['dep'] for f in cs if f.get('dep')],
                         'plan': sum(int(f.get('plan') or 0) for f in cs),
                         'sgs': sorted({s for f in cs for s in (f.get('sgs') or [])}),
                         'escort': ', '.join(f['escort'] for f in cs if f.get('escort'))}
        p['df'] = pid
    used = {p['df'] for p in people}
    hidden = {x for v in pend.values() for x in v['cands'] if x not in used}
    return [f for f in flights if f['id'] not in hidden] + list(pend.values())


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
            fileno = clean(r[ifile]).replace(' ', '') if ifile is not None else ''
            if plan and any(x.get('file') == fileno and x.get('given', '').lower() == clean(r[igiv]).lower() for x in plan.get('removed', [])):
                seen.discard(pp)
                continue
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
        flights_out = assign_domestic(work, plan, sg_order)
        by_id = {f['id']: f for f in flights_out}
        agg = collections.Counter()
        inf = collections.Counter()
        for w in work:
            w['rec']['df'] = w['df']
            k = (w['f'], w['sg'], w['df'], bool(w['auto']))
            agg[k] += 1
            fl = by_id.get(w['df'])
            a = age_on(w['dob'], (fl.get('date') or w['dom'])) if fl else None
            if a is not None and a < 2:
                inf[k] += 1
        doc['domestic'] = {
            'source': plan.get('source', ''),
            'flights': [{x: f.get(x) for x in ('id', 'date', 'no', 'dep', 'deps', 'arr', 'from', 'sgs', 'plan', 'escort', 'dmu', 'pend', 'cands') if f.get(x) not in (None, '', [])}
                        for f in flights_out],
            # rows: flight to Israel, group, domestic flight (pend:<date> = two flights that day, not yet assigned by name), unused, people, infants
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
