#!/usr/bin/env python3
"""Turn the flights workbook (the per-flight sheets) into a minimal roster for the public page.

Usage: python3 scripts/import_flights.py <workbook.xlsx> <db_dir> [label]

Writes <db_dir>/roster.json. Only these fields leave the workbook, per person:
  f   flight number (matched by the date in the sheet name, e.g. "17.11")
  he  full name in Hebrew (family name + given name)
  en  full name in English (given name + family name)
  pp4 last 4 characters of the passport number
Birth dates, file numbers, parents, addresses, phones, notes and documents are never read out.
Sheets without a date in their name (the summary, additions, "not arriving") are skipped.
"""
import json, os, re, sys
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


def main():
    src, db = sys.argv[1], sys.argv[2]
    label = sys.argv[3] if len(sys.argv) > 3 else os.path.basename(src)
    with open(os.path.join(ROOT, 'data.json'), encoding='utf8') as f:
        flights = (json.load(f).get('olim') or {}).get('flights') or []
    by_date = {fl['date'][5:]: fl['f'] for fl in flights}  # 'MM-DD' -> flight number
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    out, seen = [], set()
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
        if None in (ip, ifam, igiv, ihf, ihg):
            sys.exit(f'sheet "{ws.title}": missing one of the name or passport columns')
        for r in rows:
            pp = clean(r[ip]).replace(' ', '').upper()
            if not pp:
                continue
            if pp in seen:
                sys.exit(f'sheet "{ws.title}": passport ending {pp[-4:]} appears twice')
            seen.add(pp)
            out.append({'f': fno, 'he': ' '.join(x for x in (clean(r[ihf]), clean(r[ihg])) if x),
                        'en': ' '.join(x for x in (clean(r[igiv]), clean(r[ifam])) if x), 'pp4': pp[-4:]})
    if not out:
        sys.exit('no passengers found in ' + src)
    os.makedirs(db, exist_ok=True)
    with open(os.path.join(db, 'roster.json'), 'w', encoding='utf8') as f:
        json.dump({'source': label, 'people': out}, f, ensure_ascii=False)
    print(len(out), 'people', {fl: sum(1 for p in out if p['f'] == fl) for fl in sorted({p['f'] for p in out})})


if __name__ == '__main__':
    main()
