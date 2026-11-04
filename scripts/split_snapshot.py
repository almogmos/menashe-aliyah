#!/usr/bin/env python3
"""Split the board's public snapshot (a JSON file from its file store) into the db layout build.py reads.

Usage: python3 scripts/split_snapshot.py <snapshot.json> <db_dir>

The snapshot holds only gantt (schedule, staff, responsibilities), welcomes and visa. It never holds
olim records; if it does, this script stops.
"""
import json, os, sys

src, db = sys.argv[1], sys.argv[2]
with open(src, encoding='utf8') as f:
    s = json.load(f)
extra = set(s) - {'hash', 'at', 'gantt', 'welcomes', 'visa'}
if extra:
    sys.exit('snapshot has unexpected sections: ' + ', '.join(sorted(extra)))
for k in ('gantt', 'welcomes', 'visa'):
    if s.get(k) is None:
        continue
    os.makedirs(os.path.join(db, k), exist_ok=True)
    with open(os.path.join(db, k, 'main.json'), 'w', encoding='utf8') as f:
        json.dump(s[k], f, ensure_ascii=False)
with open(os.path.join(db, 'snapshot.json'), 'w', encoding='utf8') as f:
    json.dump({'hash': s.get('hash', ''), 'at': s.get('at', '')}, f)
print('snapshot', s.get('hash'), s.get('at'))
