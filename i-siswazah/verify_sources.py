"""Independent CSV audit: does not import the production importer."""
import csv
import json
import re
from pathlib import Path

files = ['senarai_pelajar_Penyelia.csv','senarai_pelajar_Penyelia2.csv','senarai_pelajar_Penyelia3.csv']
supervisors, identities, versions, counts = set(), set(), {}, []
for filename in files:
    rows = list(csv.reader((Path(__file__).parent / filename).open(encoding='utf-8-sig',newline='')))
    groups = students = 0
    for row in rows[2:]:
        if row[0].strip().isdigit():
            name = ' '.join(row[1].upper().split())
            staff = re.search(r'K\d+', name)
            parent = staff[0] if staff else name
            supervisors.add(parent)
            groups += 1
        if row[3].strip():
            matric = row[5].strip().upper()
            identities.add(matric)
            versions.setdefault(matric,[]).append((filename,parent,row))
            students += 1
    counts.append([groups,students])
assert counts == [[33,95],[33,94],[33,86]], counts
assert len(supervisors) == 34 and len(identities) == 95
assert sum(x[1] for x in counts) - len(identities) == 180
assert len({v[1] for v in versions['P168677']}) == 2
assert versions['P168677'][-1][1] == 'K025867'
assert versions['P169098'][-1][2][4] == 'SARJANA'
assert versions['P119396'][0][2][9] == 'KONVO 2024'
assert versions['P86119'][0][2][7] == 'GRADUAN 2024'
assert len([m for m,v in versions.items() if all(x[0] != files[-1] for x in v)]) == 9
seed = json.loads((Path(__file__).parent/'penyelia-seed.js').read_text(encoding='utf-8').split('window.PENYELIA_SEED = ',1)[1].rstrip(';\n'))
records = {p['noPelajar']:p for g in seed['groups'] for p in g['pelajar']}
assert set(records) == identities
for matric, evidence in versions.items():
    p = records[matric]
    assert [v['values'] for v in p['sourceVersions']] == [v[2] for v in evidence]
    # All nonblank latest values, including blank-preserving fallback, checked independently.
    for field, col in [('statusTerkini',11),('semesterPengajian',10),('penyeliaBersama',6)]:
        expected = next((v[2][col].strip() if col == 6 else v[2][col] for v in reversed(evidence) if v[2][col].strip()),'')
        assert p.get(field,'') == expected, (matric,field)
    assert p['tarikhStatus'] == ''
assert {m for m,p in records.items() if p['statusKhas']=='graduasi'} == {'P119396','P86119'}
assert {m for m,p in records.items() if p['statusKhas']=='diberhentikan'} == {'P107661','P108824'}
assert records['P121441']['statusKhas'] == ''
assert records['P154207']['statusKhas'] == ''
print(json.dumps(dict(groups=len(supervisors),students=len(identities),perFile=counts,occurrences=275,duplicates=180)))
