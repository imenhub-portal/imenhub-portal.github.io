"""Lossless three-source union. python import_penyelia.py (stdlib only)."""
import csv
import hashlib
import io
import json
import re
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FILES = ['senarai_pelajar_Penyelia.csv', 'senarai_pelajar_Penyelia2.csv', 'senarai_pelajar_Penyelia3.csv']
SESSIONS = ['2/2022-2023', '1/2023-2024', '2/2023-2024', '2/2025-2026']

def supervisor_key(name):
    staff = re.search(r'\(K\d+\)', name, re.I)
    return staff[0].upper() if staff else ' '.join(name.upper().split())

def classify(text):
    """Only affirmative final states; pending/future/unregistered is not final."""
    t = ' '.join(text.upper().split())
    if re.fullmatch(r'(GRADUAN|KONVO)\s+\d{4}', t) or re.fullmatch(r'TELAH BERGRADUASI(?:\s+\d{4})?', t):
        return 'graduasi'
    if t.startswith('DIBERHENTIKAN'):
        return 'diberhentikan'
    if re.fullmatch(r'(?:TELAH )?MENARIK DIRI(?: DARIPADA PENGAJIAN)?', t):
        return 'menarik_diri'
    if t == 'AKTIF SEMULA':
        return 'aktif_semula'
    return ''

def build():
    groups, students, sources = {}, {}, []
    digest = hashlib.sha256(b'three-source-union-v1')
    for source_index, filename in enumerate(FILES):
        raw = (ROOT / filename).read_bytes()
        digest.update(filename.encode() + raw)
        rows = list(csv.reader(io.StringIO(raw.decode('utf-8-sig'), newline='')))
        header = rows[1]
        assert header[:7] == ['BIL','NAMA PENYELIA','PELAJAR','','PROGRAM','NO. PELAJAR','PENYELIA BERSAMA'], filename
        assert len(header) == (15 if source_index == 2 else 12), filename
        assert 'STATUS TERKINI' in header[11]
        source = dict(filename=filename, sha256=hashlib.sha256(raw).hexdigest(),
                      reviewDate='2026-06-18' if source_index == 2 else '', headers=header, rows=rows,
                      supervisors=0, students=0, emptySupervisors=0)
        sources.append(source)
        parent = None
        counts = Counter()
        for row_number, row in enumerate(rows[2:], 3):
            assert len(row) == len(header), (filename, row_number, len(row))
            if row[0].strip().isdigit() and row[1].strip():
                key = supervisor_key(row[1])
                if key not in groups:
                    # Retain the original first-source IDs for browser compatibility.
                    gid = 'csv-supervisor-' + row[0].strip() if source_index == 0 else 'csv-supervisor-' + hashlib.sha256(key.encode()).hexdigest()[:12]
                    groups[key] = dict(id=gid, bil=len(groups)+1, namaPenyelia=row[1].strip(), pelajar=[], sourceRows=[])
                parent = groups[key]
                counts[key] += 0
                source['supervisors'] += 1
            provenance = dict(source=filename, row=row_number, values=row, supervisorId=parent['id'] if parent else None)
            if parent:
                parent['sourceRows'].append(provenance)
            if not row[3].strip():
                continue
            assert parent, (filename, row_number, 'student without parent')
            matric = re.sub(r'\s+', '', row[5]).upper()
            assert re.fullmatch(r'P\d+', matric), (filename, row_number, 'unreliable matric', matric)
            source['students'] += 1
            counts[key] += 1
            data = dict(nama=' '.join(row[3].split()), noPelajar=matric,
                        program={'SARJANA':'Sarjana Sains Kejuruteraan Mikro dan Nanoelektronik', 'KEDOKTORAN':'Doktor Falsafah'}[row[4].strip()],
                        penyeliaBersama=row[6].strip(), semesterPengajian=row[10], statusTerkini=row[11], parentId=parent['id'])
            if matric not in students:
                students[matric] = dict(id='csv-student-'+matric, tarikhStatus='', statusKhas='',
                    sesiSemesterPengajian=SESSIONS[-1], sejarahSemester=[dict(sesi=s,semesterPengajian='',catatan='') for s in SESSIONS],
                    sourceVersions=[], conflicts=[], statusEvidence=[])
            p = students[matric]
            p['sourceVersions'].append(dict(**provenance, fields=data))
            for field, value in data.items():
                if str(value).strip():
                    if p.get(field) and p[field] != value:
                        p['conflicts'].append(dict(field=field, previous=p[field], incoming=value, source=filename, row=row_number))
                    p[field] = value
            for index, session in enumerate(SESSIONS):
                value = row[7+index]
                snap = p['sejarahSemester'][index]
                if value.strip():
                    if snap['semesterPengajian'] and snap['semesterPengajian'] != value:
                        p['conflicts'].append(dict(field='semester:'+session, previous=snap['semesterPengajian'], incoming=value, source=filename, row=row_number))
                    snap.update(semesterPengajian=value, sourceValue=value, source=filename)
            for column in range(7,12):
                status = classify(row[column])
                if status:
                    evidence = dict(source=filename,row=row_number,column=header[column],text=row[column],status=status)
                    p['statusEvidence'].append(evidence)
                    p['statusKhas'] = '' if status == 'aktif_semula' else status
                    year = re.search(r'\b(20\d{2})\b', row[column])
                    if status == 'graduasi' and year:
                        p['tahunGraduasi'] = year[1]
        source['emptySupervisors'] = sum(n == 0 for n in counts.values())
    by_id = {g['id']:g for g in groups.values()}
    for p in students.values():
        p['needsDateReview'] = p['statusKhas'] == 'graduasi' and not p['tarikhStatus']
        p['missingFromLatest'] = not any(v['source'] == FILES[-1] for v in p['sourceVersions'])
        parent = by_id[p['parentId']]
        p['bil'] = len(parent['pelajar'])+1
        parent['pelajar'].append(p)
    names = {}
    for matric,p in students.items():
        for version in p['sourceVersions']:
            name = version['fields']['nama'].upper()
            names.setdefault(name,set()).add(matric)
    name_conflicts = [dict(name=n,matrics=sorted(ids)) for n,ids in names.items() if len(ids)>1]
    distribution = Counter(p['statusKhas'] or 'aktif' for p in students.values())
    report = dict(sources=[{k:s[k] for k in ['filename','supervisors','students','emptySupervisors']} for s in sources],
        supervisors=len(groups),students=len(students),studentOccurrences=sum(s['students'] for s in sources),
        duplicateOccurrences=sum(s['students'] for s in sources)-len(students),
        conflictingStudents=sum(bool(p['conflicts']) for p in students.values()),
        fieldConflicts=sum(len(p['conflicts']) for p in students.values()),
        transfers=sum(any(c['field']=='parentId' for c in p['conflicts']) for p in students.values()),
        missingFromLatest=sum(p['missingFromLatest'] for p in students.values()),
        emptySupervisors=sum(not g['pelajar'] for g in groups.values()), statuses=dict(distribution),
        needsDateReview=sum(p['needsDateReview'] for p in students.values()),nameIdentityConflicts=name_conflicts)
    return dict(version=digest.hexdigest(), sources=sources, report=report, groups=list(groups.values()))

if __name__ == '__main__':
    payload = build()
    (ROOT/'penyelia-seed.js').write_text('/* Generated by import_penyelia.py; do not edit. */\nwindow.PENYELIA_SEED = '+json.dumps(payload,ensure_ascii=False,indent=2)+';\n',encoding='utf-8')
    print(json.dumps(payload['report'],ensure_ascii=False,indent=2))
