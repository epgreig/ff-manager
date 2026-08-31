"""Minimal xlsx reader: no dependencies, returns a sheet as rows of strings."""
import zipfile, re, xml.etree.ElementTree as ET
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NSR='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def _col(ref):
    m=re.match(r'([A-Z]+)', ref); n=0
    for c in m.group(1): n=n*26+ord(c)-64
    return n-1

def sheets(path):
    z=zipfile.ZipFile(path)
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels={r.get('Id'):r.get('Target') for r in
          ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out=[]
    for s in wb.find(NS+'sheets'):
        t=rels[s.get(NSR+'id')]
        out.append((s.get('name'), 'xl/'+t.lstrip('/').replace('worksheets/','worksheets/')))
    return out

def read(path, name, max_rows=100000):
    z=zipfile.ZipFile(path)
    strings=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')):
            strings.append(''.join(t.text or '' for t in si.iter(NS+'t')))
    target=dict(sheets(path))[name]
    if target not in z.namelist(): target='xl/'+target.split('xl/')[-1]
    rows=[]
    for row in ET.fromstring(z.read(target)).iter(NS+'row'):
        cells={}
        for c in row.iter(NS+'c'):
            v=c.find(NS+'v'); t=c.get('t')
            if v is None:
                isel=c.find(NS+'is')
                val=''.join(x.text or '' for x in isel.iter(NS+'t')) if isel is not None else ''
            elif t=='s': val=strings[int(v.text)]
            else: val=v.text or ''
            cells[_col(c.get('r'))]=val
        if cells:
            rows.append([cells.get(i,'') for i in range(max(cells)+1)])
        if len(rows)>=max_rows: break
    return rows
