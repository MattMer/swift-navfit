"""Read-only PDF inspection; outputs evidence, never edits source files."""
from pathlib import Path
import hashlib, json
import subprocess
import re
from pypdf import PdfReader
from xml.etree import ElementTree as ET

root = Path(__file__).resolve().parents[1]
out = root / 'planning' / 'evidence'
out.mkdir(parents=True, exist_ok=True)
results = []
for path in sorted(root.glob('*.pdf')):
    print('Inspecting:', path.name, flush=True)
    reader = PdfReader(path)
    if reader.is_encrypted:
        reader.decrypt('')
    try:
        len(reader.pages)
    except Exception as exc:
        results.append({'file':path.name,'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
                        'pypdf_error':str(exc),'pdfinfo':subprocess.check_output(['pdfinfo',str(path)],text=True)})
        print('pypdf could not parse; recorded independent pdfinfo evidence.')
        continue
    info = {'file': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'pages': len(reader.pages), 'encrypted': reader.is_encrypted}
    form = reader.trailer['/Root'].get('/AcroForm')
    form = form.get_object() if form else {}
    info['acroform_keys'] = list(form.keys())
    info['field_count'] = len(reader.get_fields() or {})
    xfa = form.get('/XFA')
    packets = []
    if xfa:
        xfa = xfa.get_object()
        pairs = [(str(xfa[i]), xfa[i+1].get_object().get_data()) for i in range(0,len(xfa),2)] if isinstance(xfa,list) else [('whole',xfa.get_data())]
        for packet_index,(name,data) in enumerate(pairs):
            record = {'packet':name,'bytes':len(data)}
            try:
                tree = ET.fromstring(data)
                record['root'] = tree.tag
                record['named_nodes'] = [{'tag':n.tag.split('}')[-1], 'name':n.attrib['name']} for n in tree.iter() if 'name' in n.attrib]
            except ET.ParseError as exc:
                record['xml_error'] = str(exc)
            packets.append(record)
            safe_name = re.sub(r'[^A-Za-z0-9_-]', '_', name)
            (out / (path.stem + '.' + str(packet_index) + '.' + safe_name + '.xml')).write_bytes(data)
    info['xfa_packets'] = packets
    text = '\n'.join(f'\n--- PDF PAGE {i+1} ---\n' + (p.extract_text() or '') for i,p in enumerate(reader.pages))
    (out / (path.stem + '.txt')).write_text(text,encoding='utf-8')
    results.append(info)
(out / 'source-inventory.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
for item in results:
    print(json.dumps({k:v for k,v in item.items() if k != 'xfa_packets'}))
    print('XFA:', [(p['packet'],p['bytes'],len(p.get('named_nodes',[]))) for p in item.get('xfa_packets',[])])
