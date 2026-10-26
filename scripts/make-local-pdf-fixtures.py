"""Development-only PDF reader fixtures; pikepdf is an independent byte oracle.

The released Access application has no Python or pikepdf dependency.
"""
import argparse
import hashlib
import json
import re
import zlib
from pathlib import Path
import pikepdf

ROOT = Path(__file__).resolve().parents[1]
SOURCES = {
    'EVAL': 'E1-E6 EVAL BLANK NAVPERS 1616-26 05-2025_Final.pdf',
    'CHIEFEVAL': 'CHIEF EVAL BLANK 1616-27 05-2025_Final.pdf',
    'FITREP': 'W2-O6 FITREP NAVPERS 1610-2 05-2025_Final.pdf',
}

def digest(data):
    return hashlib.sha256(data).hexdigest().upper()

def classic(template, data, compression='dynamic', filter_name=b'FlateDecode'):
    def stream(raw):
        if compression == 'plain':
            return b'<< /Length %d >>\nstream\n' % len(raw) + raw + b'\nendstream'
        if compression == 'fixed':
            c = zlib.compressobj(strategy=zlib.Z_FIXED)
            packed = c.compress(raw) + c.flush()
        else:
            packed = zlib.compress(raw, 0 if compression == 'stored' else 6)
        return b'<< /Length %d /Filter /' % len(packed) + filter_name + b' >>\nstream\n' + packed + b'\nendstream'
    objects = [b'<< /Type /Catalog /Pages 2 0 R /AcroForm 3 0 R >>',
               b'<< /Type /Pages /Count 0 /Kids [] >>',
               b'<< /XFA [(template) 4 0 R (datasets) 5 0 R] >>',
               stream(template), stream(data)]
    out = b'%PDF-1.7\n'; offsets = [0]
    for n, obj in enumerate(objects, 1):
        offsets.append(len(out)); out += f'{n} 0 obj\n'.encode() + obj + b'\nendobj\n'
    xref = len(out)
    out += b'xref\n0 6\n0000000000 65535 f \n'
    out += b''.join(f'{x:010d} 00000 n \n'.encode() for x in offsets[1:])
    out += f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode()
    return out

def create(directory):
    directory.mkdir(parents=True, exist_ok=False)
    manifest = []
    templates = {}
    good = b'<xfa:datasets xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"><xfa:data><topmostSubform><Name>LOCAL, TEST</Name></topmostSubform></xfa:data></xfa:datasets>'
    def record(name, content, kind=None, data=None, rejection=None):
        path = directory / (name + '.pdf'); path.write_bytes(content)
        row = {'file': path.name, 'sourceHash': digest(content), 'kind': kind, 'rejection': rejection}
        if kind:
            # Independent extraction through qpdf/pikepdf, not the reader under test.
            with pikepdf.open(path) as pdf:
                arr = pdf.Root.AcroForm.XFA
                packets = {str(arr[i]): arr[i + 1].read_bytes() for i in range(0, len(arr), 2)}
                assert packets['template'] == templates[kind]
                assert packets['datasets'] == data
                row.update(templateHash=digest(packets['template']), dataHash=digest(data))
        manifest.append(row)
    for kind, filename in SOURCES.items():
        source = ROOT / 'FORMS_REFERENCES_TEMPLATES' / filename; before = digest(source.read_bytes())
        with pikepdf.open(source) as pdf:
            arr = pdf.Root.AcroForm.XFA
            packets = {str(arr[i]): arr[i + 1].read_bytes() for i in range(0, len(arr), 2)}
            templates[kind] = packets['template']
            record(kind.lower() + '-original', source.read_bytes(), kind, packets['datasets'])
        assert digest(source.read_bytes()) == before
    template = templates['EVAL']
    base = classic(template, good)
    unicode_data = good.replace(b'LOCAL, TEST', 'LOCAL, Jos\u00e9 \u03a9 \U0001f680 &amp;'.encode('utf-8'))
    record('unicode-utf8', classic(template, unicode_data), 'EVAL', unicode_data)
    bom_data = b'\xef\xbb\xbf' + unicode_data
    record('unicode-bom', classic(template, bom_data), 'EVAL', bom_data)
    for method in ['dynamic', 'fixed', 'stored', 'plain']:
        record('classic-' + method, classic(template, good, method), 'EVAL', good)
    def rewrite(name, **options):
        with pikepdf.open(ROOT / 'FORMS_REFERENCES_TEMPLATES' / SOURCES['EVAL'] if name == 'linear' else directory / 'classic-dynamic.pdf') as pdf:
            destination = directory / ('temp-' + name + '.pdf')
            pdf.save(destination, **options)
        result = destination.read_bytes(); destination.unlink(); return result
    record('object-stream', rewrite('objects', object_stream_mode=pikepdf.ObjectStreamMode.generate), 'EVAL', good)
    compressed = (directory / 'object-stream.pdf').read_bytes()
    stream_offset = int(re.findall(rb'startxref\s+(\d+)', compressed)[-1])
    with pikepdf.open(directory / 'object-stream.pdf') as pdf:
        root_id, root_gen = pdf.Root.objgen
        size = len(pdf.objects) + 1
    hybrid = compressed + f'xref\n0 1\n0000000000 65535 f \ntrailer\n<< /Size {size} /Root {root_id} {root_gen} R /XRefStm {stream_offset} >>\nstartxref\n{len(compressed)}\n%%EOF\n'.encode()
    record('hybrid-xref', hybrid, 'EVAL', good)
    for label in ['duplicate-xfa', 'no-xfa']:
        with pikepdf.open(directory / 'classic-dynamic.pdf') as pdf:
            if label == 'duplicate-xfa':
                pdf.Root.AcroForm.XFA.append(pikepdf.String('datasets'))
                pdf.Root.AcroForm.XFA.append(pdf.Root.AcroForm.XFA[3])
            else:
                del pdf.Root.AcroForm['/XFA']
            dest = directory / ('temp-' + label + '.pdf')
            pdf.save(dest)
        raw = dest.read_bytes(); dest.unlink()
        record(label, raw, rejection='duplicate XFA' if label == 'duplicate-xfa' else 'No supported XFA')

    with pikepdf.open(ROOT / 'FORMS_REFERENCES_TEMPLATES' / SOURCES['EVAL']) as pdf:
        a = pdf.Root.AcroForm.XFA
        original_data = {str(a[i]): a[i + 1].read_bytes() for i in range(0, len(a), 2)}['datasets']
    record('linearized', rewrite('linear', linearize=True), 'EVAL', original_data)
    for revision, aes in [(2, False), (3, False), (4, False), (4, True)]:
        record(f'security-r{revision}-aes{aes}', rewrite('security', encryption=pikepdf.Encryption(owner='synthetic-owner', user='', R=revision, aes=aes, metadata=aes)), 'EVAL', good)
    record('password-required', rewrite('password', encryption=pikepdf.Encryption(owner='owner', user='required', R=4)), rejection='password')
    record('extraction-forbidden', rewrite('forbidden', encryption=pikepdf.Encryption(owner='owner', user='', R=4, allow=pikepdf.Permissions(extract=False))), rejection='extraction')
    record('aes256-unsupported', rewrite('aes256', encryption=pikepdf.Encryption(owner='owner', user='', R=6)), rejection='encryption')
    prev = int(re.findall(rb'startxref\s+(\d+)', base)[-1])
    changed = good.replace(b'LOCAL, TEST', b'LATEST, REVISION')
    obj = b'5 0 obj\n<< /Length %d >>\nstream\n' % len(changed) + changed + b'\nendstream\nendobj\n'
    update = base + obj
    xref = len(update)
    update += f'xref\n5 1\n{len(base):010d} 00000 n \ntrailer\n<< /Size 6 /Root 1 0 R /Prev {prev} >>\nstartxref\n{xref}\n%%EOF\n'.encode()
    record('incremental-latest', update, 'EVAL', changed)
    freed = base + f'xref\n5 1\n0000000000 00001 f \ntrailer\n<< /Size 6 /Root 1 0 R /Prev {prev} >>\nstartxref\n{len(base)}\n%%EOF\n'.encode()
    record('incremental-freed', freed, rejection='freed')
    record('unknown-template', classic(template.replace(b'<', b' ', 1), good), rejection='revision')
    field = json.loads((ROOT / 'src/pdf-field-map.json').read_text())['EVAL'][0]['source']
    duplicate = f'<a><{field}>A</{field}><{field}>B</{field}></a>'.encode()
    record('duplicate-fields', classic(template, duplicate), rejection='Repeated PDF fields')
    record('invalid-xml', classic(template, b'<broken>'), rejection='structured PDF data')
    record('external-dtd', classic(template, b'<!DOCTYPE a SYSTEM "https://invalid.example/never-fetch"><a/>'), rejection='structured PDF data')
    record('invalid-utf8', classic(template, b'<a>\xc0\xaf</a>'), rejection='UTF-8')
    record('oversized-data', classic(template, b'x' * 4000001), rejection='size limit')
    record('unsupported-filter', classic(template, good, filter_name=b'LZWDecode'), rejection='compression')
    record('truncated', base[:-20], rejection='PDF')
    record('trailing-data', base + b'junk', rejection='final PDF')
    corrupt = bytearray(base); start = corrupt.find(b'stream\n') + 7
    corrupt[start + 15] ^= 1
    record('damaged-deflate', bytes(corrupt), rejection='compressed PDF')
    record('not-pdf', b'This is not a PDF file.', rejection='PDF')
    (directory / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(f'PASS: {len(manifest)} independent local-reader fixtures')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(); parser.add_argument('--output', type=Path, required=True)
    create(parser.parse_args().output.resolve())
