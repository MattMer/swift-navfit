"""Create synthetic counseling-status fixtures without changing the source PDFs.

--source is a folder made by make-pdf-expanded-fixtures.py.
"""
import argparse,json,xml.etree.ElementTree as ET
from pathlib import Path
import pikepdf
ROOT=Path(__file__).resolve().parents[1]
parser=argparse.ArgumentParser();parser.add_argument('--source',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
args=parser.parse_args();SOURCE=args.source.resolve();OUT=args.output.resolve();OUT.mkdir(parents=True,exist_ok=True)
mapping=json.loads((ROOT/'src/pdf-field-map.json').read_text())
for kind in ('EVAL','CHIEFEVAL','FITREP'):
 for status in ('NOT REQ','NOT PERF'):
  target=OUT/(kind.lower()+'-'+status.replace(' ','-')+'.pdf')
  if target.exists():raise FileExistsError(target)
  with pikepdf.open(SOURCE/(kind.lower()+'-expanded.pdf')) as pdf:
   xfa=pdf.Root['/AcroForm']['/XFA'];idx=next(i+1 for i in range(0,len(xfa),2) if str(xfa[i])=='datasets')
   dataset=ET.fromstring(xfa[idx].read_bytes());field=next(r['source'] for r in mapping[kind] if r['target']=='CounselingDate')
   dataset.find('.//'+field).text=status
   xfa[idx]=pdf.make_stream(ET.tostring(dataset,encoding='utf-8'));pdf.save(target,encryption=True if pdf.is_encrypted else None)
print('PASS: six synthetic counseling status PDFs created')
