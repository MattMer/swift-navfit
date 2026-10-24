"""Internal synthetic XFA import fixtures, not printed-layout examples."""
import argparse
import json
import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path
import pikepdf
from make_pdf_sources import SOURCES

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output/pdf/r85"

def create(output=None):
    global OUT
    if output is not None: OUT = Path(output).resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    mapping = json.loads((ROOT / "src/pdf-field-map.json").read_text())
    manifest = []
    for kind, filename in SOURCES.items():
        source = ROOT / 'FORMS_REFERENCES_TEMPLATES' / filename
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        with pikepdf.open(source) as pdf:
            xfa = pdf.Root["/AcroForm"]["/XFA"]
            packets = {str(xfa[i]): (i+1, xfa[i+1].read_bytes()) for i in range(0,len(xfa),2)}
            dataset = ET.fromstring(packets['datasets'][1])
            container = dataset.find('.//form1' if kind == 'CHIEFEVAL' else './/topmostSubform')
            values = {
                'CommandEmployment': 'SYNTHETIC BLOCK 28: Training and command achievements.',
                'OfficialName': 'EXAMPLE, TEST', 'RankRate': {'EVAL':'YN2','CHIEFEVAL':'YNC','FITREP':'LCDR'}[kind],
                'Designator': '1110' if kind == 'FITREP' else 'SW', 'MemberSSN': '000000007',
                'UnitUIC':'12345', 'ShipStation':'SYNTHETIC UNIT', 'ReportingSeniorName':'EXAMPLE, A B',
                'ReportingSeniorGrade':'CAPT', 'ReportingSeniorDesignator':'1110', 'ReportingSeniorTitle':'CO',
                'ReportingSeniorUIC':'54321', 'ReportingSeniorSSN':'000000009', 'PrimaryDuty':'ADMIN',
                'DutyDescription':"O'Brien & synthetic duties.", 'CounselorName':'EXAMPLE, C D',
                'AssignmentRecommendationA':'LPO', 'AssignmentRecommendationB':'INSTRUCTOR',
                'OfficialComments':"Expanded synthetic comments. O'Brien & example.",
                'BilletSubcategory':'SPECIAL01', 'PromotionStatus':'REGULAR',
                'PromotionRecommendation':'Must Promote', 'ReportingSeniorAddress1':'123 TEST ROAD',
                'ReportingSeniorAddress2':'SUITE 2', 'ReportingSeniorPhone':'0000000000', 'ReportingSeniorDSN':'0000000',
                'RaterName':'EXAMPLE, R', 'SeniorRaterName':'EXAMPLE, S'}
            for block in range(33, 40):
                for entry in range(1, 4):
                    values[f'Trait{block}Comment{entry}'] = f'SYNTHETIC PDF COMMENT {block}.{entry}'
            raw = {}
            for row in mapping[kind]:
                target, field, conversion = row['target'], row['source'], row['conversion']
                if conversion in ('date','date-text'): value='260101' if target != 'PeriodEnd' else '260930'
                elif conversion == 'component': value='act'
                elif conversion in ('trait','chief-trait'): value=str(3+(int(target[-2:])%3))
                elif conversion == 'flag': value='1' if target in ['ComponentActive','OccasionPeriodic','ReportRegular','StatementNo'] else '0'
                elif conversion == 'observation': value='0'
                elif conversion == 'recommendation': value='mp'
                elif conversion == 'promotion-status': value='1'
                else: value=values[target]
                raw[field] = value
            for field, value in raw.items():
                node = container.find('.//'+field)
                if node is None: node=ET.SubElement(container,field)
                node.text=value
            for suffix in ('expanded','merge'):
                target=OUT/f'{kind.lower()}-{suffix}.pdf'
                if target.exists(): raise FileExistsError(target)
                comments=container.find('.//'+next(r['source'] for r in mapping[kind] if r['target']=='OfficialComments'))
                comments.text=values['OfficialComments'] if suffix=='expanded' else 'MERGED SYNTHETIC COMMENTS'
                xfa[packets['datasets'][0]]=pdf.make_stream(ET.tostring(dataset,encoding='utf-8'))
                pdf.save(target,encryption=True if pdf.is_encrypted else None)
                with pikepdf.open(target) as written:
                    arr=written.Root['/AcroForm']['/XFA']
                    payload=next(arr[i+1].read_bytes() for i in range(0,len(arr),2) if str(arr[i])=='datasets')
                    assert ET.fromstring(payload).find('.//'+list(raw)[0]).text==raw[list(raw)[0]]
                manifest.append({'file':target.name,'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),'kind':kind,'source_sha256':before,'raw':{field:container.find('.//'+field).text for field in raw}})
        assert hashlib.sha256(source.read_bytes()).hexdigest()==before
    (OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
    print('PASS: six synthetic interactive XFA fixtures; blank sources unchanged')

if __name__=='__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output')
    create(parser.parse_args().output)
