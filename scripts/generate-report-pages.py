"""Build fixed report-page geometry; no PDF runtime dependency in the app/build."""
import json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
fields=json.loads((ROOT/'prototype/fields.mjs').read_text(encoding='utf-8-sig').split('=',1)[1].strip().rstrip(';'))
evidence=json.loads((ROOT/'build/report-page-layout.json').read_text()) if (ROOT/'build/report-page-layout.json').exists() else {}
verified=json.loads((ROOT/'src/report-page-map.json').read_text()) if (ROOT/'src/report-page-map.json').exists() else {}
mapping={}; layouts={}
bands=[['OfficialName','RankRate','Designator','MemberSSN'],['ComponentActive','ComponentTAR','ComponentInactive','ComponentATADSW','UnitUIC'],['ShipStation','PromotionStatus','DateReported'],['OccasionPeriodic','OccasionDetachment','OccasionFrocking','OccasionSpecial'],['PeriodStart','PeriodEnd','ObservationStatus','ReportRegular','ReportConcurrent','ReportOperationalCommander'],['ReadinessSummary','ReadinessDate','BilletSubcategory'],['ReportingSeniorName','ReportingSeniorGrade','ReportingSeniorDesignator','ReportingSeniorTitle','ReportingSeniorUIC','ReportingSeniorSSN'],['CommandEmployment'],['PrimaryDuty','DutyDescription'],['CounselingDate','CounselorName']]
for kind in ('EVAL','CHIEFEVAL','FITREP'):
    resolved={key:item['page'] for key,item in verified.get(kind,{}).items() if item['basis']=='Current NAVPERS XFA page subform'}
    for row in evidence.get(kind,{}).get('fields',[]):
        for loc in row['locations']:
            pages=[name.upper() for name,_ in loc['path'] if name]
            if 'PAGE1' in pages: resolved[row['key']]=0;break
            if 'PAGE2' in pages: resolved[row['key']]=1;break
    visible=[f for f in fields if not f.get('legacyHidden') and f['key'] not in ('FormKind','Title','MemberName','DraftNotes')]
    by_key={f['key']:f for f in visible}
    def block(f):
        if f['key']=='Designator': return 3
        if f['key']=='OfficialComments': return {'EVAL':43,'CHIEFEVAL':40,'FITREP':41}[kind]
        value=f.get('blocks',{}).get(kind,f.get('block',''))
        m=re.search(r'\d+',value)
        return int(m.group()) if m else 999
    boundary=37 if kind=='CHIEFEVAL' else 38
    mapping[kind]={f['key']:{'page':resolved.get(f['key'],int(block(f)>=boundary)), 'block':block(f),'basis':'Current NAVPERS XFA page subform' if f['key'] in resolved else 'Official block order'} for f in visible}
    #User-requested editing convenience: keep all seven traits together on Front.
    #This is an input workspace grouping, not a change to official printed pages.
    for f in visible:
        if f['key'].startswith('Trait'):
            mapping[kind][f['key']]['page']=0
            mapping[kind][f['key']]['basis']='Front Page editing group requested by user'
    for page in (0,1):
        rows=[];done=set();y=180;page_keys={k for k,v in mapping[kind].items() if v['page']==page}
        ordered=[]
        for band in bands:
            keys=[k for k in band if k in page_keys]
            if keys: ordered.append(keys);done.update(keys)
        traits=sorted(k for k in page_keys-done if k.startswith('Trait'))
        if traits: ordered.append(traits);done.update(traits)
        rest=sorted(page_keys-done,key=lambda k:(mapping[kind][k]['block'],list(by_key).index(k)))
        for key in rest:
            if ordered and key!='OfficialComments' and not by_key[key].get('multiline') and len(ordered[-1])<3 and not any(by_key[k].get('multiline') or k.startswith('Trait') for k in ordered[-1]): ordered[-1].append(key)
            else: ordered.append([key])
        for band in ordered:
            if band[0].startswith('Trait'):
                start=y
                for key in band:
                    rows.append(dict(key=key,x=180,y=y,w=5000,h=360,trait=True));y+=620
                rows.append(dict(key='@guidance',x=5550,y=start,w=6000,h=y-start-80));y+=200
                continue
            widths=[(11600-160*(len(band)-1))//len(band)]*len(band)
            if band[0]=='OfficialName': widths=[5400,1650,1650,2420]
            if band[0]=='ShipStation': widths=[6100,2500,2680]
            if band[0]=='PeriodStart': widths=[1900,1900,2400,1500,1600,1500]
            x=180;height=0
            for key,width in zip(band,widths):
                f=by_key[key];h=360
                if f.get('multiline'): h=950
                if key=='OfficialComments': h=1836 if kind=='CHIEFEVAL' else 3924
                rows.append(dict(key=key,x=x,y=y,w=width,h=h,trait=False));x+=width+160
                height=max(height,h+520 if not f.get('checkbox') else 600)
                if key=='OfficialComments': height=max(height,h+1500)
            y+=height+50
        layouts[f'{kind}:{page}']={'height':y+200,'fields':rows}
(ROOT/'src/report-page-map.json').write_text(json.dumps(mapping,indent=2)+'\n')
(ROOT/'src/report-page-layout.json').write_text(json.dumps(layouts,indent=2)+'\n')
lines=['Attribute VB_Name = "modReportPageMap"','Option Explicit','Public Function ReportPageForField(ByVal kind As String, ByVal key As String) As Long','Select Case kind & ":" & key']
for kind,items in mapping.items():
    for key,item in items.items(): lines.append(f'Case "{kind}:{key}": ReportPageForField = {item["page"]}')
lines+=['End Select','End Function','Public Function ReportPageRows(ByVal kind As String, ByVal page As Long) As Collection','Dim result As New Collection','Select Case kind & ":" & page']
for key,layout in layouts.items():
    lines.append(f'Case "{key}"')
    for row in layout['fields']: lines.append('result.Add "'+'|'.join(str(row[k]) for k in ('key','x','y','w','h'))+'"')
lines+=['End Select','Set ReportPageRows = result','End Function']
(ROOT/'src/modReportPageMap.bas').write_text('\n'.join(lines)+'\n')
print('Generated fixed front/back layouts and field-page maps for all three report types')
