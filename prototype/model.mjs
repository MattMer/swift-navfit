import {fields} from './fields.mjs';
export const sectionNames=['Member & report details','Reporting senior','Command & duties','Counseling','Performance traits','Recommendations & comments','Final report details','Draft notes','Review','History'];
export const clone=x=>JSON.parse(JSON.stringify(x));
export function initialState(){
 const base=Object.fromEntries(fields.map(f=>[f.key,'']));
 const report=(id,name,kind,grade,state='Draft')=>({id,revision:1,state,archived:false,folder:'Periodic 2026',values:{...base,FormKind:kind,Title:`${grade} periodic 2026`,MemberName:name,OfficialName:name.toUpperCase(),RankRate:grade,UnitUIC:'00000',ShipStation:'SYNTHETIC UNIT',PeriodStart:'2026-01-01',PeriodEnd:'2026-09-30',PromotionStatus:'REGULAR',ObservationStatus:'Observed',ReportingSeniorName:'MORGAN A',ReportingSeniorGrade:'CDR',ReportingSeniorTitle:'COMMANDING OFF',ReportingSeniorUIC:'00000',PromotionRecommendation:'Promotable',Trait33:'3.0',Trait34:'3.0',Trait35:'3.0',Trait36:'3.0',Trait37:'3.0',Trait38:'3.0',Trait39:'3.0',OfficialComments:'SYNTHETIC EXAMPLE ONLY. Demonstrated consistent performance and contributed to command readiness.'},history:[{revision:1,at:'Demo starting revision',reason:'Created synthetic example'}],comments:[]});
 const reports=[report(1,'Taylor, Jamie','EVAL','BM2'),report(2,'Lee, Jordan','CHIEFEVAL','BMC'),report(3,'Reed, Casey','FITREP','LT'),report(4,'Patel, Sam','EVAL','BM2','Approved')];
 return {reports,groups:[{id:1,title:'E5 periodic · September 2026',kind:'EVAL',grade:'E5',members:[1,4],snapshot:null,defaults:{ShipStation:'SYNTHETIC UNIT'},history:['Created synthetic group']},{id:2,title:'E7 periodic · September 2026',kind:'CHIEFEVAL',grade:'E7',members:[2],snapshot:null,defaults:{},history:['Created synthetic group']},{id:3,title:'O3 periodic · September 2026',kind:'FITREP',grade:'O3',members:[3],snapshot:null,defaults:{},history:['Created synthetic group']}],routes:[{reportId:2,revision:1,steps:[{name:'Alex Morgan',state:'Pending'},{name:'Robin Chen',state:'Waiting'}],note:''},{reportId:3,revision:0,steps:[{name:'Alex Morgan',state:'Pending'}],note:''}],recovery:{reportId:1,values:{...reports[0].values,DraftNotes:'Synthetic recovered pending notes. Not a saved report revision.'}},profile:'Alex Morgan',operationHistory:[],batches:[]};
}
export function isDirty(saved,pending){return fields.some(f=>saved[f.key]!==pending[f.key]);}
export function validate(values){const errors={};for(const f of fields){const v=values[f.key]??'';if(f.limit&&v.length>f.limit) errors[f.key]=`Use ${f.limit} characters or fewer. Your entry has ${v.length}.`;if(f.date&&v&&!validDate(v))errors[f.key]='Enter a real date as YYYY-MM-DD.';if(f.digits&&v&&!new RegExp(`^\\d{${f.digits}}$`).test(v))errors[f.key]=`Enter ${f.digits} digits, preserving leading zeros.`;if(f.options&&v&&!f.options.includes(v))errors[f.key]='Choose an available value or leave blank.';}if(values.PeriodStart&&values.PeriodEnd&&validDate(values.PeriodStart)&&validDate(values.PeriodEnd)&&values.PeriodEnd<values.PeriodStart)errors.PeriodEnd='The end date must be on or after the start date.';return errors;}
export function validDate(v){if(!/^\d{4}-\d{2}-\d{2}$/.test(v))return false;const d=new Date(v+'T12:00:00Z');return !isNaN(d)&&d.toISOString().slice(0,10)===v;}
export function saveReport(state,id,pending,reason='Saved draft'){const r=state.reports.find(r=>r.id===id);if(!r)throw Error('Report not found.');if(r.state==='Approved'||r.archived)throw Error('Reopen this report before editing.');const errors=validate(pending);if(Object.keys(errors).length)return {errors};if(!isDirty(r.values,pending))return {changed:false};r.values=clone(pending);r.revision++;r.history.unshift({revision:r.revision,at:new Date().toLocaleString(),reason});if(state.recovery?.reportId===id)state.recovery=null;return {changed:true};}
export function routeState(state,r){if(r.archived)return 'Archived';if(r.state==='Approved')return 'Local approval';const route=state.routes.find(x=>x.reportId===r.id);if(!route)return 'Draft';if(route.revision!==r.revision)return 'Review needs restart';if(route.steps.some(s=>s.state==='Returned'))return 'Returned';if(route.steps.every(s=>s.state==='Accepted'))return 'Reviews complete';return 'In review';}
export function previewBatch(state,ids,key,value){if(key==='FormKind'||!fields.some(f=>f.key===key))throw Error('Choose an editable field.');if(!ids.length)throw Error('Select reports first.');return ids.map(id=>{const r=state.reports.find(r=>r.id===id);if(!r||r.archived||r.state==='Approved')throw Error('Reopen approved or archived reports before applying a batch.');const after={...r.values,[key]:value};const errors=validate(after);if(Object.keys(errors).length)throw Error(Object.values(errors)[0]);return {id,member:r.values.MemberName,revision:r.revision,key,before:r.values[key],after:value};});}
export function applyBatch(state,preview){if(!preview.length)throw Error('Preview a change first.');for(const row of preview){const r=state.reports.find(r=>r.id===row.id);if(!r||r.revision!==row.revision||r.state==='Approved'||r.archived)throw Error('Reports changed after preview. Preview again.');const errors=validate({...r.values,[row.key]:row.after});if(Object.keys(errors).length)throw Error(Object.values(errors)[0]);}for(const row of preview)saveReport(state,row.id,{...state.reports.find(r=>r.id===row.id).values,[row.key]:row.after},'Reviewed bulk change');state.batches.push(clone(preview));state.operationHistory.unshift(`Applied bulk change to ${preview.length} reports`);}
export function traitAverage(values){
 if(values.ObservationStatus==='NOB')return {value:null,numeric:0,partial:false,nob:true};
 const numeric=fields.filter(f=>f.key.startsWith('Trait')).map(f=>values[f.key]).filter(v=>/^[1-5]\.0$/.test(v));
 return {value:numeric.length?numeric.reduce((s,v)=>s+Number(v),0)/numeric.length:null,numeric:numeric.length,partial:numeric.length<7,nob:false};
}
export function groupAverage(state,groupId){
 const group=state.groups.find(g=>g.id===Number(groupId));if(!group)return {value:null,count:0,partial:false};
 const means=group.members.map(id=>state.reports.find(r=>r.id===id)).filter(r=>r&&!r.archived).map(r=>traitAverage(r.values)).filter(a=>a.value!==null);
 const totals=groupTraitTotals(state,groupId);
 return {value:totals.count?totals.sum/totals.count:null,count:means.length,partial:means.some(a=>a.partial)};
}
export const reportingSeniorKeys=fields.filter(f=>f.key.startsWith('ReportingSenior')&&!f.legacyHidden).map(f=>f.key);
export const paygrades=['E1','E2','E3','E4','E5','E6','E7','E8','E9','W1','W2','W3','W4','W5','O1','O2','O3','O4','O5','O6'];
export function validateSenior(senior){
 const errors={};
 if(!senior.title?.trim()||senior.title.length>120)errors.title='Enter a record name of up to 120 characters.';
 for(const key of reportingSeniorKeys){const f=fields.find(f=>f.key===key),v=senior.values[key]??'';if(f.limit&&v.length>f.limit)errors[key]=`Use ${f.limit} characters or fewer.`;if(f.digits&&v&&!new RegExp(`^\\d{${f.digits}}$`).test(v))errors[key]=`Enter ${f.digits} digits.`;}
 return errors;
}
export function validateRSCA(h){
 const errors={};const count=Number(h.traitCount),avg=Number(h.rsca),sum=h.traitSum===''||h.traitSum==null?null:Number(h.traitSum);
 if(!paygrades.includes(h.grade))errors.grade='Choose a paygrade.';
 if(!validDate(h.asOf))errors.asOf='Enter a valid as-of date.';
 if(String(h.traitCount).trim()===''||!Number.isSafeInteger(count)||count<0||count>2147483647)errors.traitCount='Enter a whole trait count from 0 to 2147483647.';
 if((count>0||h.rsca!=null&&String(h.rsca).trim()!=='')&&(String(h.rsca).trim()===''||!Number.isFinite(avg)||avg<1||avg>5))errors.rsca='Enter an RSCA from 1.00 to 5.00.';
 if(sum!==null&&(!Number.isSafeInteger(sum)||sum<count||sum>5*count))errors.traitSum='Exact trait sum must be a whole number between trait count and five times trait count.';
 if(count>0&&sum!==null&&Number((sum/count).toFixed(2))!==Number(avg.toFixed(2)))errors.traitSum='Exact trait sum and count must agree with the entered RSCA rounded to two decimals.';
 if(count===0&&sum!==null&&sum!==0)errors.traitSum='A zero trait count requires a zero trait sum.';
 return errors;
}
export function groupTraitTotals(state,groupId,proposals={}){
 const g=state.groups.find(g=>g.id===Number(groupId));let sum=0,count=0,partial=false;
 for(const id of g?.members||[]){const r=state.reports.find(r=>r.id===id);if(!r||r.archived)continue;const v={...r.values,...proposals[id]},a=traitAverage(v);if(a.nob)continue;partial ||= a.partial;for(let n=33;n<=39;n++)if(/^[1-5]\.0$/.test(v['Trait'+n])){sum+=Number(v['Trait'+n]);count++;}}
 return {sum,count,partial};
}
export function psgRSCA(state,groupId,proposals={}){
 const g=state.groups.find(g=>g.id===Number(groupId)),senior=state.reportingSeniors?.find(s=>s.id===g?.seniorId),h=senior?.history?.find(h=>h.id===g?.rscaHistoryId);
 const unavailable=reason=>({value:null,reason});
 if(!senior||!h)return unavailable('Choose a reporting senior and a historical RSCA baseline.');
 if(h.grade!==g.grade)return unavailable('The baseline paygrade must match the summary group.');
 if(Object.keys(validateRSCA(h)).length)return unavailable('Correct the historical RSCA values first.');
 if(!g.baselineExcludesGroup)return unavailable('Confirm that the historical totals exclude this summary group to avoid double counting.');
 const rsSSN=senior.values.ReportingSeniorSSN;
 if(g.members.some(id=>state.reports.find(r=>r.id===id)?.archived))return unavailable('Remove or recover archived group members before calculating PSG RSCA.');
 if(!rsSSN||g.members.some(id=>{const r=state.reports.find(r=>r.id===id);return !r||r.values.ReportingSeniorSSN!==rsSSN;}))return unavailable('All group reports must use the selected reporting senior SSN. Apply the senior details to the reports first.');
 const totals=groupTraitTotals(state,groupId,proposals);if(!totals.count)return unavailable('No graded traits are available in this group.');
 const previousCount=Number(h.traitCount),exact=h.traitSum!==''&&h.traitSum!=null,previousSum=previousCount===0?0:exact?Number(h.traitSum):Number(h.rsca)*previousCount;
 return {value:(previousSum+totals.sum)/(previousCount+totals.count),estimated:previousCount>0&&!exact,previousSum,previousCount,groupSum:totals.sum,groupCount:totals.count,partial:totals.partial};
}
export function captureComparison(state,group){
 group.snapshot=group.members.map(id=>{const r=state.reports.find(r=>r.id===id);if(!r)throw Error('A group member no longer exists.');return {id,revision:r.revision,values:clone(r.values)};});
 return group.snapshot;
}
export function previewChanges(state,ids,changes){
 const keys=Object.keys(changes);if(!ids.length||!keys.length)throw Error('Change at least one block for selected reports.');
 if(keys.some(k=>k==='FormKind'||!fields.some(f=>f.key===k)))throw Error('Choose editable report values.');
 return ids.flatMap(id=>{const r=state.reports.find(r=>r.id===Number(id));if(!r||r.archived||r.state==='Approved')throw Error('Reopen approved or archived reports before applying a change.');const errors=validate({...r.values,...changes});if(Object.keys(errors).length)throw Error(Object.values(errors)[0]);return keys.filter(k=>r.values[k]!==changes[k]).map(k=>({id:r.id,member:r.values.MemberName,revision:r.revision,key:k,before:r.values[k],after:changes[k]}));});
}
export function applyChanges(state,preview){
 if(!preview.length)throw Error('No changed values to apply.');
 const updates=new Map();
 for(const row of preview){const r=state.reports.find(r=>r.id===row.id);if(!r||r.archived||r.state==='Approved'||r.revision!==row.revision)throw Error('Reports changed after preview. Preview again.');let update=updates.get(r.id);if(!update){update={r,values:clone(r.values),keys:new Set()};updates.set(r.id,update);}if(update.keys.has(row.key))throw Error('Conflicting proposals for one field.');update.keys.add(row.key);update.values[row.key]=row.after;}
 for(const u of updates.values()){const errors=validate(u.values);if(Object.keys(errors).length)throw Error(Object.values(errors)[0]);}
 for(const u of updates.values())saveReport(state,u.r.id,u.values,'Reviewed report-section changes');
 state.batches.push(clone(preview));state.operationHistory.unshift(`Applied ${preview.length} block changes to ${updates.size} reports`);
}
