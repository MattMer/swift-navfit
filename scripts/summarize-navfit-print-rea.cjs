const fs = require('fs');
const path = require('path');
const root = process.argv[2];
const out = process.argv[3];
const read = name => JSON.parse(fs.readFileSync(path.join(root, name + '.json'), 'utf8').replace(/^\uFEFF/, ''));
const artifact = read('artifact'), print = read('print-members'), dependencies = read('dependencies');
const m = print.normalized_result;
const names = ['PrintEval','PrintChief','PrintFitRep','Print_Folder','PrintSummary','NF_PrtEval','NF_PrtChief','NF_PrtFitRep','NF_PrtFolder','NF_PrtSummary','NF_EvalPage2Frame','NF_EvalPage2Text','NF_EvalPage2Data','NF_ChiefPage1Text','NF_ChiefPage2Text','NF_ChiefPage2Data','NF_FitRepPage2Text','NF_FitRepPage2Data'];
const selected = names.map(name => {
 const method = m.methods.find(x => x.name === name);
 return {name, token:method.token, public:(method.flags & 7)===6, static:!!(method.flags & 16), ilSize:method.body.il_size, instructionCount:method.body.instruction_count,
  directCalls:[...new Set(m.call_edges.filter(x=>x.caller_token===method.token).map(x=>x.target_declaring_type+'.'+x.target_name))],
  stateFields:[...new Set(m.field_accesses.filter(x=>x.method_token===method.token).map(x=>x.field_declaring_type+'.'+x.field_name))]};
});
const symbols = new Map();
for(const x of [...m.fields,...m.methods,...m.member_refs])symbols.set(x.token,(x.declaring_type||x.owner_type||x.parent_name||'')+'.'+x.name);
for(const x of m.call_edges)symbols.set(x.target_token,x.target_declaring_type+'.'+x.target_name);
for(const x of m.field_accesses)symbols.set(x.field_token,x.field_declaring_type+'.'+x.field_name);
for(const x of m.user_strings.literals)symbols.set(x.token,JSON.stringify(x.value));
const listing = [];
for(const name of ['.cctor','NF_EvalPage2Data','NF_ChiefPage2Data','NF_FitRepPage2Data']){
 const method=m.methods.find(x=>x.name===name); listing.push('\n'+name+' '+method.token+' [selected REA CIL anchors, not full source]');
 for(const x of method.body.anchors)listing.push(x.il_offset.toString(16).padStart(4,'0')+' '+x.opcode+' '+(symbols.get(x.operand)||x.operand||''));
}
const summary = {date:'2026-10-14', tool:'REA 6.5.0 installed CLI', connection:'No MCP registration in this session; CLI used without changing client configuration', target:artifact.subject, module:artifact.normalized_result.module, assembly:artifact.normalized_result.assembly, framework:artifact.normalized_result.target_frameworks,
 evidence:[artifact,print,dependencies].map(x=>({id:x.evidence_id,operation:x.operation,provider:x.provider,coverage:x.normalized_result.coverage,limitations:x.limitations})),
 selectedType:m.selection, printMethodCount:m.methods.length, methods:selected, limitations:['Static REA analysis does not execute the binary.','CIL anchors are not a source reconstruction or full behavioral proof.','Offscreen runtime proof is separately recorded and covers frames/text only.']};
fs.writeFileSync(out,JSON.stringify(summary,null,2)+'\n');
fs.writeFileSync(path.join(root,'narrative-anchors.txt'),listing.join('\n')+'\n');
console.log('Recorded '+selected.length+' focused methods; 3 complete local REA Evidence files retained.');
