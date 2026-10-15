// Summarize retained REA evidence without executing NAVFIT or claiming source recovery.
const fs = require('fs');
const path = require('path');
const directory = process.argv[2];
const record = JSON.parse(fs.readFileSync(path.join(directory, 'validation-members.json'), 'utf8').replace(/^\uFEFF/, ''));
const m = record.normalized_result, symbols = new Map();
for (const x of [...m.fields, ...m.methods, ...m.member_refs]) symbols.set(x.token, (x.declaring_type || x.owner_type || x.parent_name || '') + '.' + x.name);
for (const x of m.call_edges) symbols.set(x.target_token, x.target_declaring_type + '.' + x.target_name);
for (const x of m.field_accesses) symbols.set(x.field_token, x.field_declaring_type + '.' + x.field_name);
for (const x of m.user_strings.literals) symbols.set(x.token, JSON.stringify(x.value));
const methods = m.methods.map(method => ({name: method.name, token: method.token,
  fileOffset: method.body.file_offset, ilSize: method.body.il_size,
  strings: [...new Set(method.body.anchors.filter(x => x.opcode === 'ldstr').map(x => symbols.get(x.operand)))],
  calls: [...new Set(m.call_edges.filter(x => x.caller_token === method.token).map(x => x.target_declaring_type + '.' + x.target_name))],
  anchors: method.body.anchors.map(x => ({offset:x.il_offset, opcode:x.opcode, operand:symbols.get(x.operand) || x.operand}))}));
fs.writeFileSync(path.join(directory, 'summary.json'), JSON.stringify({evidence:record.evidence_id,subject:record.subject,coverage:m.coverage,limitations:record.limitations,methods}, null, 2));
for (const method of methods) fs.writeFileSync(path.join(directory, method.name + '.txt'), method.anchors.map(x => x.offset.toString(16).padStart(4,'0') + ' ' + x.opcode + ' ' + x.operand).join('\n'));
fs.writeFileSync(path.join(directory,'messages.txt'), methods.map(x => '\n' + x.name + ' ' + x.token + '\n' + x.strings.join('\n')).join('\n'));
console.log('Retained ' + methods.length + ' method summaries and CIL anchor listings.');
