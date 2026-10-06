"""Generate Access catalog code from reviewed JSON; never scrape wording into VBA."""
from pathlib import Path
import json

root = Path(__file__).resolve().parents[1]
items = json.loads((root / 'src/navywriter-catalog.json').read_text(encoding='utf-8-sig'))
def literal(s):
    return '"' + s.replace('"', '""').replace('\r', '').replace('\n', ' ') + '"'

lines = ['Attribute VB_Name = "modNavyWriterCatalog"', 'Option Explicit',
         'Public Function NavyWriterCount() As Long', f'NavyWriterCount = {len(items)}',
         'End Function', 'Public Function NavyWriterValue(ByVal row As Long, ByVal column As Long) As String',
         'Dim values As Variant', 'Select Case row']
for i, item in enumerate(items):
    values = [item['title'], item['body'], item['url'], '|'.join(item['rates']), item['kind']]
    if not item['url'].startswith('https://www.navywriter.com/'):
        raise ValueError('Unexpected source domain')
    if item['kind'] == 'Excerpt' and len(item['body'].split()) > 25:
        raise ValueError('Excerpt exceeds 25 words: ' + item['title'])
    lines.append(f'Case {i}: values = Array(' + ', '.join(map(literal, values)) + ')')
lines.extend(['Case Else: Err.Raise vbObjectError + 6921, , "Choose a catalog entry."',
              'End Select', 'NavyWriterValue = values(column)', 'End Function'])
rates = sorted({r for item in items for r in item['rates'] if r != 'GENERAL'})
lines.extend(['Public Function NavyWriterRates() As Variant',
              'NavyWriterRates = Array(' + ', '.join(map(literal, rates)) + ')', 'End Function'])
(root / 'src/modNavyWriterCatalog.bas').write_text('\n'.join(lines) + '\n', encoding='utf-8')
print(f'Generated {len(items)} catalog entries across {len(rates)} rate filters')
