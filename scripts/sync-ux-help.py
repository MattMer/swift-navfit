"""Emit VBA literals below the line-length limit from canonical help topics."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
topics = json.loads((ROOT/'src/ux-help-topics.json').read_text(encoding='utf-8-sig'))
path = ROOT/'src/modUXHelp.bas'
source = path.read_text(encoding='utf-8-sig')
start = source.index('Public Function UXHelpText(')
end = source.index('End Function', start)+len('End Function')
lines = ['Public Function UXHelpText(ByVal topic As Long) As String','    Select Case topic']
for index, topic in enumerate(topics):
    chunks=[topic['Body'][i:i+650].replace('"','""') for i in range(0,len(topic['Body']),650)]
    lines.append(f'        Case {index}: UXHelpText = "'+chunks[0]+'"'+(' & _' if len(chunks)>1 else ''))
    for i,chunk in enumerate(chunks[1:]):
        lines.append('            "'+chunk+'"'+(' & _' if i<len(chunks)-2 else ''))
lines += ['        Case Else: UXHelpText = "Choose a topic."','    End Select','End Function']
path.write_text(source[:start]+'\n'.join(lines)+source[end:],encoding='utf-8')
