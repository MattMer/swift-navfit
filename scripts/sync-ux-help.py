"""Emit VBA literals below the line-length limit from canonical help topics."""
import json
import re
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
lines += ['        Case Else: UXHelpText = "Choose a topic."','    End Select',
          '    UXHelpText = Replace(Replace(UXHelpText, "{APP_VERSION}", APP_VERSION), "{SCHEMA_VERSION}", CStr(SCHEMA_VERSION))',
          'End Function']
path.write_text(source[:start]+'\n'.join(lines)+source[end:],encoding='utf-8')
constants = (ROOT/'src/modSwift.bas').read_text(encoding='utf-8-sig')
version = re.search(r'APP_VERSION As String = "([^"]+)"', constants).group(1)
schema = re.search(r'SCHEMA_VERSION As Long = (\d+)', constants).group(1)
guide_path = ROOT/'planning/UX_WORKSPACE_GUIDE.md'
guide = guide_path.read_text(encoding='utf-8-sig')
version_body = topics[8]['Body'].replace('{APP_VERSION}', version).replace('{SCHEMA_VERSION}', schema)

#Keep every managed topic aligned with the native help, preserving the separate
#feature-destination table and other guide content after these sections.
for topic in topics:
    body=topic['Body'].replace('{APP_VERSION}',version).replace('{SCHEMA_VERSION}',schema)
    pattern=r'(## '+re.escape(topic['Title'])+r'\n\n).*?(?=\n## |\Z)'
    guide=re.sub(pattern,lambda match:match.group(1)+body+'\n',guide,flags=re.S)
guide = re.sub(r'^This is version [^\n]+', lambda _: version_body, guide, flags=re.M)
guide_path.write_text(guide, encoding='utf-8')
