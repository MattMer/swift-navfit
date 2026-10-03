"""Compare exports against source, allowing documented Access formatting only."""
from pathlib import Path
import re

root = Path(__file__).resolve().parents[1]
def canonical(text):
    text = re.sub(r'^Attribute VB_Name[^\n]*\n', '', text)
    text = re.sub(r'(?m)^([ \t]*Call[ \t]+\w+)\(\)([ \t]*)$', r'\1\2', text)
    parts = re.split(r'("(?:[^"]|"")*")', text)
    return ''.join(p if i % 2 else re.sub(r'\s+', '', p).lower() for i,p in enumerate(parts))

for name in ('modSwift', 'modWorkspace'):
    source = (root / 'src' / (name+'.bas')).read_text(encoding='utf-8-sig')
    installed = (root / 'build' / 'tests' / ('installed-'+name+'.bas')).read_text(encoding='utf-8-sig')
    if canonical(source) != canonical(installed):
        raise SystemExit('Installed source differs: '+name)
    print('Installed source verified:', name)
