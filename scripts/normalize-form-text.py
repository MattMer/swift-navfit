"""Trial (forms-as-source): normalize SaveAsText output by dropping
per-build noise: Checksum, GUID / NameMap / RecSrcDt hex blocks and the
BaseInfo combo cache, which embeds generated index names."""
import os, re, sys
src, dst = sys.argv[1], sys.argv[2]
os.makedirs(dst, exist_ok=True)
noisy = re.compile(r'^(\s*)(GUID|NameMap|RecSrcDt)\s*=\s*Begin\s*$')
stats = {}
for name in sorted(os.listdir(src)):
    if not name.endswith('.txt'):
        continue
    raw = open(os.path.join(src, name), 'rb').read()
    text = raw.decode('utf-16') if raw[:2] in (b'\xff\xfe', b'\xfe\xff') else raw.decode('utf-8-sig')
    out, skip, in_base = [], None, False
    for line in text.splitlines():
        if skip is not None:
            if line.strip() == 'End' and len(line) - len(line.lstrip()) == skip:
                skip = None
            continue
        m = noisy.match(line)
        if m:
            skip = len(m.group(1)); stats[m.group(2)] = stats.get(m.group(2), 0) + 1
            continue
        if line.startswith('Checksum ='):
            continue
        if in_base and line.lstrip().startswith('"'):
            continue
        in_base = line.lstrip().startswith('BaseInfo =')
        if in_base:
            stats['BaseInfo'] = stats.get('BaseInfo', 0) + 1
            continue
        out.append(line)
    open(os.path.join(dst, name), 'w', encoding='utf-8', newline='\r\n').write('\n'.join(out) + '\n')
print('removed blocks:', stats)
