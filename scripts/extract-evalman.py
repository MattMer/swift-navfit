"""Build src/evalman.json, the in-app EVALMAN text, from the official BUPERSINST 1610.10H PDF.

The PDF is a scan with recognized text. pdftotext (poppler, shipped with Git for Windows)
reads that text layer; src/evalman-corrections.json then fixes recognition errors and
supplies a reviewed transcription for each exhibit page, whose recognized text is not
usable. Running headers and page footers are dropped; each page keeps its printed number.

Usage: python scripts/extract-evalman.py <BUPERSINST-1610.10H_CHn.pdf>
"""
import hashlib, json, re, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
pdf = Path(sys.argv[1])
fixes = json.loads((ROOT / 'src/evalman-corrections.json').read_text(encoding='utf-8'))
raw = subprocess.run(['pdftotext', '-layout', str(pdf), '-'], capture_output=True, check=True).stdout.decode('utf-8', 'replace')
pages = raw.split('\f')
if pages and not pages[-1].strip():
    pages = pages[:-1]

HEADER = re.compile(r'^\s*(BUPERSINST|BUPERS INST)\b.*$|^\s*([I1l]|\d{1,2})\s?[A-Z][a-z]{2}\s?\d{4}\s*$|^\s*PERS-32\s*$')
DATELINE = re.compile(r'^\s*([I1l]|\d{1,2})\s?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s?\d{4}\s*$')
# Margin labels beside EVAL block text name the matching FITREP and CHIEFEVAL blocks.
MARGIN = re.compile(r'^\s*(FITREP|CHIEFEVAL|BLOCKS?\s+\d{1,2}(\s*(-|and|&)\s*\d{1,2})?)(\s{2,}(\S.*))?\s*$')
FOOTER = re.compile(r'^\s*([ivx]+|[I1l]?\d{0,2}-?\d{1,2}[a-b]?|I-\d)\s+Enclos\w*\s*\(.{1,3}\)\s*$|^\s*Enclos\w*\s*\(.{1,3}\)\s*$|^\s*[ivx\d]{1,3}[a-b]?\s*$')
PRINTED = re.compile(r'^\s*([ivx]+|I-\d|\d{1,2}-\d{1,2}[a-b]?|\d{1,2}[a-b]?)\b')


def clean_page(n, text):
    for old, new in fixes['replace']:
        text = text.replace(old, new)
    text = text.replace('�', '-')
    lines = text.split('\n')
    while lines and not lines[-1].strip():
        lines.pop()
    printed = ''
    # footer: page number and enclosure line(s) at the bottom
    for _ in range(3):
        if lines and FOOTER.match(lines[-1]):
            m = PRINTED.match(lines[-1])
            if m and not printed:
                printed = m.group(1)
            lines.pop()
            while lines and not lines[-1].strip():
                lines.pop()
    # header: instruction number and date at the top
    body = []
    head = True
    for line in lines:
        if head and (not line.strip() or HEADER.match(line)):
            continue
        head = False
        body.append(line.rstrip())
    override = fixes['pages'].get(str(n))
    if override is not None:
        body = override.split('\n')
    indent = min((len(l) - len(l.lstrip()) for l in body if l.strip()), default=0)
    body = [l[indent:] for l in body]
    return printed, body


TABULAR = re.compile(r'\S {3,}\S')
LISTSTART = re.compile(r'^\s*(\(?[a-z0-9]{1,3}[.)]|\d{1,2}-\d{1,2}\.|Note\b|NOTE\b|- )')
# Headings stand on their own line: chapter, exhibit, block and all-capital headings.
NUMS = r"\d{1,2}(\s*(-|and|&|,|through)\s*\d{1,2})*"
# A block heading has its number, two or more spaces and a short name without a closing
# period; a short all-capital line (such as UIC above "BLOCK 26 Enter ...") is a name line.
HEADING = re.compile(r"^\s*(CHAPTER\s*\d|EXHIBIT\s|((EVAL|FITREP|CHIEFEVAL)\s+)?BLOCKS?\s+" + NUMS + r"(\s{2,}[^.]{1,60})?\s*$|[A-Z][A-Z0-9 ,/&()'-]{1,40}$)")
BLOCKSTART = re.compile(r"^\s*((EVAL|FITREP|CHIEFEVAL)\s+)?BLOCKS?\s+\d")


def reflow(lines):
    """Join wrapped prose lines into paragraphs; keep tables and headings line by line."""
    out, para = [], ''
    margin, form_heading = [], False
    def flush():
        nonlocal para
        if para:
            out.append(para)
            para = ''
    for k, line in enumerate(lines):
        if DATELINE.match(line):
            continue
        s = line.strip()
        mm = MARGIN.match(line)
        if form_heading and mm and not re.match(r'^\s*(EVAL|FITREP|CHIEFEVAL)\s+BLOCK', line):
            margin.append(mm.group(1))
            if not mm.group(5):
                continue
            line = s = mm.group(5).strip()
        elif margin and (not s or HEADING.match(line)):
            labels, pair = [], ''
            for m in margin:
                if m in ('FITREP', 'CHIEFEVAL'):
                    pair = m
                else:
                    labels.append((pair + ' ' if pair else '') + m.lower()); pair = ''
            if labels:
                flush(); out.append('(' + '; '.join(labels) + ')')
            margin = []
        if re.match(r'^\s*(EVAL|FITREP|CHIEFEVAL)\s+BLOCK', line):
            form_heading = True
        elif re.match(r'^\s*(CHAPTER|EXHIBIT|BLOCKS?\s+\d|\d{1,2}-\d{1,2}\.)', line):
            form_heading = False
        if not s:
            following = next((x.strip() for x in lines[k + 1:] if x.strip() and not DATELINE.match(x)), '')
            if para and not re.search(r'[.:;?!)]["\']?$', para) and following[:1].islower():
                continue
            flush()
            if out and out[-1] != '':
                out.append('')
            continue
        if HEADING.match(line):
            flush(); out.append(re.sub(r' {2,}', '  ', s)); continue
        if TABULAR.search(line) or (LISTSTART.match(line) or BLOCKSTART.match(line)) and para:
            flush()
        if TABULAR.search(line):
            out.append(re.sub(r' {3,}', '    ', line.rstrip()))
            continue
        if para.endswith('-') and s[:1].islower():
            para = para[:-1] + s
        elif para:
            para += ' ' + s
        else:
            para = s
    flush()
    while out and out[-1] == '':
        out.pop()
    return out


PARA = re.compile(r'^(\d{1,2}-\d{1,2})\.\s+(.+?)((?<!\b[A-Z])\.(\s|$)|$)')
BLOCKHEAD = re.compile(r'^((EVAL|FITREP|CHIEFEVAL)\s+)?(BLOCKS?)\s+(\d{1,2}(?:\s*(?:-|and|&|,|through)\s*\d{1,2})*)\s*(\S.*)?$')
KEEP = {'NOB', 'DUINS', 'INACT', 'PIM', 'ACDU', 'EVAL', 'EVALS', 'FITREP', 'FITREPS', 'CHIEFEVAL', 'CHIEFEVALS', 'U.S.', 'ISIC', 'TEMDU', 'PFA', 'RSCA', 'UIC', 'SSN', 'CO', 'OIC', 'GCMCA', 'LDO', 'CWO', 'IET', 'TEMDUINS', 'TEMADDINS', 'AT', 'ADT', 'BCNR', 'PERS-32', 'OMPF', 'USNR', 'ADSW', 'NAVPERS', 'DOD', 'PEP', 'SNDL'}


def name_like(text):
    # A heading name reads like a title: it starts with a capital and most of its words do.
    words = [w for w in re.findall(r"[A-Za-z][A-Za-z/'-]*", text) if w.lower() not in {'a', 'an', 'and', 'or', 'of', 'on', 'to', 'the', 'for', 'in', 'by', 'at', 'with', 'from', 'if', 'any'}]
    return bool(words) and text[:1].isupper() and sum(w[:1].isupper() for w in words) >= 0.8 * len(words)


def title_case(text):
    if re.search('[a-z]', text):
        return text
    small = {'a', 'an', 'and', 'or', 'of', 'on', 'to', 'the', 'for', 'in', 'by', 'at', 'with', 'from', 'non'}
    words = []
    for i, w in enumerate(text.split()):
        core = w.strip('(),.:;')
        if core.upper() in KEEP or re.fullmatch(r'[A-Z]{1,2}-?\d.*|\d.*', core):
            words.append(w)
        elif i and core.lower() in small:
            words.append(w.lower())
        else:
            lead = '(' if w.startswith('(') else ''
            words.append(lead + '/'.join(x[:1].upper() + x[1:].lower() for x in w[len(lead):].split('/')))
    return re.sub(r'\bu\.s\.', 'U.S.', ' '.join(words))
CHAPTER = re.compile(r'^CHAPTER\s*(\d{1,2})\b')
EXHIBIT = re.compile(r'^EXHIBIT\s+(\d{1,2}-\d{1,2})\b')
TOP = re.compile(r'^(\d{1,2})\.\s+([A-Z][^.?]{2,80}[.?])')

parts = fixes['parts']  # [{"title", "first", "last", "kind"}], by PDF page
sections = []


def start(level, title, pdf_page, printed):
    sections.append({'level': level, 'title': title.strip(), 'pdfPage': pdf_page, 'page': printed, 'lines': []})


for part in parts:
    start(0, part['title'], part['first'], '')
    for n in range(part['first'], part['last'] + 1):
        printed, body = clean_page(n, pages[n - 1])
        if part['kind'] == 'skip':
            continue
        text = reflow(body)
        marker = f'[Page {printed}]' if printed else f'[PDF page {n}]'
        sections[-1]['lines'] += ['', marker, '']
        chapter_title_next = False
        prev = ''
        for line in text:
            if part['kind'] == 'enclosure2':
                m, c, e = PARA.match(line), CHAPTER.match(line), EXHIBIT.match(line)
                if c:
                    start(1, 'Chapter ' + c.group(1), n, printed); chapter_title_next = True
                    sections[-1]['lines'] += ['', marker, '']
                elif chapter_title_next and line.strip() and not line.startswith('['):
                    sep = ' ' if ' - ' in sections[-1]['title'] else ' - '
                    sections[-1]['title'] += sep + title_case(line.strip())
                    sections[-1]['lines'].append(line); continue
                elif chapter_title_next and not line.strip() and ' - ' in sections[-1]['title']:
                    chapter_title_next = False
                elif e:
                    start(2, 'Exhibit ' + e.group(1), n, printed)
                    sections[-1]['lines'] += ['', marker, '']
                elif m and len(m.group(2)) < 200:
                    start(2, m.group(1) + '. ' + m.group(2).strip(), n, printed)
                    sections[-1]['lines'] += ['', marker, '']
                elif BLOCKHEAD.match(line) and sections[-1]['level'] >= 2 and 'Exhibit' not in sections[-1]['title']:
                    b = BLOCKHEAD.match(line)
                    rest = (b.group(5) or '').strip()
                    spaced = re.match(r"^\s*((EVAL|FITREP|CHIEFEVAL)\s+)?BLOCKS?\s+" + NUMS + r"\s{2,}\S", line)
                    if rest and name_like(rest) and (spaced and len(rest) <= 60 and not rest.endswith('.') or rest.upper() == rest):
                        name = rest
                    elif re.fullmatch(r"[A-Z][A-Z0-9 ,/&()'.-]{1,60}", prev):
                        name = prev
                    else:
                        name = ''
                    if name:
                        label = ((b.group(2) or '') + ' ' + b.group(3).title() + ' ' + b.group(4).strip()).strip()
                        if name:
                            label += ': ' + title_case(name)
                            if name == prev and sections[-1]['lines'] and sections[-1]['lines'][-1].strip() == prev:
                                sections[-1]['lines'].pop()
                        start(3, label, n, printed)
                        sections[-1]['lines'] += ['', marker, '']
                        if name == prev:
                            sections[-1]['lines'].append(prev)
            elif part['kind'] == 'numbered':
                t = TOP.match(line)
                if t and len(t.group(2)) < 90:
                    start(1, t.group(1) + '. ' + t.group(2).rstrip('.'), n, printed)
                    sections[-1]['lines'] += ['', marker, '']
            sections[-1]['lines'].append(line)
            if line.strip():
                prev = line.strip()

# Exhibit titles: the first line of the exhibit text.
for s in sections:
    if s['title'].startswith('Exhibit '):
        body = [l for l in s['lines'] if l and not l.startswith('[') and not l.upper().startswith('EXHIBIT')]
        if body:
            s['title'] += ' - ' + body[0][:80]

out = {
    'title': 'BUPERSINST 1610.10H, Navy Performance Evaluation System (EVALMAN)',
    'change': fixes['change'],
    'changeDate': fixes['changeDate'],
    'source': pdf.name,
    'sourceSha256': hashlib.sha256(pdf.read_bytes()).hexdigest(),
    'pdfPages': len(pages),
    'note': 'Text from the official PDF, with recognition errors corrected and exhibits transcribed. The PDF is the authoritative copy.',
    'sections': [{k: s[k] for k in ('level', 'title', 'page', 'pdfPage')} | {'text': '\n'.join(s['lines']).strip('\n')} for s in sections if any(l.strip() for l in s['lines']) or s['level'] == 0],
}
(ROOT / 'src/evalman.json').write_text(json.dumps(out, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
chars = sum(len(s['text']) for s in out['sections'])
print(f"{len(out['sections'])} sections, {chars} characters, largest {max(len(s['text']) for s in out['sections'])}")
