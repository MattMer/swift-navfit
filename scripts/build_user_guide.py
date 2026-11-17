"""Render the maintained Swift NAVFIT Markdown guide as a bookmarked PDF.

Requires reportlab and pypdf. Run from any directory; outputs are relative to
the repository containing this script unless explicit paths are supplied.
"""
from pathlib import Path
from html import escape
import argparse
import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether, Image,
)
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader

#Palette: charcoal text, forest green structure, amber for Important notices.
FOREST = colors.HexColor('#2F5D3E')
CHARCOAL = colors.HexColor('#2B2B2B')
INK = colors.HexColor('#2B2B2B')
PALE = colors.HexColor('#F4F2EC')
RULE = colors.HexColor('#D9D4C7')
AMBER_FILL = colors.HexColor('#FFF3DD')
AMBER_BAR = colors.HexColor('#C27C0E')
NOTE_FILL = colors.HexColor('#EEF3EE')
#Body text sits right of a heading margin, so lines stay near 75 characters.
INDENT = 0.85 * inch


def font_setup():
    windows = Path('C:/Windows/Fonts')
    regular, bold = windows / 'segoeui.ttf', windows / 'segoeuib.ttf'
    italic, bold_italic = windows / 'segoeuii.ttf', windows / 'segoeuiz.ttf'
    if regular.exists() and bold.exists():
        pdfmetrics.registerFont(TTFont('Guide', str(regular)))
        pdfmetrics.registerFont(TTFont('Guide-Bold', str(bold)))
        slanted = italic.exists() and bold_italic.exists()
        if slanted:
            pdfmetrics.registerFont(TTFont('Guide-Italic', str(italic)))
            pdfmetrics.registerFont(TTFont('Guide-BoldItalic', str(bold_italic)))
        pdfmetrics.registerFontFamily('Guide', normal='Guide', bold='Guide-Bold',
                                    italic='Guide-Italic' if slanted else 'Guide',
                                    boldItalic='Guide-BoldItalic' if slanted else 'Guide-Bold')
        return 'Guide', 'Guide-Bold'
    return 'Helvetica', 'Helvetica-Bold'


def inline(text):
    text = escape(text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'(?<![\w*])\*([^*\n]+?)\*(?![\w*])', r'<i>\1</i>', text)
    return re.sub(r'`(.+?)`', r'<font color="#7A4A0E">\1</font>', text)


class GuideDoc(SimpleDocTemplate):
    """Records the page each chapter starts on, for the contents table."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.chapter_pages = []

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and flowable.style.name in ('Chapter', 'Cover'):
            self.chapter_pages.append(self.page)
            if flowable.style.name == 'Chapter':
                title = flowable.getPlainText()
                self.canv.bookmarkPage('chapter' + str(self.page))
                self.canv.addOutlineEntry(title, 'chapter' + str(self.page), 0)


def build(source, output):
    """Render twice: the first pass measures where each chapter starts, the
    second fills those page numbers into the contents table."""
    pages = render(source, output, None)
    final = render(source, output, pages)
    if final != pages:
        raise RuntimeError(f'Chapter pages moved between passes: {pages} then {final}')
    reader = PdfReader(output)
    for n, page in enumerate(reader.pages):
        if len(page.extract_text()) < 80:
            raise RuntimeError(f'Unexpected sparse or blank page {n + 1}')
    print(f'PASS: {len(reader.pages)} pages, {len(pages)} chapters on their own pages, '
          f'contents page numbers and text extraction; {output}')


def render(source, output, chapter_pages):
    regular, bold = font_setup()
    styles = {
        'body': ParagraphStyle('Body', fontName=regular, fontSize=11, leading=15.5,
                               textColor=INK, spaceAfter=8, leftIndent=INDENT),
        'chapter': ParagraphStyle('Chapter', fontName=bold, fontSize=24, leading=29,
                                  textColor=CHARCOAL, spaceAfter=17, keepWithNext=True),
        'sub': ParagraphStyle('Section', fontName=bold, fontSize=13, leading=17,
                              textColor=FOREST, spaceBefore=12, spaceAfter=7,
                              keepWithNext=True),
        'cell': ParagraphStyle('Cell', fontName=regular, fontSize=10, leading=13.2,
                               textColor=INK),
        'head': ParagraphStyle('TableHeader', fontName=bold, fontSize=10,
                               leading=13.2, textColor=colors.white),
        'step': ParagraphStyle('Step', fontName=regular, fontSize=11, leading=15.5,
                               textColor=INK, leftIndent=INDENT + 22, bulletIndent=INDENT,
                               spaceAfter=7),
        'bullet': ParagraphStyle('Bullet', fontName=regular, fontSize=11, leading=15.5,
                                 textColor=INK, leftIndent=INDENT + 18, bulletIndent=INDENT + 5,
                                 spaceAfter=5),
        'caption': ParagraphStyle('Caption', fontName=regular, fontSize=9.5, leading=12.5,
                                  textColor=colors.HexColor('#5A5650'), spaceBefore=3, spaceAfter=10,
                                  leftIndent=INDENT),
        'bullet2': ParagraphStyle('Bullet2', fontName=regular, fontSize=11, leading=15.5,
                                  textColor=INK, leftIndent=INDENT + 36, bulletIndent=INDENT + 23,
                                  spaceAfter=5),
        'notice': ParagraphStyle('Notice', fontName=regular, fontSize=11, leading=15.5,
                                 textColor=INK),
    }
    text_source = source.read_text(encoding='utf-8')
    version = re.search(r'^Version ([^|\n]+)\|', text_source, re.M)
    product = re.search(r'^For Swift NAVFIT ([^|\n]+)\|', text_source, re.M)
    header = 'Version ' + (version.group(1).strip() if version else '?') + \
        ('  |  ' + product.group(1).strip() if product else '')
    chapters = text_source.split('\n---\n')
    story = []
    width = 6.8 * inch
    body_width = width - INDENT
    for index, chapter in enumerate(chapters):
        if index:
            story.append(PageBreak())
        lines = chapter.strip().splitlines()
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
            picture = re.fullmatch(r'!\[(.*)\]\((.+)\)', line)
            if picture:
                #A screenshot: the body width at most, 4.2 inches tall at most, kept with its caption.
                path = source.parent / picture.group(2)
                pixels_w, pixels_h = ImageReader(str(path)).getSize()
                scale = min(body_width / pixels_w, 4.2 * inch / pixels_h)
                #Start the screenshot at the body indent, like the text around it.
                frame = Table([[Image(str(path), pixels_w * scale, pixels_h * scale)]],
                              hAlign='LEFT', colWidths=[pixels_w * scale])
                frame = Table([['', frame]], colWidths=[INDENT, pixels_w * scale], hAlign='LEFT',
                              style=[('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                                     ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)])
                inner = frame._cellvalues[0][1]
                inner.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), .5, RULE),
                                           ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                                           ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))
                story.append(KeepTogether([Spacer(1, 2), frame, Paragraph(inline(picture.group(1)), styles['caption'])]))
                i += 1
                continue
            if line.startswith('>'):
                #An Important or Note box, at body width, with a colored bar on its left.
                text = [line.lstrip('>').strip()]
                while i + 1 < len(lines) and lines[i + 1].strip().startswith('>'):
                    i += 1
                    text.append(lines[i].strip().lstrip('>').strip())
                joined = ' '.join(text)
                important = joined.startswith('**Important')
                box = Table([[Paragraph(inline(joined), styles['notice'])]], colWidths=[body_width],
                            hAlign='RIGHT')
                box.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), AMBER_FILL if important else NOTE_FILL),
                    ('LINEBEFORE', (0, 0), (0, -1), 4, AMBER_BAR if important else FOREST),
                    ('LEFTPADDING', (0, 0), (-1, -1), 11), ('RIGHTPADDING', (0, 0), (-1, -1), 9),
                    ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
                ]))
                box.spaceBefore, box.spaceAfter = 3, 9
                story.append(box)
                i += 1
                continue
            if line.startswith('|'):
                rows = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                    if not all(re.fullmatch(r'[:\- ]+', c) for c in cells):
                        rows.append(cells)
                    i += 1
                contents = index == 1 and rows[0][-1] == 'Page'
                if contents:
                    widths = [width * .12, width * .77, width * .11]
                elif len(rows[0]) == 2:
                    widths = [width * .32, width * .68]
                elif len(rows[0]) == 3:
                    widths = [width * .26, width * .37, width * .37]
                elif len(rows[0]) == 4:
                    widths = [width * .1, width * .2, width * .16, width * .54]
                else:
                    widths = [width / len(rows[0])] * len(rows[0])
                if contents and chapter_pages:
                    #Contents: chapter N starts where the (N + 1)th rendered chapter does.
                    for row in rows[1:]:
                        if row[0].isdigit() and int(row[0]) + 1 < len(chapter_pages):
                            row[2] = str(chapter_pages[int(row[0]) + 1])
                formatted = []
                for rn, row in enumerate(rows):
                    items = []
                    for cn, cell in enumerate(row):
                        text = inline(cell)
                        if contents and rn > 0 and cn == 2:
                            text = f'<link href="#p{cell}" color="#2F5D3E">{text}</link>'
                        items.append(Paragraph(text, styles['head' if rn == 0 else 'cell']))
                    formatted.append(items)
                table = Table(formatted, colWidths=widths, repeatRows=1, hAlign='LEFT')
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), FOREST),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [PALE, colors.white]),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 9),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 9),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ('LINEBELOW', (0, -1), (-1, -1), .5, RULE),
                ]))
                table.spaceAfter = 8
                story.append(table)
                continue
            if line.startswith('# '):
                if index == 0:
                    story.append(Spacer(1, 47))
                    cover_style = ParagraphStyle('Cover', parent=styles['chapter'],
                                                 fontSize=36, leading=43)
                    story.append(Paragraph(inline(line[2:]), cover_style))
                    story.append(Spacer(1, 14))
                else:
                    story.append(Paragraph(inline(line[2:]), styles['chapter']))
            elif line.startswith(('## ', '### ')):
                story.append(Paragraph(inline(line.lstrip('#').strip()), styles['sub']))
            elif re.match(r'^\d+\. ', line):
                #The step number hangs in the margin, so wrapped lines align with the text.
                number, step = line.split(' ', 1)
                story.append(Paragraph(inline(step), styles['step'], bulletText=number))
            elif re.match(r'^\s*- ', lines[i]):
                nested = len(lines[i]) - len(lines[i].lstrip()) >= 2
                text = [line[2:]]
                while i + 1 < len(lines) and lines[i + 1].strip() and not re.match(
                        r'^\s*(#|\||\d+\. |- )', lines[i + 1]):
                    i += 1
                    text.append(lines[i].strip())
                story.append(Paragraph(inline(' '.join(text)), styles['bullet2' if nested else 'bullet'],
                                       bulletText='\u2022'))
            else:
                text = [line]
                while i + 1 < len(lines) and lines[i + 1].strip() and not re.match(
                        r'^(#|\||\d+\. |- )', lines[i + 1].strip()):
                    i += 1
                    text.append(lines[i].strip())
                story.append(Paragraph(inline(' '.join(text)), styles['body']))
            i += 1

    def furniture(canvas, doc):
        canvas.saveState()
        canvas.bookmarkPage('p' + str(doc.page))
        w, h = doc.pagesize
        canvas.setFillColor(FOREST)
        canvas.rect(0, h - 9, w, 9, fill=1, stroke=0)
        canvas.setFont(bold, 8)
        canvas.setFillColor(CHARCOAL)
        canvas.drawString(doc.leftMargin, h - 32, 'SWIFT NAVFIT  /  USER GUIDE')
        canvas.setFont(regular, 8)
        canvas.drawRightString(w - doc.rightMargin, h - 32, header)
        canvas.setStrokeColor(RULE)
        canvas.line(doc.leftMargin, 43, w - doc.rightMargin, 43)
        canvas.setFillColor(INK)
        canvas.setFont(regular, 7.5)
        canvas.drawString(doc.leftMargin, 29, 'Local workflow guide  |  Independent development application')
        canvas.drawRightString(w - doc.rightMargin, 29, str(doc.page))
        canvas.restoreState()

    output.parent.mkdir(parents=True, exist_ok=True)
    doc = GuideDoc(str(output), pagesize=(8.5 * inch, 11 * inch),
                   leftMargin=.85 * inch, rightMargin=.85 * inch,
                   topMargin=.78 * inch, bottomMargin=.76 * inch,
                   title='Swift NAVFIT User Guide', author='Swift NAVFIT Project',
                   subject='Operating procedures for Swift NAVFIT ' + (product.group(1).strip() if product else ''))
    doc.build(story, onFirstPage=furniture, onLaterPages=furniture)
    if len(doc.chapter_pages) != len(chapters):
        raise RuntimeError(f'Expected {len(chapters)} titled sections; found {len(doc.chapter_pages)}')
    return doc.chapter_pages


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    root = Path(__file__).resolve().parents[1]
    parser.add_argument('--source', type=Path, default=root / 'docs/Swift_NAVFIT_User_Guide.md')
    parser.add_argument('--output', type=Path, default=root / 'docs/Swift_NAVFIT_User_Guide.pdf')
    args = parser.parse_args()
    build(args.source, args.output)
