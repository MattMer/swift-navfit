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
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether,
)
from pypdf import PdfReader

NAVY = colors.HexColor('#15304D')
TEAL = colors.HexColor('#0D7475')
INK = colors.HexColor('#23374A')
PALE = colors.HexColor('#EFF4F8')


def font_setup():
    windows = Path('C:/Windows/Fonts')
    regular, bold = windows / 'segoeui.ttf', windows / 'segoeuib.ttf'
    if regular.exists() and bold.exists():
        pdfmetrics.registerFont(TTFont('Guide', str(regular)))
        pdfmetrics.registerFont(TTFont('Guide-Bold', str(bold)))
        pdfmetrics.registerFontFamily('Guide', normal='Guide', bold='Guide-Bold',
                                    italic='Guide', boldItalic='Guide-Bold')
        return 'Guide', 'Guide-Bold'
    return 'Helvetica', 'Helvetica-Bold'


def inline(text):
    text = escape(text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    return re.sub(r'`(.+?)`', r'<font color="#0D7475">\1</font>', text)


class GuideDoc(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and flowable.style.name == 'Chapter':
            title = flowable.getPlainText()
            self.canv.bookmarkPage('chapter' + str(self.page))
            self.canv.addOutlineEntry(title, 'chapter' + str(self.page), 0)


def build(source, output):
    regular, bold = font_setup()
    styles = {
        'body': ParagraphStyle('Body', fontName=regular, fontSize=9.5, leading=13.2,
                               textColor=INK, spaceAfter=7),
        'chapter': ParagraphStyle('Chapter', fontName=bold, fontSize=23, leading=28,
                                  textColor=NAVY, spaceAfter=17, keepWithNext=True),
        'sub': ParagraphStyle('Section', fontName=bold, fontSize=12, leading=16,
                              textColor=TEAL, spaceBefore=9, spaceAfter=7,
                              keepWithNext=True),
        'cell': ParagraphStyle('Cell', fontName=regular, fontSize=8.8, leading=11.5,
                               textColor=INK),
        'head': ParagraphStyle('TableHeader', fontName=bold, fontSize=9.1,
                               leading=11.5, textColor=colors.white),
        'step': ParagraphStyle('Step', fontName=regular, fontSize=9.5, leading=13.2,
                               textColor=INK, leftIndent=21, firstLineIndent=-19,
                               spaceAfter=7),
    }
    chapters = source.read_text(encoding='utf-8').split('\n---\n')
    story = []
    width = 6.95 * inch
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
            if line.startswith('|'):
                rows = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                    if not all(re.fullmatch(r'[:\- ]+', c) for c in cells):
                        rows.append(cells)
                    i += 1
                if len(rows[0]) == 3:
                    widths = [width * .12, width * .77, width * .11]
                else:
                    widths = [width * .32, width * .68]
                formatted = []
                for rn, row in enumerate(rows):
                    items = []
                    for cn, cell in enumerate(row):
                        text = inline(cell)
                        if index == 1 and rn > 0 and cn == 2:
                            text = f'<link href="#p{cell}" color="#0D7475">{text}</link>'
                        items.append(Paragraph(text, styles['head' if rn == 0 else 'cell']))
                    formatted.append(items)
                table = Table(formatted, colWidths=widths, repeatRows=1, hAlign='LEFT')
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), NAVY),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [PALE, colors.white]),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 9),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 9),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ('LINEBELOW', (0, -1), (-1, -1), .5, colors.HexColor('#CCD8E0')),
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
            elif line.startswith('## '):
                story.append(Paragraph(inline(line[3:]), styles['sub']))
            elif re.match(r'^\d+\. ', line):
                story.append(Paragraph(inline(line), styles['step']))
            else:
                text = [line]
                while i + 1 < len(lines) and lines[i + 1].strip() and not re.match(
                        r'^(#|\||\d+\. )', lines[i + 1].strip()):
                    i += 1
                    text.append(lines[i].strip())
                story.append(Paragraph(inline(' '.join(text)), styles['body']))
            i += 1

    def furniture(canvas, doc):
        canvas.saveState()
        canvas.bookmarkPage('p' + str(doc.page))
        w, h = doc.pagesize
        canvas.setFillColor(TEAL)
        canvas.rect(0, h - 9, w, 9, fill=1, stroke=0)
        canvas.setFont(bold, 8)
        canvas.setFillColor(NAVY)
        canvas.drawString(doc.leftMargin, h - 32, 'SWIFT NAVFIT  /  USER GUIDE')
        canvas.setFont(regular, 8)
        canvas.drawRightString(w - doc.rightMargin, h - 32, 'Version 1.0  |  r123a')
        canvas.setStrokeColor(colors.HexColor('#CCD8E0'))
        canvas.line(doc.leftMargin, 43, w - doc.rightMargin, 43)
        canvas.setFillColor(INK)
        canvas.setFont(regular, 7.5)
        canvas.drawString(doc.leftMargin, 29, 'Local workflow guide  |  Independent development application')
        canvas.drawRightString(w - doc.rightMargin, 29, str(doc.page))
        canvas.restoreState()

    output.parent.mkdir(parents=True, exist_ok=True)
    doc = GuideDoc(str(output), pagesize=(8.5 * inch, 11 * inch),
                   leftMargin=.775 * inch, rightMargin=.775 * inch,
                   topMargin=.78 * inch, bottomMargin=.76 * inch,
                   title='Swift NAVFIT User Guide', author='Swift NAVFIT Project',
                   subject='Operating procedures for Swift NAVFIT r123a')
    doc.build(story, onFirstPage=furniture, onLaterPages=furniture)
    reader = PdfReader(output)
    if len(reader.pages) != len(chapters):
        raise RuntimeError(f'Expected {len(chapters)} pages; got {len(reader.pages)}. Review overflow and contents.')
    for n, page in enumerate(reader.pages):
        text = page.extract_text()
        if len(text) < 200:
            raise RuntimeError(f'Unexpected sparse or blank page {n + 1}')
    print(f'PASS: {len(reader.pages)} pages, chapter boundaries and text extraction; {output}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    root = Path(__file__).resolve().parents[1]
    parser.add_argument('--source', type=Path, default=root / 'docs/Swift_NAVFIT_User_Guide.md')
    parser.add_argument('--output', type=Path, default=root / 'docs/Swift_NAVFIT_User_Guide.pdf')
    args = parser.parse_args()
    build(args.source, args.output)
