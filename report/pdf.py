from reportlab import platypus
from reportlab.lib import styles
from reportlab.lib.units import mm

from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.linecharts import HorizontalLineChart

"""
This module generates the PDF report.
"""

class Pdf(object):
    PARAGRAPH_STYLE = styles.ParagraphStyle(
        "default", fontSize=12, leading=13
    )

    def __init__(self, filename):
        frame = platypus.Frame(
            20 * mm,  20 * mm,  170 * mm, 227 * mm, showBoundary=0,
            topPadding=0, bottomPadding=0, leftPadding=0, rightPadding=0
        )
        page = platypus.PageTemplate("main", frames=[frame],
                                     onPage=self._decorate_page)
        self.doc = platypus.BaseDocTemplate(filename, pageTemplates=[page])
        self.story = []
        self.header = None

    def _decorate_page(self, canvas, document):
        text = canvas.beginText(20 * mm, 277 * mm)
        for line in self.header:
            text.textLine(line)
        canvas.drawText(text)

    def set_header(self, header_lines):
        self.header = header_lines

    def add_paragraph(self, text):
        self.story.append(platypus.Paragraph(text, self.PARAGRAPH_STYLE))
        
    def add_table(self, rows, column_sizes, style=[]):
        real_style = platypus.TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.8, (0, 0, 0)),
            ("BACKGROUND", (0, 0), (-1, 0), (0.8471, 0.8941, 0.7373)),
        ] + style)
        sizes = [s * mm for s in column_sizes]
        self.story.extend([
            platypus.Spacer(1, 3 * mm),
            platypus.Table(rows, colWidths=sizes, style=real_style),
            platypus.Spacer(1, 3 * mm),
        ])

    def save(self):
        assert self.header is not None, \
            "Use #set_header to set header"
        self.doc.build(self.story)