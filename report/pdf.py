from reportlab import platypus
from reportlab.lib import styles
from reportlab.lib.units import mm

from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib import colors

"""
This module generates the PDF report.
"""

class Pdf(object):
    PARAGRAPH_STYLE = styles.ParagraphStyle(
        "default", fontSize=12, leading=13
    )

    def __init__(self, filename):
        self.chart_colors = [
                colors.HexColor("#0000e5"),
                colors.HexColor("#cc0000"),
                colors.HexColor("#00e500"),
                colors.HexColor("#cccc00"),
                colors.HexColor("#ff00ff"),
                colors.HexColor("#5757f0"),
                colors.HexColor("#8f8ff5"),
                colors.HexColor("#1f1feb"),
                colors.HexColor("#c7c7fa"),
                colors.HexColor("#f5c2c2"),
                colors.HexColor("#eb8585"),
                colors.HexColor("#e04747"),
                colors.HexColor("#d60a0a"),
                colors.HexColor("#ff0000"),
                ]

        frame = platypus.Frame(
            20 * mm,  20 * mm,  170 * mm, 239 * mm, showBoundary=0,
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

    @staticmethod
    def _find_min_max(data):
        _min, _max = data[0][0], data[0][0]
        for d in data:
            tmp_min = min(d)
            tmp_max = max(d)
            _min = _min if _min < tmp_min else tmp_min
            _max = _max if _max > tmp_max else tmp_max
        return _min, _max

    def add_line_chart(self, width, height, labels, data, series_names, minv=None, maxv=None):
        n_series = len(data)

        pad = 10

        _min, _max = self._find_min_max(data)
        minv = _min if minv is None else minv
        maxv = _max if maxv is None else maxv

        lc = HorizontalLineChart()
        lc.x = pad * mm
        lc.y = pad * mm
        lc.width = (width - 2 * pad) * mm
        lc.height = (height - 2 * pad) * mm

        lc.categoryAxis.categoryNames = labels
        lc.data = data
        lc.valueAxis.valueMin = minv
        lc.valueAxis.valueMax = maxv

        lc.joinedLines = 1
        lc.categoryAxis.labels.boxAnchor = "n"
        lc.lines.strokeWidth = 2
        for i in range(n_series):
            lc.lines[i].strokeColor = self.chart_colors[i]

        legend = Legend()
        legend.x = lc.width - 20 * mm
        legend.y = ( 5 + 11*n_series ) * mm
        legend.dx = 8
        legend.dy = 8
        legend.fontSize = 9
        legend.boxAnchor = 'nw'
        legend.columnMaximum = 10
        legend.strokeWidth = 1
        legend.strokeColor = colors.black
        legend.deltax = 75
        legend.deltay = 10
        legend.autoXPadding = 5
        legend.yGap = 0
        legend.dxTextSpace = 5
        legend.alignment = 'right'
        legend.subCols.rpad = 30
        legend.colorNamePairs = list(zip(self.chart_colors[:n_series],
            series_names))

        drawing = Drawing(width * mm, height * mm)
        drawing.hAlign = "CENTER"
        drawing.add(lc)
        drawing.add(legend)
        self.story.append(drawing)

    def new_page(self):
        self.story.append(platypus.PageBreak())

    def save(self):
        assert self.header is not None, \
            "Use #set_header to set header"
        self.doc.build(self.story)
