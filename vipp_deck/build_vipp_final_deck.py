"""
Build the VIPP final presentation (Study 1 and Study 2 of the final report)
on top of the VIPP progress-deck template.

The template supplies the AUB slide master, layouts, logos and the four
scenario images (roll plane, a_y profile, top view, bump profile). Its own
slides are removed and replaced by the slides built below. The only other
input is figures/fig6_baseline_time.png (Fig. 6 of the report).

Windows 10 (PowerShell or cmd), from this folder:
    pip install python-pptx
    python build_vipp_final_deck.py --template "VIPP_Progress_Presentation.pptx"

Options:
    --template  path to the progress deck used as template
    --figures   folder holding fig6_baseline_time.png   (default: ./figures)
    --out       output file                          (default: VIPP_Final_Presentation.pptx)
"""

import argparse
import io
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# --------------------------------------------------------------------------
# Deck text that changes between versions
# --------------------------------------------------------------------------
HEADER = "VIPP Final Presentation"
DECK_TITLE = "VIPP FINAL PRESENTATION"
SUBTITLE = "Preview-based active suspension for rollover margin in cornering over road defects"
DATE = "September 2026"
AFFILIATION = "American University of Beirut\nMoataz Maarouf"

# --------------------------------------------------------------------------
# Template style (taken from the progress deck)
# --------------------------------------------------------------------------
MAROON = "840032"
TEXT = "1A1A1A"
MUTED = "6E6E6E"
PINK = "F4E9EE"
PINK_MID = "E6CCD6"
PINK_DARK = "C28099"
GRAY_CARD = "EEEEEE"
WHITE = "FFFFFF"
GRID = "E0E0E0"
AXIS = "BFBFBF"

# one colour per controller, used in every chart
C_P = "A6A6A6"   # passive
C_B = "404040"   # feedback LQR
C_PV = "C28099"  # road preview
C_PA = "840032"  # road + a_y preview

HEAD_FONT = "Helvetica"
BODY_FONT = "Calibri"
EQ_FONT = "Cambria"

# layout names in the template
LAYOUT_TITLE = "3_Title Slide"
LAYOUT_CONTENT = "Title and Content"
LAYOUT_DIVIDER = "1_Title Slide"
LAYOUT_END = "2_Title Slide"

I = Inches


def rgb(hex_str):
    return RGBColor.from_string(hex_str)


# --------------------------------------------------------------------------
# Text helpers
# --------------------------------------------------------------------------
_SUB = re.compile(r"_\{([^}]*)\}|_(\w)")


def _split_sub(s):
    """'m_s a_y h_{cg}' -> [('m', False), ('s', True), ...]; subscripts marked True."""
    out, pos = [], 0
    for m in _SUB.finditer(s):
        if m.start() > pos:
            out.append((s[pos:m.start()], False))
        out.append((m.group(1) if m.group(1) is not None else m.group(2), True))
        pos = m.end()
    if pos < len(s):
        out.append((s[pos:], False))
    return out


def _add_runs(p, spec, size, bold, italic, color, font):
    """spec: str, or list of str / (str, {b, i, color, size, font}) runs."""
    if isinstance(spec, str):
        spec = [spec]
    for item in spec:
        txt, o = (item, {}) if isinstance(item, str) else item
        for piece, sub in _split_sub(txt):
            r = p.add_run()
            r.text = piece
            f = r.font
            f.size = Pt(o.get("size", size))
            f.bold = o.get("b", bold)
            f.italic = o.get("i", italic)
            f.name = o.get("font", font)
            f.color.rgb = rgb(o.get("color", color))
            if sub:
                r._r.get_or_add_rPr().set("baseline", "-25000")


def _bullet(p, indent_in=0.45):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(I(indent_in))))
    pPr.set("indent", str(-int(I(indent_in))))
    bu_font = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    bu_char = pPr.makeelement(qn("a:buChar"), {"char": "•"})
    pPr.append(bu_font)
    pPr.append(bu_char)


def text(slide, x, y, w, h, content, size=28, bold=False, italic=False, color=TEXT,
         font=BODY_FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=0):
    """content: str (\\n starts a new paragraph) or list of paragraph specs."""
    tb = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = I(0.05)
    tf.margin_top = tf.margin_bottom = I(0.03)
    paras = content.split("\n") if isinstance(content, str) else content
    for k, spec in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = align
        if space_after:
            p.space_after = Pt(space_after)
        _add_runs(p, spec, size, bold, italic, color, font)
    return tb


def bullets(slide, x, y, w, h, items, size=28, color=TEXT, space_after=14):
    """items: str, or (bold label, text) tuples."""
    tb = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = I(0.05)
    tf.margin_top = tf.margin_bottom = I(0.03)
    for k, item in enumerate(items):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(space_after)
        _bullet(p)
        if isinstance(item, tuple):
            label, body = item
            _add_runs(p, [(label + " ", {"b": True}), body], size, False, False, color, BODY_FONT)
        else:
            _add_runs(p, item, size, False, False, color, BODY_FONT)
    return tb


def footnote(slide, content):
    text(slide, 0.85, 13.95, 24.9, 0.6, content, size=20, italic=True, color=MUTED)


# --------------------------------------------------------------------------
# Shape helpers
# --------------------------------------------------------------------------
def card(slide, x, y, w, h, fill=PINK, radius=0.05):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, I(x), I(y), I(w), I(h))
    shp.adjustments[0] = radius
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def line(slide, x1, y1, x2, y2, color=TEXT, width=2.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


def number_badge(slide, x, y, d, label, fill=MAROON, color=WHITE, size=26):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, I(x), I(y), I(d), I(d))
    o.fill.solid()
    o.fill.fore_color.rgb = rgb(fill)
    o.line.fill.background()
    o.shadow.inherit = False
    tf = o.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _add_runs(p, label, size, True, False, color, HEAD_FONT)
    return o


def stat(slide, x, y, w, h, big, label, fill=PINK, big_size=66, label_size=26):
    card(slide, x, y, w, h, fill)
    big_h = big_size / 72 * 1.35
    text(slide, x + 0.45, y + 0.3, w - 0.9, big_h, big, size=big_size, bold=True,
         color=MAROON, font=HEAD_FONT)
    text(slide, x + 0.45, y + 0.35 + big_h, w - 0.9, h - big_h - 0.6, label, size=label_size)


def table(slide, x, y, col_w, rows, row_h=0.7, size=24, banded=True, emph=None, first_col_align=PP_ALIGN.LEFT):
    """rows[0] is the header. emph(i, j) -> True draws that cell bold maroon."""
    nr, nc = len(rows), len(rows[0])
    gf = slide.shapes.add_table(nr, nc, I(x), I(y), I(sum(col_w)), I(row_h * nr))
    tbl = gf.table
    for j, cw in enumerate(col_w):
        tbl.columns[j].width = I(cw)
    for i in range(nr):
        tbl.rows[i].height = I(row_h)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = rgb(MAROON)
            else:
                cell.fill.fore_color.rgb = rgb(PINK if (banded and i % 2 == 1) else WHITE)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = I(0.12)
            cell.margin_top = cell.margin_bottom = I(0.03)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = first_col_align if j == 0 else PP_ALIGN.CENTER
            if i == 0:
                _add_runs(p, val, size, True, False, WHITE, BODY_FONT)
            elif emph and emph(i, j):
                _add_runs(p, val, size, True, False, MAROON, BODY_FONT)
            else:
                _add_runs(p, val, size, False, False, TEXT, BODY_FONT)
    return tbl


# --------------------------------------------------------------------------
# Chart helpers (native, editable charts in the template's style)
# --------------------------------------------------------------------------
def _style_chart(ch, title, legend, axis_size=20, vmin=None, vmax=None, major=None,
                 val_fmt=None, cat_title=None):
    ch.font.name = BODY_FONT
    ch.font.size = Pt(18)
    ch.font.color.rgb = rgb(TEXT)
    if title:
        ch.has_title = True
        tf = ch.chart_title.text_frame
        tf.text = title
        for r in tf.paragraphs[0].runs:
            r.font.size = Pt(26)
            r.font.bold = False
            r.font.color.rgb = rgb(TEXT)
    else:
        ch.has_title = False
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(22)
    ca = ch.category_axis
    ca.tick_labels.font.size = Pt(axis_size)
    ca.format.line.color.rgb = rgb(AXIS)
    ca.has_major_gridlines = False
    if cat_title:
        ca.has_title = True
        ca.axis_title.text_frame.text = cat_title
        for r in ca.axis_title.text_frame.paragraphs[0].runs:
            r.font.size = Pt(22)
            r.font.bold = False
            r.font.color.rgb = rgb(MUTED)
    va = ch.value_axis
    va.tick_labels.font.size = Pt(axis_size)
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb(GRID)
    va.format.line.fill.background()
    if vmin is not None:
        va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    if major is not None:
        va.major_unit = major
    if val_fmt:
        va.tick_labels.number_format = val_fmt
        va.tick_labels.number_format_is_linked = False


def bar_chart(slide, x, y, w, h, cats, series, title=None, horizontal=False, fmt="0.00",
              labels=True, label_size=20, legend=True, gap=60, overlap=-10, **axis):
    """series: list of (name, values, colour)."""
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals, _ in series:
        cd.add_series(name, vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    ch = slide.shapes.add_chart(kind, I(x), I(y), I(w), I(h), cd).chart
    _style_chart(ch, title, legend, **axis)
    plot = ch.plots[0]
    plot.gap_width = gap
    plot.overlap = overlap
    for s, (_, _, colour) in zip(plot.series, series):
        s.invert_if_negative = False
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = rgb(colour)
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = fmt
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.font.size = Pt(label_size)
    return ch


def line_chart(slide, x, y, w, h, cats, series, title=None, legend=True, **axis):
    """series: list of (name, values, colour, dashed)."""
    cd = CategoryChartData()
    cd.categories = cats
    for name, vals, _, _ in series:
        cd.add_series(name, vals)
    ch = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, I(x), I(y), I(w), I(h), cd).chart
    _style_chart(ch, title, legend, **axis)
    for s, (_, _, colour, dashed) in zip(ch.plots[0].series, series):
        s.smooth = False
        s.format.line.color.rgb = rgb(colour)
        s.format.line.width = Pt(3)
        if dashed:
            s.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        s.marker.style = XL_MARKER_STYLE.CIRCLE
        s.marker.size = 9
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = rgb(colour)
        s.marker.format.line.color.rgb = rgb(colour)
    return ch


def point_label(series, idx, label, position=XL_LABEL_POSITION.ABOVE, size=22, color=MAROON):
    dl = series.points[idx].data_label
    dl.has_text_frame = True
    dl.text_frame.text = label
    for r in dl.text_frame.paragraphs[0].runs:
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = rgb(color)
    dl.position = position


# --------------------------------------------------------------------------
# Slide scaffolding
# --------------------------------------------------------------------------
class Deck:
    def __init__(self, template, figures):
        self.prs = Presentation(template)
        self.figures = Path(figures)
        self.layouts = {l.name: l for l in self.prs.slide_layouts}
        for name in (LAYOUT_TITLE, LAYOUT_CONTENT, LAYOUT_DIVIDER, LAYOUT_END):
            if name not in self.layouts:
                sys.exit(f"Template has no layout named '{name}'. Found: {list(self.layouts)}")
        self.images = self._collect_template_images()
        self._remove_template_slides()

    def _collect_template_images(self):
        found = {}
        for s in self.prs.slides:
            for shp in s.shapes:
                if shp.shape_type == 13:  # picture
                    descr = shp._element.nvPicPr.cNvPr.get("descr", "")
                    found[descr] = shp.image.blob
        return found

    def _remove_template_slides(self):
        sld_ids = self.prs.slides._sldIdLst
        for sld_id in list(sld_ids):
            self.prs.part.drop_rel(sld_id.rId)
            sld_ids.remove(sld_id)

    def _new(self, layout, notes):
        s = self.prs.slides.add_slide(self.layouts[layout])
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def content(self, title, notes=None):
        s = self._new(LAYOUT_CONTENT, notes)
        text(s, 17.45, 0.72, 7.39, 0.55, HEADER, size=23, font=HEAD_FONT, align=PP_ALIGN.RIGHT)
        text(s, 0.85, 2.55, 24.9, 1.0, title, size=50, bold=True, color=MAROON, font=HEAD_FONT)
        return s

    def divider(self, kicker, title, sub, notes=None):
        s = self._new(LAYOUT_DIVIDER, notes)
        text(s, 2.0, 3.1, 22.67, 0.9, kicker, size=36, bold=True, color=PINK_MID,
             font=HEAD_FONT, align=PP_ALIGN.CENTER)
        text(s, 2.0, 4.2, 22.67, 1.4, title, size=64, bold=True, color=WHITE,
             font=HEAD_FONT, align=PP_ALIGN.CENTER)
        text(s, 4.0, 6.2, 18.67, 2.4, sub, size=32, color=WHITE, align=PP_ALIGN.CENTER)
        return s

    def picture(self, slide, key, x, y, w=None, h=None):
        """key: alt text of a template image, or a file name in the figures folder."""
        blob = self.images.get(key)
        path = self.figures / key
        if blob is not None:
            src = io.BytesIO(blob)
        elif path.exists():
            src = str(path)
        else:
            print(f"warning: image '{key}' not found in template or {self.figures}", file=sys.stderr)
            card(slide, x, y, w or 8, h or 5, GRAY_CARD)
            text(slide, x + 0.3, y + 0.3, (w or 8) - 0.6, 1.5, f"missing image: {key}", size=22, color=MUTED)
            return None
        return slide.shapes.add_picture(src, I(x), I(y), I(w) if w else None, I(h) if h else None)


# --------------------------------------------------------------------------
# Slides
# --------------------------------------------------------------------------
def slide_title(d):
    s = d._new(LAYOUT_TITLE, "Final results of VIPP. Two studies: what feedback active suspension does "
               "on a straight rough road with defects, and whether preview of the road and of the "
               "planned lateral acceleration protects the rollover margin when a defect is met in a corner.")
    text(s, 1.93, 6.1, 22.5, 0.9, DECK_TITLE, size=45, bold=True, color=WHITE, font=HEAD_FONT)
    text(s, 1.93, 7.2, 22.5, 1.3, SUBTITLE, size=34, color=WHITE, font=HEAD_FONT)
    text(s, 1.93, 9.05, 10.0, 0.7, DATE, size=30, color=WHITE, font=HEAD_FONT)
    line(s, 1.95, 9.95, 13.8, 9.95, WHITE, 1.5)
    text(s, 1.93, 10.15, 15.0, 1.4, AFFILIATION, size=27, color=WHITE, font=HEAD_FONT)


def slide_outline(d):
    s = d.content("OUTLINE")
    items = [
        ("01.", "Setup: model, metric, scenario"),
        ("02.", "Study 1: feedback in straight-line driving"),
        ("03.", "Study 2: cornering limits and controllers"),
        ("04.", "Study 2: baseline and statistical results"),
        ("05.", "Study 2: generalisation, severity, actuator"),
        ("06.", "Discussion and conclusion"),
    ]
    for k, (num, label) in enumerate(items):
        x = 2.8 if k < 3 else 15.3
        y = 4.3 + (k % 3) * 2.9
        text(s, x, y, 1.8, 0.9, num, size=44, bold=True, color=MAROON, font=HEAD_FONT)
        text(s, x, y + 0.95, 10.5, 1.4, label, size=36, font=HEAD_FONT)


def slide_setup_model(d):
    s = d.content("1. SETUP: MODEL AND METRIC",
                  "Research question and the model in one slide. Seven-degree-of-freedom full car, "
                  "tyres that cannot pull, travel and force limits. Lateral acceleration is an input from "
                  "the path planner: it both creates the corner and is the signal the preview controller "
                  "can use. The rollover metric is the load-transfer ratio; LTR = 1 means both inner "
                  "wheels are unloaded, which is necessary for untripped rollover but is not a rollover.")
    d.picture(s, "roll_plane.png", 0.85, 3.9, w=11.0)
    text(s, 13.2, 3.9, 12.6, 0.7, "Research question", size=32, bold=True, color=MAROON, font=HEAD_FONT)
    text(s, 13.2, 4.6, 12.6, 1.4,
         "How much rollover margin can active suspension protect when a defect has to be ridden over in a corner?",
         size=31)
    bullets(s, 13.2, 6.0, 12.6, 3.8, [
        ("7-DOF full car:", "heave, pitch, roll and four wheel hops; 14 states, 1500 kg sprung mass"),
        ("Tyres cannot pull:", "F_n = 0 means the wheel has left the road"),
        ("Limits:", "±100 mm travel with bump stops, ±15 kN actuator force"),
        ("Corner input:", "a_y from the path planner; no lateral or yaw dynamics"),
    ], size=29, space_after=8)
    text(s, 13.2, 9.95, 12.6, 0.7, "Metric: load-transfer ratio", size=32, bold=True, color=MAROON, font=HEAD_FONT)
    text(s, 13.2, 10.7, 12.6, 0.7, "LTR  =  (ΣF_{outer} − ΣF_{inner}) / (ΣF_{outer} + ΣF_{inner})",
         size=30, bold=True, font=EQ_FONT)
    x0, x6, x1 = 13.6, 20.0, 24.8
    rect(s, x0, 11.75, x6 - x0, 0.45, PINK_MID)
    rect(s, x6, 11.75, x1 - x6, 0.45, PINK_DARK)
    for xv, label in ((x0, "0\nlevel"), (x6, "0.6\nincipient"), (x1, "1.0\ntrip condition")):
        line(s, xv, 11.55, xv, 12.4, MAROON, 2.0)
        text(s, xv - 1.6, 12.45, 3.2, 1.2, label, size=24, bold=True, color=MAROON, align=PP_ALIGN.CENTER)
    footnote(s, "Trip condition = both inner wheels unloaded: necessary for untripped rollover, not a rollover event.")


def slide_setup_scenario(d):
    s = d.content("1. SETUP: ROAD, CORNER, DEFECT",
                  "Class C road at 50 km/h. Nine-second left turn: straight, 1.5 s ramp in, 3 s steady "
                  "hold, ramp out. Baseline 0.5 g at 39 m radius; sweeps go from 0.3 to 0.8 g. A 44 mm, "
                  "0.5 m bump sits under the inner front wheel at the start of the hold. The rear axle "
                  "runs 103 mm inside the front, so the footprint is widened and the inner rear also hits "
                  "it about 0.2 s later. The twin-defect case adds a second bump one wheelbase later.")
    d.picture(s, "ay_profile.png", 0.85, 3.9, w=12.2)
    d.picture(s, "top_view.png", 13.5, 3.9, w=12.2)
    d.picture(s, "bump_profile.png", 0.85, 9.7, w=6.3)
    bullets(s, 7.6, 9.8, 18.2, 4.0, [
        ("Road:", "ISO 8608 Class C, partially correlated left/right tracks, averaged over a 0.18 m tyre patch"),
        ("Corner:", "9 s manoeuvre with a 3 s steady hold; baseline 0.5 g, R = 39 m, 50 km/h"),
        ("Defect:", "44 mm × 0.5 m bump under the inner front wheel at the start of the hold; the inner rear meets it 0.2 s later"),
        ("Twin defect:", "a second bump one wheelbase later unloads both inner wheels together"),
    ], size=29, space_after=8)


def slide_study1_passive(d):
    s = d.content("2. STUDY 1: PASSIVE CAR ON A DEFECT ROAD",
                  "Straight line, 50 km/h, 60 s. On Class C alone the tyres stay loaded: minimum normal "
                  "force 1048 N, DLC 0.181, no lift. Adding discrete defects barely moves the comfort "
                  "averages but changes road holding qualitatively: nine lift events, 0.140 s off the "
                  "road, travel doubles. Averaged metrics understate the defect.")
    rows = [
        ["Passive, 50 km/h, 60 s", "Class C", "+ defects", "change"],
        ["W_k heave RMS [m/s²]", "0.679", "0.723", "+6.5 %"],
        ["VDV ratio", "1.31", "1.40", "+7.4 %"],
        ["Dynamic load coefficient", "0.181", "0.193", "+6.4 %"],
        ["Crest factor", "3.86", "5.70", "+47.6 %"],
        ["Suspension travel, max [mm]", "32.4", "70.1", "+116 %"],
        ["Minimum normal force [N]", "1048", "0", "−100 %"],
        ["Time off the road [s]", "0.000", "0.140", "—"],
        ["Lift events", "0", "9", "—"],
        ["LTR, max", "0.232", "0.537", "—"],
    ]
    table(s, 0.85, 3.9, [6.6, 2.4, 2.6, 2.6], rows, row_h=0.82, size=26,
          emph=lambda i, j: i >= 4 and j >= 2)
    stat(s, 15.8, 3.9, 10.0, 3.6, "+6.5 %",
         "weighted heave acceleration with defects: comfort stays in the same ISO 2631-1 band", label_size=29)
    stat(s, 15.8, 7.9, 10.0, 3.6, "9 lifts",
         "0.140 s off the road, longest 20 ms; minimum tyre load falls from 1048 N to 0", label_size=29)
    text(s, 15.8, 11.9, 10.0, 1.8, "Averages barely move; road holding fails. The safety case sits at the discrete defect.",
         size=30, bold=True, color=MAROON)
    footnote(s, "Baseline vehicle of Table I, ISO 8608 Class C. Static load 3973 N per corner.")


def slide_study1_selection(d):
    s = d.content("2. STUDY 1: PASSIVE TRADE-OFF AND LQR TUNING",
                  "First the passive design itself: 25 spring/damper tunings, 12 on the Pareto front. "
                  "No tuning reaches the 'not uncomfortable' band. The balanced point is the mechanical "
                  "reference, so passive and active differ only by the actuator force. Then 90 LQR "
                  "candidates were screened with a strict acceptance rule; the selected gain is "
                  "alpha = 0.05, rho = 300.")
    card(s, 0.85, 3.9, 12.2, 9.7, PINK)
    text(s, 1.35, 4.25, 11.2, 0.8, "Passive design sweep", size=34, bold=True, color=MAROON, font=HEAD_FONT)
    text(s, 1.35, 5.1, 11.2, 1.8,
         "25 designs: ride frequency 1.00–1.60 Hz × damping ratio 0.15–0.45. 12 lie on the comfort / road-holding Pareto front.",
         size=29)
    table(s, 1.35, 7.1, [5.6, 2.8, 2.8], [
        ["Pareto endpoint", "W_k [m/s²]", "DLC"],
        ["Best comfort", "0.512", "0.249"],
        ["Best road holding", "0.760", "0.177"],
    ], row_h=0.8, size=27, banded=False)
    bullets(s, 1.35, 9.8, 11.2, 3.7, [
        "Endpoint to endpoint: DLC changes 40 %, working space 106 %",
        "No tuning reaches the ISO 2631-1 “not uncomfortable” band",
        "Reference for the active car: balanced point, 1.00 Hz, ζ = 0.30",
    ], size=28, space_after=10)

    card(s, 13.6, 3.9, 12.2, 9.7, GRAY_CARD)
    text(s, 14.1, 4.25, 11.2, 0.8, "Active LQR: selection rule", size=34, bold=True, color=MAROON, font=HEAD_FONT)
    steps = [
        (5.25, 1.4, "90 candidate gains: comfort vs road-holding priority α × actuator-effort penalty ρ"),
        (6.85, 1.4, "Each simulated on 9 nonlinear tuning cases, actuator saturated at ±15 kN"),
        (8.45, 2.7, "Admitted only if: travel limit held, no lift on Class C, ≤ 20 ms extra lift on defect "
                    "roads, ≤ 1 % saturation, comfort and road holding within 5 % of passive"),
    ]
    for k, (y, h, label) in enumerate(steps, start=1):
        number_badge(s, 14.1, y, 0.8, str(k))
        text(s, 15.2, y + 0.02, 10.1, h, label, size=28)
    card(s, 14.1, 11.55, 11.2, 1.55, MAROON)
    text(s, 14.1, 11.55, 11.2, 1.55, "Selected gain:  α = 0.05,  ρ = 300", size=32, bold=True,
         color=WHITE, font=HEAD_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_study1_feedback(d):
    s = d.content("2. STUDY 1: WHAT FEEDBACK ACHIEVES",
                  "Matched passive against the selected LQR on six road scenarios, ten roads each. "
                  "Gains are fractions of a percent; the actuator barely works. Wheel lift at the defects "
                  "is identical to the millisecond in every scenario: state feedback can only respond "
                  "after the wheel is struck. That is the limitation preview removes.")
    bar_chart(s, 0.85, 3.9, 14.3, 9.7,
                   ["Class C", "Potholes", "Asymmetric", "Speed bumps", "Potholes + asym.", "All defects"],
                   [("Passive", [0.000, 0.117, 0.023, 0.001, 0.143, 0.144], C_P),
                    ("Active LQR", [0.000, 0.117, 0.023, 0.001, 0.143, 0.144], C_PA)],
                   title="Time off the road [s], mean of 10 roads", fmt="0.000", label_size=18,
                   vmin=0, vmax=0.16, major=0.04, val_fmt="0.00", axis_size=21)
    positions = [(15.7, 3.9), (21.0, 3.9), (15.7, 7.2), (21.0, 7.2)]
    tiles = [("−0.4 %", "W_k heave RMS"), ("−0.3 %", "dynamic load coefficient"),
             ("+4 %", "pitch and roll RMS"), ("0.02 kN", "actuator RMS (limit 15 kN)")]
    for (x, y), (big, label) in zip(positions, tiles):
        stat(s, x, y, 4.9, 3.0, big, label, big_size=52, label_size=26)
    text(s, 15.7, 10.6, 10.2, 3.0,
         "No candidate prevents wheel lift at a defect. State feedback acts on the car’s own motion, "
         "so it responds only after the wheel is struck.",
         size=30, bold=True, color=MAROON)
    footnote(s, "Same mechanics, roads and defect placements; only the actuator force differs. "
                "Gains up to ≈ 11 % on W_k exist but raise DLC or attitude above passive.")


def slide_study1_carry(d):
    s = d.content("2. STUDY 1: WHAT CARRIES FORWARD",
                  "Three conclusions. On a maintained road the passive car keeps its tyres loaded, so "
                  "the safety case is at defects. Feedback cannot stop the lift at a defect, for a "
                  "structural reason: that motivates preview. Straight-line roll has no safety meaning; "
                  "in a corner roll and lateral load transfer are the rollover criterion: that motivates "
                  "the cornering scenario of Study 2.")
    cards = [
        ("01", "Class C: tyres stay loaded",
         "On a maintained road, active control is a comfort and attitude tool. The safety case has to be made at discrete defects.",
         "→ focus on the defect"),
        ("02", "Feedback cannot stop the lift",
         "Lift time is unchanged to the millisecond at every weighting tested: a structural limit, not a tuning one.",
         "→ motivates preview"),
        ("03", "Roll matters only in a corner",
         "Straight-line roll is small and road-induced. In a corner, roll and lateral load transfer set the rollover criterion.",
         "→ motivates cornering"),
    ]
    for k, (num, head, body, tag) in enumerate(cards):
        x = 0.85 + k * 8.4
        card(s, x, 3.9, 8.0, 7.9, PINK if k != 1 else PINK_MID)
        text(s, x + 0.45, 4.2, 7.1, 1.2, num, size=60, bold=True, color=MAROON, font=HEAD_FONT)
        text(s, x + 0.45, 5.5, 7.1, 1.3, head, size=33, bold=True)
        text(s, x + 0.45, 6.75, 7.1, 3.9, body, size=30)
        text(s, x + 0.45, 10.75, 7.1, 0.7, tag, size=29, bold=True, color=MAROON)
    text(s, 0.85, 12.4, 24.9, 0.9, "Study 2 asks the cornering question with preview control.",
         size=34, bold=True, color=MAROON, font=HEAD_FONT)


def slide_corner_limits(d):
    s = d.content("3. STUDY 2: WHAT CORNERING ALONE CAN DO",
                  "Two limits frame the study. Cornering alone cannot reach LTR = 1 on this car: that "
                  "needs 1.61 g and the tyres saturate first. And in a steady corner the actuator forces "
                  "are internal to the moment balance; only the CG shift from body lean is reachable, "
                  "about 6 % of the load transfer. So the target is the transient peak at the defect.")
    text(s, 0.85, 3.9, 12.6, 0.8, "Steady corner: the car slides first", size=34, bold=True,
         color=MAROON, font=HEAD_FONT)
    text(s, 0.85, 4.75, 12.4, 2.6,
         "Linear statics: LTR = 0.6 at 0.97 g and LTR = 1 at 1.61 g, both above the tyre grip limit. "
         "In simulation the passive car crosses 0.6 only at 0.8 g and never trips.", size=31)
    text(s, 0.85, 7.55, 12.6, 0.8, "What any suspension can change", size=34, bold=True,
         color=MAROON, font=HEAD_FONT)
    text(s, 0.85, 8.45, 12.6, 0.9,
         "(F_{out} − F_{in}) T/2 = m_s a_y h_{cg} + m_u a_y h_u + m_s g h_R sin φ",
         size=30, bold=True, font=EQ_FONT)
    text(s, 0.85, 9.6, 12.4, 3.0,
         "Actuator forces act between sprung and unsprung masses, so they drop out of this balance. "
         "Only the last term, the CG shift from body lean, is reachable.", size=31)
    stat(s, 14.0, 3.9, 11.8, 4.4, "1.61 g",
         "lateral acceleration for LTR = 1 in a steady corner; tyres give about 0.8–0.9 g", label_size=28, big_size=72)
    stat(s, 14.0, 8.7, 11.8, 4.4, "≈ 6 %",
         "of steady load transfer any suspension can remove at 0.5 g: ΔLTR ≈ 0.017 (0.307 → 0.288 without lean)",
         label_size=28, big_size=72)
    text(s, 14.0, 13.3, 11.8, 0.7, "So target the transient peak at the defect.", size=31,
         bold=True, color=MAROON)


def slide_controllers(d):
    s = d.content("3. STUDY 2: CONTROLLERS COMPARED",
                  "Five controllers. Passive is the ride-frequency design. A and B are output-weighted "
                  "LQR on the full state with Bryson-normalised, rollover-led weights; B adds LTR as an "
                  "output. Pv adds LQ preview of 1.04 m of road ahead; Pa also previews the planned "
                  "lateral acceleration, which an automated vehicle has from its planner. The feedback "
                  "part of the preview controllers equals the ordinary LQR gain.")
    ctrl = [
        ("P", "Passive", "ride-frequency springs and dampers", "BASELINE", GRAY_CARD),
        ("A", "Roll LQR", "output-weighted feedback, roll angle weighted", "FEEDBACK", PINK),
        ("B", "Roll + LTR LQR", "A plus a load-transfer output in the cost", "FEEDBACK", PINK),
        ("Pv", "Road preview", "B plus 1.04 m of road ahead (75 ms)", "PREVIEW", PINK_MID),
        ("Pa", "Road + a_y preview", "Pv plus the planned lateral acceleration", "PREVIEW", PINK_MID),
    ]
    for k, (letter, name, desc, tag, fill) in enumerate(ctrl):
        x = 0.85 + k * 5.05
        card(s, x, 4.1, 4.72, 5.4, fill)
        text(s, x + 0.3, 4.35, 4.12, 1.3, letter, size=60, bold=True, color=MAROON, font=HEAD_FONT)
        text(s, x + 0.3, 5.75, 4.12, 1.2, name, size=27, bold=True)
        text(s, x + 0.3, 7.05, 4.12, 1.6, desc, size=26)
        text(s, x + 0.3, 8.75, 4.12, 0.6, tag, size=20, bold=True, color=MUTED, font=HEAD_FONT)
    bullets(s, 0.85, 10.0, 25.0, 3.8, [
        ("Output-weighted LQR:", "body accelerations, travel, dynamic tyre load, roll angle and LTR as outputs; "
                                 "Bryson-normalised, rollover-led weights; discretised at 1 kHz"),
        ("LQ preview:", "state augmented with a shift register of future disturbance samples, "
                        "u_k = −K_x x_k − Σ K_{ff,j} w_{k+j}"),
        ("Structure:", "K_x equals the ordinary LQR gain; preview gains follow a recursion and decay with lead "
                       "time; the tail beyond the horizon is summed in closed form (matters for the steady a_y)"),
    ], size=29, space_after=10)


def slide_horizon(d):
    s = d.content("3. STUDY 2: CHOOSING THE PREVIEW HORIZON",
                  "Sweep of look-ahead distance at the baseline case. The LQ cost saturates early: "
                  "97 % of the reduction by 0.28 m. Peak LTR keeps improving to about 1.04 m. Since the "
                  "reported metric is the transient margin, 1.04 m is used: 75 ms at 50 km/h, about one "
                  "wheel-hop period. A clairvoyant controller given the whole road does no better and "
                  "uses no more force: the limit is information, not actuator authority.")
    cats = ["0", "0.07", "0.14", "0.28", "0.49", "0.69", "1.04", "1.39", "2.78"]
    ch = line_chart(s, 0.85, 3.9, 12.3, 6.7, cats,
                    [("Cost reduction captured [%]", [0, 30.0, 71.0, 96.9, 97.5, 98.2, 99.2, 99.6, 99.9], C_PA, False)],
                    title="LQ cost reduction captured [%]", legend=False, vmin=0, vmax=110, major=20,
                    cat_title="look-ahead distance [m]", axis_size=21)
    point_label(ch.plots[0].series[0], 3, "0.28 m", XL_LABEL_POSITION.ABOVE)
    ch = line_chart(s, 13.5, 3.9, 12.3, 6.7, cats,
                    [("Peak LTR", [0.796, 0.718, 0.635, 0.589, 0.544, 0.537, 0.500, 0.494, 0.489], C_PA, False)],
                    title="Peak LTR at the baseline case", legend=False, vmin=0.4, vmax=0.85, major=0.1,
                    val_fmt="0.0", cat_title="look-ahead distance [m]", axis_size=21)
    point_label(ch.plots[0].series[0], 6, "1.04 m", XL_LABEL_POSITION.BELOW)
    tiles = [("0.28 m", "cost knee: 96.9 % of the LQ cost reduction"),
             ("1.04 m", "where peak LTR stops improving; used (75 ms at 50 km/h)"),
             ("0.499", "clairvoyant peak LTR: the gap is information, not force")]
    for k, (big, label) in enumerate(tiles):
        stat(s, 0.85 + k * 8.4, 10.9, 8.0, 2.8, big, label, big_size=50, label_size=27)


def slide_baseline(d):
    s = d.content("4. STUDY 2: THE BASELINE CASE",
                  "One road at 0.5 g. Passive peaks at 0.742. Both feedback controllers are worse, and B "
                  "leans further than passive. Road preview brings the peak to 0.500; adding the planned "
                  "lateral acceleration gives 0.457 and holds the body within 0.6 degrees. The largest "
                  "peak comes when the inner rear wheel meets the bump, 0.2 s after the front. Preview "
                  "retracts the wheel before impact, cutting the tyre impact force from 10.0 to 6.4 kN.")
    d.picture(s, "fig6_baseline_time.png", 0.85, 3.9, w=13.6)
    text(s, 0.85, 11.65, 13.6, 1.0, "a_y, body roll and LTR; the dashed line marks the inner front wheel meeting the bump.",
         size=24, italic=True, color=MUTED)
    table(s, 14.95, 3.9, [2.0, 3.0, 2.7, 3.2], [
        ["0.5 g", "peak LTR", "roll [°]", "past 0.6 [ms]"],
        ["P", "0.742", "3.04", "33"],
        ["A", "0.865", "1.28", "50"],
        ["B", "0.822", "5.44", "22"],
        ["Pv", "0.500", "5.13", "0"],
        ["Pa", "0.457", "0.61", "0"],
    ], row_h=0.75, size=26, emph=lambda i, j: i == 5)
    bullets(s, 14.95, 8.75, 10.9, 5.1, [
        "Feedback raises the peak; B leans further than passive (5.4° vs 3.0°)",
        "Road + a_y preview: peak 0.457, body held within 0.6°",
        "Largest peak comes 0.2 s after the front strike, when the inner rear meets the bump",
        "Preview retracts the wheel before impact: tyre impact force 10.0 → 6.4 kN",
    ], size=28, space_after=8)
    footnote(s, "One road realisation, 50 km/h, R = 39 m, 44 mm × 0.5 m bump under the inner front wheel at the start of the hold.")


def slide_statistics(d):
    s = d.content("4. STUDY 2: OVER 25 ROAD REALISATIONS",
                  "Twenty-five roads, differences paired by road. Road plus planned-a_y preview lowers "
                  "peak LTR on every road at both levels: by 0.236 at 0.5 g and 0.194 at 0.8 g. Road "
                  "preview alone helps but leaves the body leaning 5 to 8 degrees and rides the bump "
                  "stops at 0.8 g. Feedback gives no significant gain at 0.5 g and at 0.8 g reaches the "
                  "trip condition on 11 of 25 roads, where passive does on one.")
    bar_chart(s, 0.85, 3.9, 13.3, 9.8, ["P", "B", "Pv", "Pa"],
              [("0.5 g", [0.714, 0.688, 0.524, 0.478], PINK_DARK),
               ("0.8 g", [0.870, 0.926, 0.771, 0.676], MAROON)],
              title="Mean peak LTR over 25 roads (lower is better)", fmt="0.00",
              vmin=0, vmax=1.0, major=0.2, val_fmt="0.0", axis_size=24)
    table(s, 14.6, 3.9, [1.6, 1.6, 5.0, 3.0], [
        ["a_y", "Ctrl", "Δ vs passive [95 % CI]", "better on"],
        ["0.5 g", "B", "−0.027 [−0.063, +0.010]", "16/25"],
        ["0.5 g", "Pv", "−0.190 [−0.210, −0.171]", "25/25"],
        ["0.5 g", "Pa", "−0.236 [−0.256, −0.216]", "25/25"],
        ["0.8 g", "B", "+0.056 [+0.013, +0.099]", "7/25"],
        ["0.8 g", "Pv", "−0.099 [−0.130, −0.068]", "22/25"],
        ["0.8 g", "Pa", "−0.194 [−0.218, −0.170]", "25/25"],
    ], row_h=0.74, size=24, emph=lambda i, j: i in (3, 6))
    bullets(s, 14.6, 9.4, 11.2, 4.4, [
        ("Pa", "lowers the peak on every road at both levels"),
        ("Pv", "also helps but leaves 5–8° of lean; at 0.8 g it rides the bump stops for 2.9 s of 9 s"),
        ("Feedback:", "no significant gain at 0.5 g; at 0.8 g it reaches LTR = 1 on 11 of 25 roads (passive: 1)"),
    ], size=28, space_after=10)
    footnote(s, "ISO 8608 Class C, 50 km/h, 44 mm bump under the inner front wheel. Differences paired by road.")


def slide_robustness(d):
    s = d.content("4. STUDY 2: IS THE GAIN REAL AND ROBUST?",
                  "LTR is a ratio, so a controller could lower it by unloading the whole car. It does "
                  "not: preview raises the minimum total tyre load, lowers the roll rate, and the inner "
                  "wheel keeps 500 N where passive reaches zero. Keeping the body level is not the same "
                  "as keeping the tyres loaded: A holds 1.3 degrees yet raises the peak. A higher LTR "
                  "weight helps preview and hurts feedback. The preview gain holds over CG heights "
                  "0.30 to 0.60 m.")
    card(s, 0.85, 3.9, 8.0, 9.8, PINK)
    text(s, 1.3, 4.2, 7.1, 0.8, "Is the lower LTR real?", size=30, bold=True, color=MAROON, font=HEAD_FONT)
    text(s, 1.3, 5.05, 7.1, 1.3, "LTR is a ratio: it could fall if the whole car unloads.", size=28, italic=True)
    bullets(s, 1.3, 6.6, 7.1, 6.9, [
        "Preview raises the minimum total tyre load and lowers the roll rate",
        "Minimum inner-wheel load: 0 N (passive) → 500 N",
        "A level body is not loaded tyres: A holds 1.3° yet raises the peak",
    ], size=29, space_after=14)
    line_chart(s, 9.25, 3.9, 8.1, 7.3, ["4", "8", "16", "32"],
               [("B (feedback)", [0.767, 0.758, 0.822, 1.000], C_B, True),
                ("Pa (preview)", [0.520, 0.489, 0.457, 0.425], C_PA, False)],
               title="Peak LTR vs LTR weight", vmin=0.3, vmax=1.1, major=0.2, val_fmt="0.0",
               cat_title="LTR weight", axis_size=21)
    line_chart(s, 17.7, 3.9, 8.1, 7.3, ["0.30", "0.40", "0.508", "0.55", "0.60"],
               [("P (passive)", [0.641, 0.691, 0.742, 0.766, 0.788], C_P, True),
                ("Pa (preview)", [0.313, 0.382, 0.457, 0.487, 0.522], C_PA, False)],
               title="Peak LTR vs CG height", vmin=0.2, vmax=0.9, major=0.1, val_fmt="0.0",
               cat_title="CG height [m]", axis_size=21)
    text(s, 9.25, 11.45, 8.1, 2.3,
         "Raising the LTR weight improves preview (0.520 → 0.425) but drives feedback to the trip condition.",
         size=27)
    text(s, 17.7, 11.45, 8.1, 2.3,
         "SSF 2.58 → 1.29: preview still lowers the peak by 0.27–0.33 (redesigned at each height).",
         size=27)


def slide_generalisation(d):
    s = d.content("5. STUDY 2: GENERALISATION OVER 55 CELLS",
                  "The same design procedure, applied without re-tuning: build the model from the vehicle "
                  "parameters, same weights, 1.04 m of road plus the planned a_y. Six vehicles, speeds, "
                  "radii, defect type, length, location, height and timing. 51 of 55 cells pass. Preview "
                  "lowered peak LTR on 10 of 10 roads in every cell and never reached the trip condition. "
                  "The four failures have clear causes, listed on the right.")
    tiles = [("51 / 55", "scenario cells pass"),
             ("10 / 10", "roads improved in every cell"),
             ("0", "cells where preview trips (passive 11, feedback 15)"),
             ("0.10–0.37", "reduction in peak LTR across cells")]
    for k, (big, label) in enumerate(tiles):
        stat(s, 0.85 + k * 6.34, 3.9, 5.94, 2.75, big, label, big_size=50, label_size=25)
    rows = [
        ["Axis", "cells", "Δ peak LTR", "pass"],
        ["Vehicle (six, redesigned)", "12", "−0.32 to −0.15", "12/12"],
        ["Vehicle (mismatched gains)", "2", "−0.10, −0.04", "0/2"],
        ["Speed at fixed radius, 30–60 km/h", "4", "−0.37 to −0.22", "4/4"],
        ["Speed at fixed a_y, 30–70 km/h", "3", "−0.32 to −0.23", "3/3"],
        ["Radius, 100–25 m", "4", "−0.29 to −0.20", "4/4"],
        ["Defect type and length", "12", "−0.34 to −0.12", "12/12"],
        ["Defect location", "8", "−0.24 to −0.10", "8/8"],
        ["Defect height, 80 and 150 mm", "4", "−0.19 to −0.18", "3/4"],
        ["Point in the corner", "6", "−0.27 to −0.19", "6/6"],
    ]
    table(s, 0.85, 7.0, [7.6, 1.6, 4.0, 2.2], rows, row_h=0.64, size=24,
          emph=lambda i, j: j == 3 and i in (2, 8))
    text(s, 16.8, 7.0, 9.0, 0.7, "The four failures", size=30, bold=True, color=MAROON, font=HEAD_FONT)
    bullets(s, 16.8, 7.75, 9.0, 3.9, [
        ("SUV on the sedan’s gains (2):", "peak still falls on every road, but the lean stays near passive"),
        ("Full-size van, 0.5 g:", "misses the roll criterion by 0.001"),
        ("150 mm bump:", "exceeds travel; bump stops for 7 ms, roll rate 30°/s"),
    ], size=26, space_after=8)
    text(s, 16.8, 11.4, 9.0, 0.6, "Trends", size=28, bold=True, color=MAROON, font=HEAD_FONT)
    text(s, 16.8, 12.05, 9.0, 1.7,
         "Gain shrinks as a_y rises (steady share grows) and with speed at fixed a_y (fixed distance = shorter time).",
         size=25)
    footnote(s, "Pass: lower peak LTR on ≥ 90 % of roads, no trip where passive has none, roll ≤ 25 % of passive, "
                "no bump-stop contact. 10 roads per cell.")


def slide_severity(d):
    s = d.content("5. STUDY 2: DEFECT SEVERITY",
                  "Defect height from 44 to 150 mm at 0.5 g. With a single bump, preview never reaches "
                  "the trip condition at any height tested; passive first reaches it at 80 mm at 0.8 g. "
                  "The twin bump unloads both inner wheels together: passive trips on every road at "
                  "every height, preview avoids it at 44 mm and cuts the time at LTR = 1 by four to ten "
                  "times above that.")
    heights = ["44", "60", "80", "100", "120", "150"]
    line_chart(s, 0.85, 3.9, 12.3, 7.4, heights,
               [("P", [0.733, 0.834, 0.865, 0.891, 0.918, 0.957], C_P, True),
                ("Pv", [0.561, 0.595, 0.661, 0.693, 0.720, 0.755], C_PV, False),
                ("Pa", [0.513, 0.542, 0.625, 0.653, 0.679, 0.720], C_PA, False)],
               title="Single bump: worst-road peak LTR", vmin=0.4, vmax=1.0, major=0.1, val_fmt="0.0",
               cat_title="bump height [mm]", axis_size=21)
    bar_chart(s, 13.5, 3.9, 12.3, 7.4, heights,
              [("P", [34, 49, 65, 81, 100, 141], C_P),
               ("Pv", [0, 1, 12, 21, 29, 42], C_PV),
               ("Pa", [0, 1, 11, 19, 26, 38], C_PA)],
              title="Twin bump: time at LTR = 1 [ms]", fmt="0", label_size=17, gap=50, overlap=0,
              vmin=0, vmax=160, major=40, val_fmt="0", cat_title="bump height [mm]", axis_size=21)
    bullets(s, 0.85, 11.6, 24.9, 2.2, [
        ("Single bump:", "preview never reaches the trip condition up to 150 mm; passive first does at 80 mm (0.8 g)"),
        ("Twin bump:", "passive trips on every road at every height; preview avoids it at 44 mm and cuts the dwell 4–10× above"),
    ], size=28, space_after=8)
    footnote(s, "0.5 g, mean of five road realisations. Twin bump: second identical bump one wheelbase later, a constructed worst case.")


def slide_actuator(d):
    s = d.content("5. STUDY 2: ACTUATOR REQUIREMENT",
                  "The nominal controller with a first-order lag and slew limit, not redesigned. The peak-"
                  "LTR gain survives to about 15 Hz and is essentially complete by 30 Hz. At 10 Hz and "
                  "below the force arrives after the 11.2 Hz wheel-hop event, so it is worse than "
                  "passive. 300 kN/s slew keeps most of the gain; 150 kN/s removes it. Removing the lean "
                  "is a low-frequency action and survives at every bandwidth.")
    cases = ["Ideal", "30 Hz", "15 Hz", "10 Hz", "5 Hz", "3 Hz",
             "30 Hz, 300 kN/s", "30 Hz, 150 kN/s", "30 Hz, 75 kN/s"]
    vals = [0.475, 0.487, 0.644, 0.763, 0.873, 0.850, 0.610, 0.778, 0.803]
    passive = 0.698
    ch = bar_chart(s, 0.85, 3.9, 14.6, 9.9, cases[::-1], [("Pa peak LTR", vals[::-1], C_PA)],
                   title="Pa peak LTR by actuator (passive: 0.698)", horizontal=True, fmt="0.000",
                   legend=False, gap=45, overlap=0, vmin=0, vmax=1.0, major=0.2, val_fmt="0.0", axis_size=22)
    for k, v in enumerate(vals[::-1]):
        if v >= passive:
            pt = ch.plots[0].series[0].points[k]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(C_P)
    stat(s, 15.9, 3.9, 9.9, 3.1, "15–30 Hz",
         "bandwidth: the gain survives to ≈ 15 Hz and is complete by 30 Hz", big_size=56, label_size=28)
    stat(s, 15.9, 7.35, 9.9, 3.1, "≈ 300 kN/s",
         "slew rate keeps most of the gain; 150 kN/s removes it", big_size=56, label_size=28)
    bullets(s, 15.9, 10.8, 9.9, 3.0, [
        "≤ 10 Hz: force arrives after the 11.2 Hz wheel-hop event; worse than passive (gray bars)",
        "Lean removal (roll ≈ 0.6–0.7°) survives at every bandwidth",
    ], size=27, space_after=8)
    footnote(s, "First-order lag and slew limit on the nominal controller, not redesigned. Baseline case, mean of five roads.")


def slide_discussion(d):
    s = d.content("6. DISCUSSION",
                  "Why feedback degrades the margin: it leans further than passive, a lower effective "
                  "roll stiffness. Its model has no lateral-acceleration term, so it cannot tell the "
                  "steady corner load from a road disturbance. This is a hypothesis; closed-loop modal "
                  "analysis is future work. The two horizon criteria differ by four times, so pick the "
                  "horizon against the reported metric. The vertical model cannot represent tripped "
                  "rollover, such as an outer tyre striking a pothole edge.")
    cards = [
        ("Why feedback degrades the margin",
         "B leans 5.4° against passive 3.0°: lower effective roll stiffness. With no a_y term in its model it "
         "treats the steady corner load as a disturbance and softens roll against it. A higher LTR weight makes "
         "it worse. Hypothesis; closed-loop modal analysis pending."),
        ("Two criteria for the horizon",
         "LQ cost saturates at 0.28 m; peak LTR needs 1.04 m, about 4× longer. The cost is dominated by the "
         "near present; the peak is one instant that needs action before impact. Choose the horizon against the "
         "reported metric. 1 m is about one wheel-hop wavelength, within sensor range."),
        ("What the vertical model cannot say",
         "It resolves how tyre loads redistribute and whether the inner wheels unload. It cannot show tripped "
         "rollover: an outer-wheel pothole looks mild here, but the tyre can strike the far edge and trip the car."),
    ]
    for k, (head, body) in enumerate(cards):
        x = 0.85 + k * 8.4
        card(s, x, 3.9, 8.0, 9.8, PINK if k != 1 else GRAY_CARD)
        number_badge(s, x + 0.45, 4.3, 0.9, f"{k + 1}", size=28)
        text(s, x + 0.45, 5.5, 7.1, 1.4, head, size=31, bold=True, color=MAROON, font=HEAD_FONT)
        text(s, x + 0.45, 7.1, 7.1, 6.4, body, size=30)


def slide_limitations(d):
    s = d.content("6. LIMITATIONS AND NEXT STEPS",
                  "What the results rest on, and the work that follows from each limitation. All next "
                  "steps will be run in CarSim for higher-fidelity, visual simulation.")
    card(s, 0.85, 3.95, 12.2, 9.7, PINK)
    text(s, 1.35, 4.3, 11.2, 0.9, "Limitations", size=36, bold=True, color=MAROON, font=HEAD_FONT)
    bullets(s, 1.35, 5.45, 11.2, 8.0, [
        "Prescribed a_y: no lateral or yaw dynamics, no tyre force saturation",
        "Unsprung load transfer not modelled: LTR ≈ 7 % low, all controllers equally",
        "Exact preview: no sensor noise, position error or delay",
        "Weights and horizon chosen on the baseline scenario",
        "Measured-vehicle suspension and tyre data assumed; limits not in the design",
    ], size=34, space_after=24)
    card(s, 13.6, 3.95, 12.2, 9.7, GRAY_CARD)
    text(s, 14.1, 4.3, 11.2, 0.9, "Next steps", size=36, bold=True, color=MAROON, font=HEAD_FONT)
    bullets(s, 14.1, 5.45, 11.2, 8.0, [
        "Model-error study: nominal design on similar-but-different vehicles",
        "Sensor noise and defect-position error in the preview",
        "Actuator dynamics inside the controller design",
        "Constrained formulation (MPC) where travel or force limits act",
        "CarSim validation with lateral tyre forces, incl. the pothole edge strike",
    ], size=34, space_after=24)


def slide_conclusion(d):
    s = d.content("6. CONCLUSION",
                  "An active suspension cannot change the load transfer of a steady corner; it can change "
                  "the transient at a defect met in that corner. Feedback LQR does not: it leans further "
                  "and at high lateral acceleration reaches the trip condition where passive does not. "
                  "Previewing the road and the planned lateral acceleration does, on every road and in "
                  "every cell tested, with an actuator of about 15 to 30 Hz and 300 kN/s.")
    tiles = [("≈ 6 %", "of steady-corner load transfer any suspension can change"),
             ("11 / 25", "roads where feedback reached LTR = 1 at 0.8 g (passive: 1)"),
             ("51 / 55", "cells passed by road + a_y preview, no re-tuning"),
             ("15–30 Hz", "actuator bandwidth needed, with ≈ 300 kN/s slew")]
    for k, (big, label) in enumerate(tiles):
        stat(s, 0.85 + k * 6.34, 3.9, 5.94, 4.4, big, label, big_size=56, label_size=28)
    bullets(s, 0.85, 8.8, 24.9, 3.2, [
        ("Road preview handles the defect; a_y preview handles the corner.",
         "Together they lower peak LTR on every road, remove the lean and prevent the trip condition across the grid."),
        ("Feedback LQR degrades the margin:", "lower effective roll stiffness, and no way to tell the corner from the road."),
    ], size=31, space_after=12)
    text(s, 0.85, 12.2, 24.9, 1.4,
         "Within a vertical model, information an automated vehicle already has protects a significant part of the rollover margin.",
         size=33, bold=True, color=MAROON, font=HEAD_FONT)


def slide_end(d):
    s = d._new(LAYOUT_END, None)
    text(s, 0.0, 11.2, 26.67, 1.0, "THANK YOU  ·  QUESTIONS", size=40, bold=True, color=WHITE,
         font=HEAD_FONT, align=PP_ALIGN.CENTER)


def build(template, figures, out):
    d = Deck(template, figures)
    slide_title(d)
    slide_outline(d)
    slide_setup_model(d)
    slide_setup_scenario(d)
    d.divider("STUDY 1", "Feedback control in straight-line driving",
              "What output-weighted LQR achieves on a rough road and at a discrete defect: "
              "the baseline Study 2 departs from.",
              "Study 1 fixes the baseline: what a feedback active suspension can and cannot do before the corner is added.")
    slide_study1_passive(d)
    slide_study1_selection(d)
    slide_study1_feedback(d)
    slide_study1_carry(d)
    d.divider("STUDY 2", "Preview control for rollover margin in cornering",
              "Can the road ahead and the planned lateral acceleration protect the margin "
              "when a defect is met mid-corner?",
              "Study 2 is the core of the project: the cornering scenario, preview control, and the transient rollover metric.")
    slide_corner_limits(d)
    slide_controllers(d)
    slide_horizon(d)
    slide_baseline(d)
    slide_statistics(d)
    slide_robustness(d)
    slide_generalisation(d)
    slide_severity(d)
    slide_actuator(d)
    slide_discussion(d)
    slide_limitations(d)
    slide_conclusion(d)
    slide_end(d)
    d.prs.save(out)
    print(f"saved {out} ({len(d.prs.slides)} slides)")


if __name__ == "__main__":
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", default=str(here / "VIPP_Progress_Presentation.pptx"))
    ap.add_argument("--figures", default=str(here / "figures"))
    ap.add_argument("--out", default=str(here / "VIPP_Final_Presentation.pptx"))
    args = ap.parse_args()
    if not Path(args.template).exists():
        sys.exit(f"Template not found: {args.template}\nPass it with --template \"path\\to\\VIPP_Progress_Presentation.pptx\"")
    build(args.template, args.figures, args.out)
