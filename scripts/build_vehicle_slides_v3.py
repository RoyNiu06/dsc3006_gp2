"""Build a cleaner, higher-contrast DSC3006 Vehicle Sales deck (v3)."""
from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

ROOT = Path(r"F:\Study\2627semA\3006-机器学习\Project")
FIG = Path(r"F:\Study\ALL_PROJECT\repos\DSC3006\vehicle_sales\results\figures")
OUT = ROOT / "DSC3006_VehicleSales_Slides.pptx"

NAVY = RGBColor(14, 28, 45)
TEAL = RGBColor(12, 106, 120)
TEAL_LIGHT = RGBColor(226, 241, 243)
ORANGE = RGBColor(200, 92, 60)
INK = RGBColor(28, 37, 48)
MUTED = RGBColor(91, 105, 120)
LINE = RGBColor(211, 220, 230)
BG = RGBColor(250, 251, 253)
WHITE = RGBColor(255, 255, 255)
PALE = RGBColor(242, 246, 249)
FONT = "Segoe UI"
TITLE_FONT = "Segoe UI Semibold"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def set_bg(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def text(slide, x, y, w, h, s, size=16, color=INK, bold=False,
         font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    r = p.add_run()
    r.text = s
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def bullets(slide, x, y, w, h, items, size=16, color=INK, spacing=1.10, gap=8):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = spacing
        p.space_after = Pt(gap)
        r1 = p.add_run()
        r1.text = "•  "
        r1.font.name = FONT
        r1.font.size = Pt(size)
        r1.font.color.rgb = TEAL
        r1.font.bold = True
        r2 = p.add_run()
        r2.text = item
        r2.font.name = FONT
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
    return box


def rect(slide, x, y, w, h, fill=WHITE, line=None, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is not None:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    if radius:
        shp.adjustments[0] = 0.05
    return shp


def line(slide, x1, y1, x2, y2, color=LINE, width=1):
    c = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


def picture_fit(slide, path, x, y, w, h):
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    ww, hh = iw * scale, ih * scale
    xx = x + (w - ww) / 2
    yy = y + (h - hh) / 2
    return slide.shapes.add_picture(str(path), Inches(xx), Inches(yy), Inches(ww), Inches(hh))


def header(slide, kicker, title, num=None):
    text(slide, 0.7, 0.34, 10.0, 0.3, kicker.upper(), 11, TEAL, True)
    text(slide, 0.7, 0.68, 11.8, 0.8, title, 30, NAVY, True, TITLE_FONT)
    line(slide, 0.7, 1.50, 12.65, 1.50, LINE, 1)
    if num:
        text(slide, 11.85, 0.34, 0.8, 0.3, f"{num:02d}", 11, MUTED, True, FONT, PP_ALIGN.RIGHT)


def footer(slide):
    line(slide, 0.7, 7.12, 12.65, 7.12, LINE, 0.8)
    text(slide, 0.7, 7.20, 8.0, 0.22, "DSC3006 · Vehicle Sales · 29 November 2026", 9.5, MUTED)
    text(slide, 11.7, 7.20, 0.95, 0.22, "DSC3006", 9.5, MUTED, True, FONT, PP_ALIGN.RIGHT)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def fmt_cell(cell, s, size=13, color=INK, bold=False, fill=None, align=PP_ALIGN.LEFT):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    cell.margin_left = Inches(0.10)
    cell.margin_right = Inches(0.10)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.04)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def table(slide, x, y, w, h, data, widths=None, size=13):
    rows, cols = len(data), len(data[0])
    shp = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    t = shp.table
    if widths:
        total = sum(widths)
        for i, cw in enumerate(widths):
            t.columns[i].width = Inches(w * cw / total)
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            if r == 0:
                fmt_cell(cell, str(val), size, WHITE, True, NAVY,
                         PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
            else:
                fmt_cell(cell, str(val), size, INK, c == 0,
                         WHITE if r % 2 else PALE,
                         PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    return t


# =====================================================================
# 1. TITLE
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s, NAVY)
text(s, 0.85, 0.62, 6.0, 0.3, "DSC3006 · FINAL PROJECT", 11, RGBColor(143, 211, 220), True)
text(s, 0.85, 1.00, 6.0, 0.3, "Statistical Machine Learning", 15, RGBColor(203, 213, 225))
text(s, 0.85, 1.75, 11.6, 1.65,
     "Used-Car Price Prediction:\nTemporal Validation and Error Analysis",
     41, WHITE, True, TITLE_FONT, spacing=0.95)
text(s, 0.87, 3.65, 11.0, 0.5,
     "Vehicle Sales auction data  ·  Group XX  ·  Member A, Member B, Member C",
     16, RGBColor(203, 213, 225))
text(s, 0.87, 4.15, 5.0, 0.35, "29 November 2026", 13, RGBColor(148, 163, 184))
text(s, 0.87, 5.15, 11.6, 0.35,
     "558,837 auction records   |   403,815 final-fit rows   |   153,684 held-out test rows",
     14, RGBColor(203, 213, 225))
text(s, 0.87, 5.62, 11.6, 0.35,
     "Target: sellingprice   |   Features available at auction", 14, RGBColor(148, 163, 184))
line(s, 0.87, 6.55, 12.45, 6.55, RGBColor(54, 80, 106), 1)
text(s, 0.87, 6.72, 11.2, 0.3,
     "A regression study of what car attributes explain, and where the model fails.", 11, RGBColor(203, 213, 225))
notes(s, "This is a regression project on used-car auction prices. The focus is a clean temporal validation design and a careful analysis of errors, not a commercial valuation product.")

# =====================================================================
# 2. THREE QUESTIONS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "01 · Questions", "Three questions", 2)
qs = [
    ("01", "How much of the price can auction-time attributes explain?"),
    ("02", "Does a more flexible model beat a linear baseline?"),
    ("03", "Where should we not trust the prediction?"),
]
for i, (num, q) in enumerate(qs):
    y = 2.05 + i * 1.45
    text(s, 0.75, y, 0.8, 0.35, num, 14, TEAL, True)
    text(s, 1.55, y - 0.06, 10.6, 0.9, q, 21, NAVY, True, TITLE_FONT, spacing=1.05)
    if i < 2:
        line(s, 1.55, y + 1.05, 12.3, y + 1.05, LINE, 0.8)
text(s, 0.75, 6.35, 11.8, 0.35,
     "Data: 558,837 auctions  ·  Test: 153,684 rows from May–July 2015  ·  Models: Ridge, Spline, HGB",
     11, MUTED)
footer(s)
notes(s, "Open with three questions. We answer them at the end: how predictable price is, whether flexibility helps, and where the prediction is weak.")

# =====================================================================
# 3. DATA
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "02 · Data", "Data and sample", 3)
rows = [
    ["Stage", "Rows", "Period"],
    ["Raw auction records", "558,837", "Dec 2014 – Jul 2015"],
    ["Clean records", "558,592", "after exclusions"],
    ["Training", "354,936", "Dec 2014 – Feb 2015"],
    ["Validation", "47,526", "Mar – Apr 2015"],
    ["Final fit", "403,815", "Dec 2014 – Apr 2015"],
    ["Test", "153,684", "May – Jul 2015"],
]
table(s, 0.7, 1.75, 6.65, 4.25, rows, [2.2, 1.2, 2.0], 13)
text(s, 0.7, 6.18, 6.4, 0.3, "Source: Kaggle Vehicle Sales Data (public).", 10.5, MUTED)

text(s, 7.7, 1.78, 5.0, 0.3, "INPUTS", 11, TEAL, True)
text(s, 7.7, 2.12, 4.9, 1.0,
     "model_year, log(odometer), make, model, trim, body, transmission, state, color, interior, condition",
     15, INK, spacing=1.08)
text(s, 7.7, 3.28, 5.0, 0.3, "EXCLUDED", 11, ORANGE, True)
bullets(s, 7.7, 3.62, 4.85, 1.55, [
    "MMR: an existing price estimate.",
    "VIN: grouping only.",
    "Seller and sale date: identity / split variables.",
], 13.5, spacing=1.06, gap=5)
text(s, 7.7, 5.35, 4.9, 0.8,
     "245 raw rows removed for invalid price, date, state, year or odometer.",
     11, MUTED, spacing=1.05)
footer(s)
notes(s, "Vehicle Sales is the only dataset. The table shows the temporal stages; the right side lists model inputs and excluded fields. MMR is excluded because it is an existing price estimate; VIN is only an identity; sale date is only used for splitting.")

# =====================================================================
# 4. VALIDATION
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "03 · Validation", "Temporal validation", 4)
blocks = [
    ("TRAIN", "Dec 2014 – Feb 2015", "354,936 rows", NAVY),
    ("VALIDATION", "Mar – Apr 2015", "47,526 rows", TEAL),
    ("TEST", "May – Jul 2015", "153,684 rows", ORANGE),
]
for i, (label, dates, n, color) in enumerate(blocks):
    x = 0.75 + i * 4.07
    rect(s, x, 1.85, 3.55, 1.35, color)
    text(s, x + 0.22, 2.02, 3.0, 0.25, label, 10.5, WHITE, True)
    text(s, x + 0.22, 2.34, 3.1, 0.35, dates, 15, WHITE, True)
    text(s, x + 0.22, 2.78, 3.1, 0.25, n, 11, RGBColor(226, 232, 240))

bullets(s, 0.75, 3.55, 11.9, 1.45, [
    "The test period is later than training and validation.",
    "A vehicle appearing later is removed from earlier splits; VIN overlap is zero.",
    "HGB31 is selected on validation log RMSE, refit on train + validation, and tested once.",
], 15.5, spacing=1.08, gap=7)
text(s, 0.75, 5.45, 11.8, 0.35,
     "2,245 train rows and 201 validation rows were purged to keep VINs separate.",
     11, MUTED)
text(s, 0.75, 6.0, 11.8, 0.35,
     "Random K-fold is not used: it would put future prices, and the same vehicle, into training.",
     11, MUTED)
footer(s)
notes(s, "The split is temporal, not random. Test is later than train and validation. Any vehicle seen later is removed from earlier data. HGB31 is selected on validation, refit on train plus validation, and evaluated once on test.")

# =====================================================================
# 5. FORMULAS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "04 · Models", "Model formulation", 5)
text(s, 0.7, 1.62, 11.7, 0.3,
     "All models predict log price; predictions are exponentiated back to price.", 12, MUTED)

cards = [
    ("MEAN BASELINE", "y_hat = exp( mean(log y_train) )",
     "Fitted on 403,815 rows: exp(9.164) = 9,549."),
    ("RIDGE", "y_hat = exp( β0 + Σ βj xj )",
     "Minimises squared error + α Σ βj²; α = 10."),
    ("SPLINE RIDGE", "y_hat = exp( β0 + Σ βk Bk(year) + Σ γl Bl(log odo) + Σ δc 1[cat] )",
     "Cubic B-splines with 6 knots; ridge coefficients."),
    ("HGB31", "y_hat = exp( Σm=1..220 ν fm(x) )",
     "fm is a regression tree (≤31 leaves); ν = 0.07; no closed form."),
]
pos = [(0.7, 1.98, 5.95, 1.55), (6.92, 1.98, 5.95, 1.55),
       (0.7, 3.75, 5.95, 1.55), (6.92, 3.75, 5.95, 1.55)]
for (label, formula, note), (x, y, w, h) in zip(cards, pos):
    rect(s, x, y, w, h, WHITE, LINE)
    text(s, x + 0.22, y + 0.14, w - 0.4, 0.25, label, 10, TEAL, True)
    text(s, x + 0.22, y + 0.46, w - 0.4, 0.45, formula, 13, NAVY, True, FONT, spacing=1.0)
    text(s, x + 0.22, y + 1.08, w - 0.4, 0.35, note, 10.5, MUTED, spacing=1.03)

rect(s, 0.7, 5.55, 12.15, 0.72, TEAL_LIGHT)
text(s, 0.92, 5.68, 11.7, 0.22, "VALIDATION SELECTION", 9.5, TEAL, True)
text(s, 0.92, 5.94, 11.7, 0.25,
     "Ridge 0.3214   ·   Spline 0.2718   ·   HGB15 0.2570   ·   HGB31 0.2503   →   HGB31",
     12, NAVY, True)
text(s, 0.7, 6.48, 11.8, 0.3,
     "HGB15 was a validation candidate only; it was not carried to test after HGB31 was selected.",
     10.5, MUTED)
footer(s)
notes(s, "Mean baseline uses the geometric mean of training prices. Ridge is a linear model on the log scale with an L2 penalty. The spline model expands year and mileage into cubic B-spline bases and applies ridge. HGB is an additive sum of 220 small trees with shrinkage 0.07; it has no closed-form equation.")

# =====================================================================
# 6. RESULTS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "05 · Results", "Test results", 6)
text(s, 0.7, 1.62, 8.0, 0.3, "Held-out test: 153,684 rows, May–July 2015", 12, MUTED)

cd = CategoryChartData()
cd.categories = ["Ridge", "Spline", "HGB31"]
cd.add_series("MAE", (2388.2, 1929.3, 1626.1))
chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(2.0),
                           Inches(7.0), Inches(4.0), cd).chart
chart.has_legend = False
chart.has_title = True
chart.chart_title.text_frame.text = "Test MAE (price units, lower is better)"
chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
chart.chart_title.text_frame.paragraphs[0].font.color.rgb = MUTED
plot = chart.plots[0]
plot.gap_width = 65
plot.has_data_labels = True
dl = plot.data_labels
dl.number_format = '#,##0'
dl.number_format_is_linked = False
dl.font.size = Pt(11)
dl.font.name = FONT
dl.font.color.rgb = NAVY
for i, p in enumerate(plot.series[0].points):
    p.format.fill.solid()
    p.format.fill.fore_color.rgb = [RGBColor(158, 171, 186), RGBColor(76, 150, 166), TEAL][i]
chart.value_axis.maximum_scale = 2800
chart.value_axis.minimum_scale = 0
chart.value_axis.has_major_gridlines = True
chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(230, 236, 242)
chart.value_axis.tick_labels.font.size = Pt(9.5)
chart.value_axis.tick_labels.font.color.rgb = MUTED
chart.category_axis.tick_labels.font.size = Pt(11)
chart.category_axis.tick_labels.font.color.rgb = INK

rows = [
    ["Model", "MAE", "R²", "Log RMSE"],
    ["Mean", "7,272", "−0.025", "0.892"],
    ["Ridge", "2,388", "0.841", "0.302"],
    ["Spline", "1,929", "0.877", "0.256"],
    ["HGB31", "1,626", "0.900", "0.232"],
]
table(s, 7.95, 2.0, 4.65, 2.55, rows, [1.65, 1.0, 0.8, 1.1], 12)
text(s, 7.95, 4.78, 4.65, 0.3, "HGB31 vs Ridge: −31.9% MAE", 13, TEAL, True)
text(s, 7.95, 5.18, 4.65, 0.3, "Log RMSE: 0.302 → 0.232", 11.5, MUTED)
text(s, 0.7, 6.28, 11.9, 0.3,
     "MAE is the primary metric; R² is descriptive. Percentage errors are shown by price band later.",
     10.5, MUTED)
footer(s)
notes(s, "On the test period, HGB31 reduces MAE from 2,388 for Ridge to 1,626. Log RMSE falls from 0.302 to 0.232. Raw R-squared is 0.90, but it is descriptive; the price range matters.")

# =====================================================================
# 7. INTERPRETATION
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "06 · Interpretation", "Feature importance and condition", 7)
text(s, 0.7, 1.62, 7.4, 0.3, "Validation permutation importance", 11, MUTED)
picture_fit(s, FIG / "vehicle_sales_importance.png", 0.55, 1.95, 7.6, 3.65)

rows = [
    ["Model", "MAE", "Log RMSE"],
    ["HGB31", "1,626", "0.232"],
    ["without condition", "1,763", "0.256"],
]
table(s, 8.55, 1.98, 3.95, 1.3, rows, [1.9, 0.9, 1.1], 11.5)
text(s, 8.55, 3.5, 3.95, 0.3, "Condition: −7.8% MAE", 14, TEAL, True)
text(s, 8.55, 3.95, 3.95, 1.25,
     "model_year, model and mileage dominate. Condition adds a moderate increment.",
     13, INK, spacing=1.08)
text(s, 8.55, 5.55, 3.95, 0.3, "Condition codes are categorical, not ordinal.", 10.5, MUTED)
footer(s)
notes(s, "Age, model and mileage are the main predictors. Removing condition raises MAE from 1,626 to 1,763, so condition adds about 7.8 percent. The condition codes are treated as categories, not as an ordinal scale.")

# =====================================================================
# 8. ERRORS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "07 · Error analysis", "Absolute and percentage error", 8)
text(s, 0.7, 1.62, 9.0, 0.3,
     "Left: absolute error (price units). Right: median absolute percentage error.", 11, MUTED)
picture_fit(s, FIG / "vehicle_sales_error_bands.png", 0.55, 1.98, 8.5, 3.35)

rect(s, 9.35, 2.0, 3.2, 1.22, PALE, LINE)
text(s, 9.58, 2.16, 2.8, 0.24, "LOW PRICE", 9.5, ORANGE, True)
text(s, 9.58, 2.44, 2.9, 0.4, "Median APE 24.9%", 16, NAVY, True, TITLE_FONT)
text(s, 9.58, 2.88, 2.8, 0.22, "mean APE 45.6%", 9.5, MUTED)
rect(s, 9.35, 3.42, 3.2, 1.22, PALE, LINE)
text(s, 9.58, 3.58, 2.8, 0.24, "HIGH PRICE", 9.5, TEAL, True)
text(s, 9.58, 3.86, 2.9, 0.4, "MAE 5,890", 16, NAVY, True, TITLE_FONT)
text(s, 9.58, 4.30, 2.8, 0.22, "bias −3,676", 9.5, MUTED)

text(s, 0.7, 5.55, 11.9, 0.75,
     "Absolute error rises with price; relative error is worst at the bottom. High-price cars are underpriced.",
     16, NAVY, True, TITLE_FONT, spacing=1.05)
text(s, 0.7, 6.40, 11.9, 0.3,
     "High-price median APE 8.9%, close to the 6.3% of the 15–30k band. July has only 1,300 rows.",
     10.5, MUTED)
footer(s)
notes(s, "The left panel is absolute error; the right panel is median percentage error. Expensive cars have larger absolute errors and are underpriced, but their median percentage error is close to the middle band. The cheapest cars have the largest relative error.")

# =====================================================================
# 9. LIMITATIONS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s)
header(s, "08 · Limitations", "Limitations", 9)
text(s, 0.75, 1.75, 3.7, 0.3, "DATA", 11, ORANGE, True)
bullets(s, 0.75, 2.10, 3.65, 3.2, [
    "Price definition and currency are not independently verified.",
    "Condition codes are categorical, not an ordinal grade.",
], 14, spacing=1.10, gap=10)

text(s, 4.65, 1.75, 3.7, 0.3, "TIME", 11, ORANGE, True)
bullets(s, 4.65, 2.10, 3.65, 3.2, [
    "One annual cycle: month and time are confounded.",
    "May/June MAE 1,609/1,627; July 2,267 with 1,300 rows.",
], 14, spacing=1.10, gap=10)

text(s, 8.55, 1.75, 3.7, 0.3, "SCOPE", 11, ORANGE, True)
bullets(s, 8.55, 2.10, 3.65, 3.2, [
    "Single auction market; UK data was not merged.",
    "Results are predictive, not causal.",
], 14, spacing=1.10, gap=10)

text(s, 0.75, 5.55, 11.9, 0.7,
     "Within this market and window, the model is most useful for relative comparisons; high-price point predictions need the most caution.",
     14, NAVY, True, TITLE_FONT, spacing=1.05)
footer(s)
notes(s, "Keep the limitations short. Price definition and currency are not independently verified. The condition codes are categorical. The data cover one annual cycle, so month and time are confounded; May and June are stable, July is partial. The UK data was not merged because its price definition and timestamps differ. These are predictive associations, not causal effects.")

# =====================================================================
# 10. ANSWERS
# =====================================================================
s = prs.slides.add_slide(BLANK)
set_bg(s, NAVY)
text(s, 0.85, 0.62, 5.0, 0.3, "09 · ANSWERS", 10.5, RGBColor(143, 211, 220), True)
text(s, 0.85, 0.98, 11.6, 0.75, "Answers", 32, WHITE, True, TITLE_FONT)
line(s, 0.85, 1.85, 12.45, 1.85, RGBColor(54, 80, 106), 1)

answers = [
    ("Q1 · HOW MUCH?", "R² 0.900; MAE 1,626; median APE 8.25%. Attributes explain a large share, but the price range matters."),
    ("Q2 · DOES FLEXIBILITY HELP?", "Yes. MAE falls from 2,388 (Ridge) to 1,929 (Spline) to 1,626 (HGB31)."),
    ("Q3 · WHERE IS IT WEAK?", "High-price cars: MAE 5,890, bias −3,676, median APE 8.9%. Low-price cars: MAE 808, median APE 24.9%."),
]
for i, (label, body) in enumerate(answers):
    y = 2.10 + i * 1.32
    text(s, 0.85, y, 11.2, 0.25, label, 9.8, RGBColor(143, 211, 220), True)
    text(s, 0.85, y + 0.30, 11.5, 0.75, body, 14.5, WHITE, spacing=1.06)

text(s, 0.85, 6.05, 11.6, 0.4,
     "Use: relative pricing and screening.  |  Reproducible: frozen split, 232/232 audit checks.",
     13, RGBColor(203, 213, 225))
line(s, 0.85, 6.85, 12.45, 6.85, RGBColor(54, 80, 106), 1)
text(s, 0.85, 6.98, 7.0, 0.25, "Thank you  ·  Questions", 10.5, RGBColor(203, 213, 225))
text(s, 9.9, 6.98, 2.55, 0.25, "DSC3006 · Group XX", 9.5, RGBColor(159, 179, 200), False, FONT, PP_ALIGN.RIGHT)
notes(s, "Answer the opening questions. Q1: price is predictable on average, but R-squared overstates percentage accuracy. Q2: flexibility helps, with HGB best. Q3: high-price cars are underpriced in absolute terms; low-price cars have the largest percentage error. Use the model for relative pricing and screening.")

prs.save(OUT)
print("saved", OUT)
