from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

INK = RGBColor(0x1a, 0x1a, 0x1a)
MUTED = RGBColor(0x5a, 0x5a, 0x5a)
ACCENT = RGBColor(0x3b, 0x6e, 0x8f)
PAPER = RGBColor(0xfa, 0xf9, 0xf6)
RULE = RGBColor(0xcf, 0xc9, 0xbd)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

def add_bg(slide, color=PAPER):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def add_textbox(slide, left, top, width, height, text, size, color=INK, bold=False,
                 align=PP_ALIGN.LEFT, font="Georgia", italic=False, line_spacing=1.0):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing != 1.0:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    return box

def add_rule(slide, left, top, width, color=RULE, weight=Pt(1)):
    height = Emu(int(weight))
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = color
    box.line.fill.background()
    box.shadow.inherit = False

def add_kicker(slide, text, left=Inches(0.7), top=Inches(0.45)):
    add_textbox(slide, left, top, Inches(6), Inches(0.4), text.upper(), 13,
                color=ACCENT, bold=True, font="Verdana")

def slide_number(slide, n):
    add_textbox(slide, Inches(12.5), Inches(7.0), Inches(0.6), Inches(0.4), str(n), 11, color=MUTED, font="Verdana")

# ---------------------------------------------------------------------------
# Slide 1: Title
# ---------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
add_bg(s)
add_textbox(s, Inches(0.7), Inches(2.6), Inches(10), Inches(0.4),
            "IMDB CATALOG, 2010–2019", 14, color=ACCENT, bold=True, font="Verdana")
add_textbox(s, Inches(0.7), Inches(3.0), Inches(11.5), Inches(1.6),
            "Ratings tell a different story than audiences do", 36, color=INK, bold=True)
add_textbox(s, Inches(0.7), Inches(4.1), Inches(9.5), Inches(0.6),
            "A first look at what the IMDB movie tables reveal, where the gaps are, and what's worth chasing next.",
            16, color=MUTED, italic=True)
add_rule(s, Inches(0.7), Inches(4.85), Inches(3))
add_textbox(s, Inches(0.7), Inches(6.6), Inches(6), Inches(0.4),
            "Danielle Shokrian  —  Part 2 Exploratory Analysis", 12, color=MUTED, font="Verdana")
slide_number(s, 1)

# ---------------------------------------------------------------------------
# Slide 2: Exploration findings
# ---------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
add_bg(s)
add_kicker(s, "What the data shows")
add_textbox(s, Inches(0.7), Inches(0.85), Inches(10), Inches(0.8),
            "Genre tags are crowded, and engagement isn't evenly spread", 26, bold=True)
add_rule(s, Inches(0.7), Inches(1.65), Inches(11.9))

rows = [
    ("30,788 / 17,753 / 17,290",
     "Drama, Documentary, and Comedy dominate the catalog by title count — nearly half of every rated movie carries one of these three tags."),
    ("Most titles carry 2–3 genre tags at once",
     "\"Comedy,Drama,Romance\" and similar combinations are common, so a movie tagged this way needs to be split and counted under each genre, not treated as its own category."),
    ("7,620 missing runtimes · 804 missing genre tags",
     "Both fields have real gaps in the movie_basics table that had to be excluded outright — no filling in, per the assignment's instructions."),
    ("Reach and acclaim move independently",
     "Documentary titles average a 7.3 rating but draw only ~270 votes per title; Adventure titles average a lower 6.2 rating but draw over 22,000 votes."),
]

top = Inches(2.05)
row_h = Inches(1.15)
for i, (lead, body) in enumerate(rows):
    y = top + row_h * i
    add_textbox(s, Inches(0.7), y, Inches(3.6), row_h, lead, 16, color=ACCENT, bold=True, line_spacing=1.05)
    add_textbox(s, Inches(4.5), y, Inches(8.0), row_h, body, 14, color=INK, line_spacing=1.15)
    if i < len(rows) - 1:
        add_rule(s, Inches(0.7), y + Inches(1.0), Inches(11.9), color=RGBColor(0xe5,0xe1,0xd6), weight=Pt(0.75))
slide_number(s, 2)

# ---------------------------------------------------------------------------
# Slide 3: Visualization
# ---------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
add_bg(s)
add_kicker(s, "Engagement vs. acclaim")
add_textbox(s, Inches(0.7), Inches(0.85), Inches(11), Inches(0.8),
            "Action and Adventure pull crowds; Documentary pulls praise", 26, bold=True)
s.shapes.add_picture("images/genre_votes_chart.png", Inches(1.1), Inches(1.75), height=Inches(5.0))
add_textbox(s, Inches(9.6), Inches(2.1), Inches(3.2), Inches(4.2),
            "Average vote count per title, by genre (2010–2019). "
            "Labels show each genre's average rating for comparison.\n\n"
            "Adventure titles draw roughly 80x more votes on average than Documentary titles, "
            "despite rating over a full point lower.",
            13, color=MUTED, line_spacing=1.2, font="Verdana")
slide_number(s, 3)

# ---------------------------------------------------------------------------
# Slide 4: Business question
# ---------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
add_bg(s, color=ACCENT)
add_kicker(s, "Where this leads", left=Inches(0.7), top=Inches(0.5))
add_textbox(s, Inches(0.7), Inches(0.5), Inches(10), Inches(0.4), "", 1)
title_box = s.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(11.9), Inches(2.6))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
run = p.add_run()
run.text = "If a studio could only chase one of these, which genre bet pays off — reach or reputation?"
run.font.size = Pt(30)
run.font.bold = True
run.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
run.font.name = "Georgia"
p.line_spacing = 1.15

add_textbox(s, Inches(0.7), Inches(4.2), Inches(11.5), Inches(1.6),
            "Specifically: do high-vote, average-rated genres like Action and Adventure convert into stronger "
            "long-run financial outcomes than high-rated, low-vote genres like Documentary and Biography — "
            "and does that tradeoff hold once runtime and release year are controlled for?",
            16, color=RGBColor(0xf0, 0xef, 0xe9), line_spacing=1.25)
slide_number(s, 4)

# ---------------------------------------------------------------------------
# Slide 5: Data cleaning tasks
# ---------------------------------------------------------------------------
s = prs.slides.add_slide(blank)
add_bg(s)
add_kicker(s, "Before any of that analysis happens")
add_textbox(s, Inches(0.7), Inches(0.85), Inches(11), Inches(0.8),
            "Four cleanup tasks come first", 26, bold=True)
add_rule(s, Inches(0.7), Inches(1.65), Inches(11.9))

tasks = [
    ("01", "Split combined genre tags",
     "Explode the comma-separated genres column into one row per genre before grouping, so multi-genre titles aren't miscounted."),
    ("02", "Drop, don't fill, missing values",
     "runtime_minutes and genres both have real nulls; exclude those rows from the relevant analysis rather than imputing a value."),
    ("03", "Enforce the 2019 cutoff",
     "start_year contains values past 2019; every query needs an explicit filter since the dataset is only validated through that year."),
    ("04", "Watch for join fan-out",
     "principals, directors, and writers are many-to-many tables — joining them onto movie_basics can multiply rows per movie if not deduplicated first."),
]

col_w = Inches(5.75)
row_h = Inches(2.35)
positions = [(Inches(0.7), Inches(2.05)), (Inches(6.85), Inches(2.05)),
             (Inches(0.7), Inches(4.55)), (Inches(6.85), Inches(4.55))]
for (num, lead, body), (x, y) in zip(tasks, positions):
    add_textbox(s, x, y, Inches(1.0), Inches(0.6), num, 22, color=RULE, bold=True, font="Verdana")
    add_textbox(s, x, y + Inches(0.55), col_w, Inches(0.5), lead, 17, color=INK, bold=True)
    add_textbox(s, x, y + Inches(1.05), col_w, Inches(1.2), body, 13, color=MUTED, line_spacing=1.2)
slide_number(s, 5)

prs.save("Movie_Data_Exploration_Slides.pptx")
print("saved pptx")
