"""
CPPL Prototype Planning — Minimalist presentation with Webteam branding.
Brand colors extracted from logo: Red #E8000D  |  Navy #1A0080
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── brand palette ─────────────────────────────────────────────────────────────
RED        = RGBColor(0xE8, 0x00, 0x0D)   # webteam red
NAVY       = RGBColor(0x1A, 0x00, 0x80)   # webteam navy
INK        = RGBColor(0x1A, 0x1A, 0x2E)   # near-black body text
GRAY       = RGBColor(0x55, 0x55, 0x66)   # secondary text
LIGHT      = RGBColor(0xF5, 0xF6, 0xFA)   # card background
BORDER     = RGBColor(0xE2, 0xE4, 0xED)   # card border
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GREEN      = RGBColor(0x0A, 0x7A, 0x4B)
ORANGE     = RGBColor(0xD4, 0x5F, 0x00)

W = Inches(13.33)
H = Inches(7.5)

LOGO = r"C:\Users\Dell\CPPL\webteam logo.png"

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


# ── primitives ────────────────────────────────────────────────────────────────

def rect(slide, x, y, w, h, fill=None, border=None, bw=Pt(0.75)):
    s = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if border:
        s.line.color.rgb = border; s.line.width = bw
    else:
        s.line.fill.background()
    return s


def txt(slide, text, x, y, w, h,
        size=12, bold=False, italic=False,
        color=INK, align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.word_wrap = wrap
    tf = tb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color
    return tb


def logo(slide, x=Inches(11.9), y=Inches(0.18), w=Inches(1.2), h=Inches(0.5)):
    """Place logo — uses black bg area in top-right corner box."""
    rect(slide, x - Inches(0.1), y - Inches(0.08),
         w + Inches(0.2), h + Inches(0.14), fill=INK)
    slide.shapes.add_picture(LOGO, x, y, w, h)


def top_bar(slide, color=RED, height=Inches(0.055)):
    """Thin brand accent line at top."""
    rect(slide, 0, 0, W, height, fill=color)


def slide_title(slide, title, subtitle=None, y=Inches(0.22)):
    """Left-aligned title block below top bar."""
    top_bar(slide)
    logo(slide)
    txt(slide, title, Inches(0.55), y, Inches(10.8), Inches(0.62),
        size=28, bold=True, color=NAVY)
    if subtitle:
        txt(slide, subtitle, Inches(0.55), y + Inches(0.62),
            Inches(10.8), Inches(0.35), size=13, color=GRAY)
    # thin divider under title
    rect(slide, Inches(0.55), y + Inches(0.98) + (Inches(0.35) if subtitle else 0),
         Inches(12.23), Inches(0.025), fill=BORDER)


def card(slide, x, y, w, h, head=None, head_color=NAVY, lines=None, line_size=11):
    """Clean card with optional header strip and bullet lines."""
    rect(slide, x, y, w, h, fill=LIGHT, border=BORDER)
    cy = y
    if head:
        rect(slide, x, y, w, Inches(0.42), fill=head_color)
        txt(slide, head, x + Inches(0.18), y + Inches(0.07),
            w - Inches(0.25), Inches(0.3),
            size=12, bold=True, color=WHITE)
        cy += Inches(0.48)
    if lines:
        for line in lines:
            txt(slide, f"  •   {line}", x + Inches(0.12), cy,
                w - Inches(0.2), Inches(0.32),
                size=line_size, color=INK)
            cy += Inches(0.31)
    return cy


def pill(slide, label, x, y, fill=NAVY, size=10):
    w = Inches(max(len(label) * 0.095 + 0.35, 1.6))
    h = Inches(0.28)
    r = rect(slide, x, y, w, h, fill=fill)
    r.line.fill.background()
    tf = r.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    rn = p.add_run(); rn.text = label
    rn.font.size = Pt(size); rn.font.bold = True
    rn.font.color.rgb = WHITE
    return w


def slide_number(slide, n, total=9):
    txt(slide, f"{n} / {total}",
        Inches(12.6), Inches(7.18), Inches(0.6), Inches(0.25),
        size=9, color=BORDER, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)

# left red accent column
rect(s, 0, 0, Inches(0.22), H, fill=RED)

# big title block
txt(s, "CPPL", Inches(0.55), Inches(1.6), Inches(9), Inches(1.4),
    size=72, bold=True, color=NAVY)
txt(s, "Inbound Process Automation",
    Inches(0.58), Inches(2.9), Inches(9), Inches(0.65),
    size=26, color=INK)
txt(s, "Prototype Planning  —  Data Requirements",
    Inches(0.58), Inches(3.5), Inches(9), Inches(0.45),
    size=16, color=GRAY)

# thin red rule
rect(s, Inches(0.55), Inches(4.1), Inches(5.5), Inches(0.04), fill=RED)

txt(s, "Webteam Private Limited  ×  TransExcel Consulting",
    Inches(0.58), Inches(4.28), Inches(8), Inches(0.38),
    size=12, color=GRAY)
txt(s, "May 2026",
    Inches(0.58), Inches(4.65), Inches(4), Inches(0.32),
    size=11, color=GRAY)

# right info panel
rect(s, Inches(9.8), Inches(1.4), Inches(3.3), Inches(4.0),
     fill=LIGHT, border=BORDER)
txt(s, "Modules in Scope",
    Inches(10.0), Inches(1.65), Inches(3.0), Inches(0.38),
    size=12, bold=True, color=NAVY)
rect(s, Inches(10.0), Inches(2.05), Inches(2.5), Inches(0.025), fill=BORDER)
for i, m in enumerate([
    "M01  Unique Item Code",
    "      Creation",
    "",
    "M02  Material Inward",
    "      & GRN",
]):
    txt(s, m, Inches(10.0), Inches(2.2 + i * 0.38),
        Inches(2.9), Inches(0.36), size=11, color=INK,
        bold=(m.startswith("M0")))

txt(s, "Integration",
    Inches(10.0), Inches(3.85), Inches(3.0), Inches(0.36),
    size=12, bold=True, color=NAVY)
rect(s, Inches(10.0), Inches(4.21), Inches(2.5), Inches(0.025), fill=BORDER)
txt(s, "Tally Prime  ⇄  Our System",
    Inches(10.0), Inches(4.35), Inches(2.9), Inches(0.36),
    size=11, color=INK)

# logo bottom-right on white
slide.shapes.add_picture(LOGO, Inches(10.0), Inches(5.9), Inches(1.4), Inches(0.58)) if False else None
s.shapes.add_picture(LOGO, Inches(10.1), Inches(5.9), Inches(1.4), Inches(0.58))

txt(s, "Confidential  •  Pre-Sales",
    Inches(0.55), Inches(7.1), Inches(5), Inches(0.3),
    size=9, italic=True, color=BORDER)
slide_number(s, 1)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Agenda
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "Agenda", "What we will cover in this session")
logo(s)

items = [
    ("01", "What We Are Building",       "Overview of M01 and M02"),
    ("02", "How Tally Prime Connects",    "Data flow — pull, enrich, push"),
    ("03", "Data Needed  —  Item Code",   "What we need for Module 1"),
    ("04", "Data Needed  —  GRN",        "What we need for Module 2"),
    ("05", "Minimum to Start Prototype", "5 inputs to kick off immediately"),
    ("06", "Next Steps",                 "Actions, owners, timeline"),
]

for i, (num, title, sub) in enumerate(items):
    col, row = i % 2, i // 2
    x = Inches(0.55 + col * 6.42)
    y = Inches(1.65 + row * 1.75)
    rect(s, x, y, Inches(6.0), Inches(1.5), fill=WHITE, border=BORDER)
    # left number tab
    rect(s, x, y, Inches(0.7), Inches(1.5), fill=NAVY if col == 0 else RED)
    txt(s, num, x, y + Inches(0.52), Inches(0.7), Inches(0.45),
        size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, title, x + Inches(0.85), y + Inches(0.32),
        Inches(4.95), Inches(0.45), size=14, bold=True, color=INK)
    txt(s, sub, x + Inches(0.85), y + Inches(0.82),
        Inches(4.95), Inches(0.5), size=11, color=GRAY)

slide_number(s, 2)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — What We Are Building
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "What We Are Building", "Phase I prototype — 2 modules")
logo(s)

for ci, (title, col, points, tag_l) in enumerate([
    (
        "M01   Unique Item Code Creation", NAVY,
        [
            "Hierarchical: Type  →  Category  →  Class  →  Characteristics",
            "Distinguishes Service vs Goods; RM vs SFG vs FG",
            "Alphanumeric code, fixed length, intelligently structured",
            "Deduplication — no duplicate items regardless of source",
            "Separate description field (human-readable)",
            "New items synced back to Tally as Stock Items",
        ],
        [("Pull: Stock Items & Groups from Tally", GREEN),
         ("Push: New items back to Tally", ORANGE)],
    ),
    (
        "M02   Material Inward & GRN", RED,
        [
            "Receive material against open Purchase Orders from Tally",
            "Auto-fill lines from PO  (item, qty, rate, UOM)",
            "QC / Inspection checklist  —  pass or fail per line",
            "Split Accepted qty vs Rejected qty",
            "Partial delivery tracking  —  pending qty stays open",
            "Attach Challan, Invoice, E-way Bill",
            "On approval  →  push Receipt Note back to Tally",
        ],
        [("Pull: POs, Suppliers, Godowns from Tally", GREEN),
         ("Push: Receipt Note back to Tally", ORANGE)],
    ),
]):
    x = Inches(0.4 + ci * 6.52)
    w = Inches(6.22)
    rect(s, x, Inches(1.6), w, Inches(5.65), fill=WHITE, border=BORDER)
    rect(s, x, Inches(1.6), w, Inches(0.48), fill=col)
    txt(s, title, x + Inches(0.18), Inches(1.66),
        w - Inches(0.25), Inches(0.38), size=13, bold=True, color=WHITE)

    cy = Inches(2.2)
    for pt in points:
        txt(s, f"  •   {pt}", x + Inches(0.18), cy,
            w - Inches(0.28), Inches(0.32), size=11, color=INK)
        cy += Inches(0.33)

    cy = max(cy, Inches(6.5)) + Inches(0.05)
    for label, tc in tag_l:
        pill(s, label, x + Inches(0.18), cy, fill=tc, size=10)
        cy += Inches(0.34)

slide_number(s, 3)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Tally Prime Connection
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "How Tally Prime Connects",
            "Data flow between Tally Prime and our system")
logo(s)

cols_data = [
    ("PULL from Tally", GREEN,
     ["Supplier Ledgers\n(Sundry Creditors)",
      "Open Purchase Orders\n& line items",
      "Stock Items &\nStock Groups",
      "Godowns (Warehouses)",
      "Units of Measure",
      "Batches / Lots\n(if enabled at CPPL)"]),
    ("OUR SYSTEM ADDS", NAVY,
     ["Approval workflow\nDraft → QC → Submit",
      "QC inspection checklist\nper line item",
      "Accepted vs Rejected\nqty split",
      "Document attachments\nChallan / Invoice / E-way",
      "GRN reference series\n(our own numbering)",
      "Real-time dashboard\n& pending PO tracker"]),
    ("PUSH back to Tally", RED,
     ["Receipt Note voucher\n(inventory auto-updated)",
      "New Stock Items\n(from M01 item codes)",
      "Batch / Lot entries\n(if applicable)",
      "Narration & references\ncarried through",
      "Triggered only after\nGRN approval — not on save"]),
]

col_w = Inches(4.0)
for ci, (heading, hcol, lines) in enumerate(cols_data):
    x = Inches(0.4 + ci * 4.32)
    card(s, x, Inches(1.65), col_w, Inches(5.6),
         head=heading, head_color=hcol, lines=lines, line_size=11)
    # arrow between columns
    if ci < 2:
        txt(s, "→", x + col_w + Inches(0.08),
            Inches(4.3), Inches(0.22), Inches(0.5),
            size=20, color=BORDER, align=PP_ALIGN.CENTER)

slide_number(s, 4)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Data: Item Code
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "Data Required  —  Item Code Module (M01)",
            "What we need from CPPL to build and demo this module")
logo(s)

card(s, Inches(0.4), Inches(1.6), Inches(5.95), Inches(2.65),
     head="From Tally Prime  —  Export as Excel / CSV",
     head_color=GREEN,
     lines=["Stock Items list (all existing materials)",
            "Stock Groups / Categories hierarchy",
            "Units of Measure (UOM) list",
            "Any existing numbering pattern used today"])

card(s, Inches(6.98), Inches(1.6), Inches(5.95), Inches(2.65),
     head="From CPPL Team  —  Discussion / Verbal",
     head_color=NAVY,
     lines=["How do you currently classify materials?",
            "Do you distinguish RM / SFG / FG today?",
            "5–10 sample items used daily (real names)",
            "Do you track Batch / Heat numbers?"])

# why-it-matters strip
rect(s, Inches(0.4), Inches(4.5), Inches(12.55), Inches(0.04), fill=RED)
txt(s, "Why this data matters",
    Inches(0.4), Inches(4.62), Inches(4), Inches(0.36),
    size=12, bold=True, color=NAVY)

reasons = [
    ("Prevents duplicates",
     "Algorithm checks against existing Tally items — needs the full list."),
    ("Realistic demo",
     "Real CPPL material names make the prototype credible to management."),
    ("Correct hierarchy",
     "Code structure is built around how CPPL actually classifies materials."),
]
for i, (h, d) in enumerate(reasons):
    x = Inches(0.4 + i * 4.3)
    rect(s, x, Inches(5.1), Inches(4.1), Inches(2.1), fill=WHITE, border=BORDER)
    rect(s, x, Inches(5.1), Inches(0.05), Inches(2.1), fill=RED)
    txt(s, h, x + Inches(0.18), Inches(5.18),
        Inches(3.8), Inches(0.38), size=12, bold=True, color=INK)
    txt(s, d, x + Inches(0.18), Inches(5.58),
        Inches(3.8), Inches(1.4), size=11, color=GRAY)

slide_number(s, 5)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Data: GRN
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "Data Required  —  GRN Module (M02)",
            "What we need from CPPL to build and demo this module")
logo(s)

card(s, Inches(0.4), Inches(1.6), Inches(5.95), Inches(3.0),
     head="From Tally Prime  —  Export as Excel / CSV",
     head_color=GREEN,
     lines=["Supplier Ledgers (Sundry Creditors group)",
            "2–3 open Purchase Orders with line items",
            "Godowns list (warehouse / store locations)",
            "Units of Measure",
            "Batch / Lot names (if batch tracking is on)"])

card(s, Inches(6.98), Inches(1.6), Inches(5.95), Inches(3.0),
     head="From CPPL Team  —  Discussion / Verbal",
     head_color=NAVY,
     lines=["Who approves a GRN? (role or person name)",
            "What QC checks are done on inbound material?",
            "Rejections — return to vendor or hold in store?",
            "Sample challan or delivery document (any format)",
            "Do partial deliveries happen frequently?"])

# workflow
rect(s, Inches(0.4), Inches(4.78), Inches(12.55), Inches(0.04), fill=NAVY)
txt(s, "Proposed GRN Workflow",
    Inches(0.4), Inches(4.9), Inches(4), Inches(0.34),
    size=11, bold=True, color=NAVY)

steps = [
    ("Create GRN\n(Draft)",    BORDER,  INK),
    ("Attach\nDocuments",      BORDER,  INK),
    ("QC\nInspection",         BORDER,  INK),
    ("Submit for\nApproval",   BORDER,  INK),
    ("Approved",               GREEN,   WHITE),
    ("Push to\nTally",         RED,     WHITE),
]
sw = Inches(1.9)
for i, (label, bc, tc) in enumerate(steps):
    x = Inches(0.4 + i * 2.12)
    fill = LIGHT if tc == INK else bc
    rect(s, x, Inches(5.35), sw, Inches(1.75),
         fill=fill, border=bc, bw=Pt(1.5))
    txt(s, label, x + Inches(0.1), Inches(5.78),
        sw - Inches(0.15), Inches(0.9),
        size=11, bold=(tc == WHITE), color=tc if tc == WHITE else NAVY,
        align=PP_ALIGN.CENTER)
    if i < 5:
        txt(s, "→", x + sw + Inches(0.06), Inches(6.1),
            Inches(0.2), Inches(0.4),
            size=14, color=GRAY, align=PP_ALIGN.CENTER)

slide_number(s, 6)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Minimum to Start
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "Minimum to Start Prototype",
            "These 5 inputs are enough to begin building immediately")
logo(s)

rows = [
    (RED,   "01", "Stock Items + Stock Groups export from Tally",
     "Excel / CSV  •  Seed data for Item Code module and deduplication check"),
    (RED,   "02", "Supplier Ledgers export from Tally",
     "Sundry Creditors group only  •  Needed for GRN supplier dropdown"),
    (NAVY,  "03", "2–3 sample Purchase Orders",
     "Even dummy POs in CPPL's format  •  To demo the receive-against-PO flow"),
    (NAVY,  "04", "5–10 real material names used daily",
     "Verbal is fine  •  Makes the prototype credible to CPPL management"),
    (GRAY,  "05", "Name of person / role who approves GRN",
     "One name is enough  •  We wire the approval workflow around this role"),
]

for i, (col, num, title, sub) in enumerate(rows):
    y = Inches(1.65 + i * 1.05)
    rect(s, Inches(0.4), y, Inches(12.55), Inches(0.92),
         fill=WHITE, border=BORDER)
    rect(s, Inches(0.4), y, Inches(0.65), Inches(0.92), fill=col)
    txt(s, num, Inches(0.4), y + Inches(0.24),
        Inches(0.65), Inches(0.45),
        size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, title, Inches(1.2), y + Inches(0.1),
        Inches(11.5), Inches(0.38), size=13, bold=True, color=INK)
    txt(s, sub, Inches(1.2), y + Inches(0.52),
        Inches(11.5), Inches(0.32), size=11, color=GRAY)

slide_number(s, 7)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Next Steps
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)
slide_title(s, "Next Steps", "Actions and owners after this call")
logo(s)

headers  = ["#", "Action", "Owner", "By When"]
col_widths = [Inches(0.55), Inches(5.8), Inches(3.4), Inches(2.8)]
table_rows = [
    ["1", "Export Stock Items + Stock Groups from Tally",           "CPPL  Accounts / Store",  "Before next call"],
    ["2", "Export Supplier Ledgers (Sundry Creditors) from Tally",  "CPPL  Accounts",          "Before next call"],
    ["3", "Share 2–3 sample POs (or dummy in same format)",    "CPPL  Purchase Team",     "Before next call"],
    ["4", "Confirm GRN approver role / name",                       "CPPL  Management",        "On this call"],
    ["5", "Share 5–10 daily-use material names",               "CPPL  Store Team",        "On this call"],
    ["6", "Confirm Batch / Lot tracking status in Tally",           "CPPL  Accounts",          "On this call"],
    ["7", "Set up Tally Prime API access for Webteam",              "CPPL IT + Webteam",       "After data received"],
    ["8", "Scaffold project + Tally sync integration",              "Webteam",                 "Within 1 week"],
    ["9", "Prototype demo  —  both modules",                   "Webteam",                 "Within 3 weeks"],
]

row_h = Inches(0.42)
ty = Inches(1.72)

# header row
cx = Inches(0.4)
for ci, hdr in enumerate(headers):
    rect(s, cx, ty, col_widths[ci], row_h, fill=NAVY)
    txt(s, hdr, cx + Inches(0.1), ty + Inches(0.08),
        col_widths[ci] - Inches(0.12), row_h - Inches(0.1),
        size=11, bold=True, color=WHITE)
    cx += col_widths[ci]

for ri, row in enumerate(table_rows):
    ry = ty + row_h * (ri + 1)
    fill = LIGHT if ri % 2 == 0 else WHITE
    cx = Inches(0.4)
    # highlight Webteam rows
    if row[2] == "Webteam":
        fill = RGBColor(0xFF, 0xF0, 0xF1)
    for ci, cell in enumerate(row):
        rect(s, cx, ry, col_widths[ci], row_h,
             fill=fill, border=BORDER, bw=Pt(0.5))
        col = RED if row[2] == "Webteam" and ci in (1, 2, 3) else INK
        txt(s, str(cell), cx + Inches(0.1), ry + Inches(0.08),
            col_widths[ci] - Inches(0.12), row_h - Inches(0.1),
            size=10, color=col)
        cx += col_widths[ci]

slide_number(s, 8)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Thank You
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=WHITE)

# left accent block
rect(s, 0, 0, Inches(0.22), H, fill=RED)
rect(s, 0, 0, W, Inches(0.055), fill=RED)

# large closing line
txt(s, "Thank You.",
    Inches(0.6), Inches(1.5), Inches(8), Inches(1.4),
    size=64, bold=True, color=NAVY)

rect(s, Inches(0.6), Inches(2.85), Inches(4.5), Inches(0.045), fill=RED)

txt(s, "We look forward to building this with CPPL.",
    Inches(0.6), Inches(3.0), Inches(8), Inches(0.5),
    size=16, color=INK)

txt(s, "Questions, clarifications, or additional context\nfrom your side is very welcome on this call.",
    Inches(0.6), Inches(3.55), Inches(7.5), Inches(0.75),
    size=13, color=GRAY)

# contact card
rect(s, Inches(0.6), Inches(4.7), Inches(7.5), Inches(1.8),
     fill=LIGHT, border=BORDER)
contacts = [
    ("Email",   "arvind.warule@webteam.in"),
    ("Project", "CPPL Inbound Process Automation  —  Phase I"),
    ("Repo",    "github.com/AviWarule/CPPL"),
]
for i, (label, val) in enumerate(contacts):
    txt(s, label, Inches(0.85), Inches(4.88 + i * 0.48),
        Inches(1.2), Inches(0.38), size=11, bold=True, color=NAVY)
    txt(s, val, Inches(2.05), Inches(4.88 + i * 0.48),
        Inches(5.8), Inches(0.38), size=11, color=INK)

# logo — prominent on thank you
s.shapes.add_picture(LOGO, Inches(9.5), Inches(2.5), Inches(2.8), Inches(1.18))
rect(s, Inches(9.4), Inches(2.42),
     Inches(3.0), Inches(1.35), fill=INK)
s.shapes.add_picture(LOGO, Inches(9.5), Inches(2.5), Inches(2.8), Inches(1.18))

txt(s, "Confidential  •  Pre-Sales",
    Inches(0.6), Inches(7.1), Inches(5), Inches(0.3),
    size=9, italic=True, color=BORDER)
slide_number(s, 9)


# ── save ─────────────────────────────────────────────────────────────────────
out = r"C:\Users\Dell\CPPL\CPPL_Prototype_DataRequirements_v2.pptx"
prs.save(out)
print(f"Saved: {out}")
