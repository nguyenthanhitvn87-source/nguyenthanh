"""Fill the LPNVi content into the DKSH template skeleton.

skeleton.pptx keeps the original deck's master, layouts and slide artwork;
only the text runs and the extra "SAP" label are written here.
"""
from copy import deepcopy

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _template_para(tf):
    """First paragraph that carries a run — its formatting is the model."""
    for p in tf.paragraphs:
        if p.runs:
            return p._p
    return tf.paragraphs[0]._p


def set_lines(tf, lines, size=None):
    """Replace the text frame's paragraphs with `lines`, keeping the original
    paragraph/run formatting. "" gives a blank spacer paragraph; "\n" inside a
    line becomes a soft line break."""
    tmpl = deepcopy(_template_para(tf))
    body = tf._txBody
    for p in body.findall(qn("a:p")):
        body.remove(p)

    run_tmpl = deepcopy(tmpl.findall(qn("a:r"))[0])
    for child in list(tmpl):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            tmpl.remove(child)

    for line in lines:
        para = deepcopy(tmpl)
        if line:
            for i, chunk in enumerate(line.split("\n")):
                if i:
                    para.append(para.makeelement(qn("a:br"), {}))
                r = deepcopy(run_tmpl)
                r.find(qn("a:t")).text = chunk
                if size is not None:
                    rPr = r.find(qn("a:rPr"))
                    if rPr is None:
                        rPr = r.makeelement(qn("a:rPr"), {})
                        r.insert(0, rPr)
                    rPr.set("sz", str(int(size * 100)))
                para.append(r)
        body.append(para)


def by_name(slide):
    return {sh.name: sh for sh in slide.shapes}


prs = Presentation("skeleton.pptx")
S = list(prs.slides)

# ---------------------------------------------------------------- slide 1
s = by_name(S[0])
title = s["Title 1"]
title.height = Inches(1.70)          # subtitle starts at 3.51", so 1.13+1.70 clears it
set_lines(title.text_frame,
          ["PROCESS LONG PRODUCTION NAME WITH VIETNAMESE LANGUAGE – LPNVi"],
          size=28)
set_lines(s["Subtitle 2"].text_frame, [
    "Business Owner: Luong Ngoc Truong, HEC Sales Manager",
    "Process Owner: Tran Quang Thuan, IT CSSC Vietnam",
    "Technical Owner: Nguyen Vu Thanh, IT CSSC Vietnam",
])
set_lines(s["Text Placeholder 3"].text_frame, ["HCMC, 21 Aug 2026"])

icon = s["Picture 4"]
W = Inches(2.6)
box = S[0].shapes.add_textbox(int(icon.left + icon.width / 2 - W / 2),
                              int(icon.top + icon.height + Inches(0.10)),
                              int(W), Inches(0.70))
box.name = "SAP Label"
tf = box.text_frame
tf.word_wrap = False
tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
tf.vertical_anchor = MSO_ANCHOR.TOP
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "SAP"
run.font.name = "Arial"
run.font.size = Pt(36)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# ---------------------------------------------------------------- slide 2
s = by_name(S[1])
set_lines(s["Title 1"].text_frame, ["Solution Overview"])
cards = [
    ("TextBox 7",  "TextBox 8",  "Rounded Rectangle 9",
     "Master Data\nSetup", "CMIR · MG0156 · INSP", "MASTER DATA"),
    ("TextBox 15", "TextBox 16", "Rounded Rectangle 17",
     "Billing\nOutput", "Mat. Desc + INSP text", "SAP"),
    ("TextBox 23", "TextBox 24", "Rounded Rectangle 25",
     "Delivery\nForms", "PGN · PXKDL · PXKNB", "PP#"),
    ("TextBox 31", "TextBox 32", "Rounded Rectangle 33",
     "E-Invoice\nDisplay", "Separator \"| \" on EIV", "VNPT"),
]
for title_n, cap_n, badge_n, title, cap, badge in cards:
    set_lines(s[title_n].text_frame, [title])
    set_lines(s[cap_n].text_frame, [cap])
    set_lines(s[badge_n].text_frame, [badge])
set_lines(s["TextBox 34"].text_frame,
          ["Scope: all BU nationwide  —  logic priority CMIR › MG0156 › INSP."])

# ---------------------------------------------------------------- slide 3
s = by_name(S[2])
set_lines(s["Title 1"].text_frame, ["Hiện trạng – Current limitation"])
set_lines(s["Text Placeholder 2"].text_frame, ["1"])
set_lines(s["Subtitle 2"].text_frame, [
    "Field Mat. Desc is limited to 40 characters — too short for long product names.",
    "",
    "Custom logic already enhanced: CMIR for Tender, Basis Text for MG0156.",
    "",
    "All remaining products still cannot carry the full Vietnamese name.",
], size=22)
set_lines(s["TextBox 8"].text_frame, ["MAT. DESC"])
set_lines(s["Rounded Rectangle 9"].text_frame, ["40 CHAR"])
set_lines(s["TextBox 12"].text_frame, ["Standard SAP field limit"])
set_lines(s["TextBox 13"].text_frame, ["Enhanced already:\nCMIR (Tender) · MG0156"])
set_lines(s["TextBox 14"].text_frame, ["No Vietnamese diacritics"])
set_lines(s["TextBox 16"].text_frame, ["Rest of scope still open"])
set_lines(s["TextBox 17"].text_frame, ["CURRENT  ·  Field Limitation"])
set_lines(s["Rounded Rectangle 18"].text_frame, ["AS-IS"])

# ---------------------------------------------------------------- slide 4
s = by_name(S[3])
set_lines(s["Title 1"].text_frame, ["Yêu cầu – Business requirement"])
set_lines(s["Text Placeholder 2"].text_frame, ["2"])
set_lines(s["Subtitle 2"].text_frame, [
    "Product name longer than the standard 40 characters, with Vietnamese diacritics.",
    "",
    "On the invoice: Vietnamese name on line 1, English name on line 2 "
    "(smaller font on EIV).",
    "",
    "Extend to Phiếu Giao Nhận, Phiếu Xuất Kho Đại Lý, Phiếu Xuất Kho Nội Bộ.",
    "",
    "Scope: apply for all BU nationwide.",
], size=20)
set_lines(s["TextBox 8"].text_frame, ["HÓA ĐƠN / EIV"])
set_lines(s["Rounded Rectangle 9"].text_frame, ["LAYOUT"])
set_lines(s["TextBox 12"].text_frame, ["Tên tiếng Việt — dòng 1"])
set_lines(s["TextBox 13"].text_frame, ["English name — dòng 2\nchữ nhỏ hơn trên EIV"])
set_lines(s["TextBox 14"].text_frame, ["PGN · PXKDL · PXKNB"])
set_lines(s["TextBox 16"].text_frame, ["All BU nationwide"])
set_lines(s["TextBox 17"].text_frame, ["REQUIREMENT  ·  Display Format"])
set_lines(s["Rounded Rectangle 18"].text_frame, ["TO-BE"])

# ---------------------------------------------------------------- slide 5
s = by_name(S[4])
set_lines(s["Title 1"].text_frame, ["Giải pháp – Solution design"])
set_lines(s["Text Placeholder 2"].text_frame, ["3"])
set_lines(s["Subtitle 2"].text_frame, [
    "Use field INSP (max 220) for Vietnamese text with diacritics.",
    "",
    "Concatenate it with Mat. Desc (default) to build the long product name.",
    "",
    "Separator \"| \" joins the two fields (Mat. Desc + INSP text) for EIV VNPT.",
    "",
    "Logic applies in order: CMIR › MG0156 › INSP.",
], size=20)
set_lines(s["TextBox 8"].text_frame, ["INSP TEXT"])
set_lines(s["Rounded Rectangle 9"].text_frame, ["220 CHAR"])
set_lines(s["TextBox 12"].text_frame, ["Mat. Desc  |  INSP"])
set_lines(s["TextBox 13"].text_frame, ["Separator \"| \""])
set_lines(s["TextBox 15"].text_frame, ["Vietnamese with diacritics"])
set_lines(s["TextBox 16"].text_frame, ["Priority: CMIR › MG0156 › INSP"])
set_lines(s["TextBox 17"].text_frame, ["SOLUTION  ·  Field Concatenation"])
set_lines(s["Rounded Rectangle 18"].text_frame, ["DESIGN"])
# solution slide carries the same red badges as slides 3-4, not the grey
# "manual step" colouring it inherited from the source slide
s["Rounded Rectangle 9"].fill.solid()
s["Rounded Rectangle 9"].fill.fore_color.rgb = RGBColor(0xEF, 0x23, 0x3C)
s["Rounded Rectangle 18"].fill.solid()
s["Rounded Rectangle 18"].fill.fore_color.rgb = RGBColor(0xBE, 0x00, 0x28)

# ---------------------------------------------------------------- slide 6
s = by_name(S[5])
set_lines(s["Title 1"].text_frame, ["Lập trình – Development scope"])
set_lines(s["Text Placeholder 2"].text_frame, ["4"])
set_lines(s["Subtitle 2"].text_frame, [
    "Billing Output @SAP — concatenate Mat. Desc + INSP with separator \"| \".",
    "",
    "Form PGN / PXKDL / PXKNB @PP# — print the long product name.",
    "",
    "@VNPT — detect the separator \"| \" and lay out the long name on EIV.",
], size=22)
set_lines(s["TextBox 8"].text_frame, ["DEV SCOPE"])  # short: badge sits at 11.23"
set_lines(s["Rounded Rectangle 9"].text_frame, ["BUILD"])
set_lines(s["TextBox 12"].text_frame, ["Billing Output"])
set_lines(s["TextBox 13"].text_frame, ["@SAP"])
set_lines(s["TextBox 15"].text_frame, ["Forms PGN/PXKDL/PXKNB — @PP#"])
set_lines(s["TextBox 16"].text_frame, ["Separator handling — @VNPT"])
set_lines(s["TextBox 17"].text_frame, ["DEVELOPMENT  ·  Three Parties"])
set_lines(s["Rounded Rectangle 18"].text_frame, ["IN PROGRESS"])

# ---------------------------------------------------------------- slide 7
s = by_name(S[6])
set_lines(s["Title 1"].text_frame, ["Tóm lại – Summary"])
set_lines(s["TextBox 2"].text_frame, ["CMIR › MG0156 › INSP"])  # one line, clears the icon
table = next(sh for sh in S[6].shapes if sh.has_table).table
rows = [
    ["#", "Item", "Owner", "Mode"],
    ["1", "Scope: all BU nationwide", "Business", "Confirmed"],
    ["2", "Master data: CMIR, MG0156, INSP", "Master Data", "Setup"],
    ["3", "Priority: CMIR › MG0156 › INSP", "SAP", "Develop"],
    ["4", "Dentsply: remove CMIR records first", "CIT", "Action"],
]
for r, values in enumerate(rows):
    for c, value in enumerate(values):
        set_lines(table.cell(r, c).text_frame, [value])

prs.save("lpnvi.pptx")
print("saved lpnvi.pptx")
