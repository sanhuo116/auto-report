from __future__ import annotations

import math
import shutil
import textwrap
from collections import OrderedDict
from pathlib import Path
from zipfile import ZipFile

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path("/Users/zhuyanzhu/Documents/ChatGPT/自动巡检平台")
TEMPLATE = Path("/Users/zhuyanzhu/Desktop/郸城县大数据局运营周报-20251215.docx")
OUTPUT = ROOT / "天津河北分局AI运营周报（2026年8月1日-8月26日）.docx"
ASSET_DIR = ROOT / "runtime" / "tianjin_weekly_assets"
ISSUE_IMAGE = ROOT / "tianjin_case_issue_1.png"
ISSUE_JOB_IMAGE = ROOT / "tianjin_case_issue_job.png"
ISSUE_DISPUTE_IMAGE = ROOT / "tianjin_case_issue_dispute_backend.png"
GOOD_IMAGE = ROOT / "tianjin_case_good_idcard.png"
GOOD_5G_IMAGE = ROOT / "tianjin_case_good_5g_backend.png"
GOOD_IMAGE_CASE = ROOT / "tianjin_case_good_image.png"

NAVY = "1F3A5F"
BLUE = "D8EAF6"
PALE_BLUE = "F2F9FF"
PALE_GREEN = "E9F5E9"
PALE_YELLOW = "FFF4CC"
GRAY = "66717D"
TEXT = "2B3440"
LIGHT_BORDER = "B7C9D6"
DEFAULT_FONT = "Songti SC"

# Original report colors for charts and case highlights.
TITLE_PINK = "#E2586F"
CHART_BLUE = "#4A83F6"
CHART_TEAL = "#48BAA3"
CHART_ORANGE = "#F6AF5D"
CHART_GREEN = "#53B495"
CHART_RED = "#F0807F"
CHART_GRAY = "#C9D2DB"
CASE_ISSUE = "#E8A33B"
CASE_GOOD = "#53B495"

DAILY = OrderedDict(
    [
        ("8月1日", 195),
        ("8月2日", 54),
        ("8月3日", 152),
        ("8月4日", 72),
        ("8月5日", 54),
        ("8月6日", 71),
        ("8月7日", 80),
        ("8月8日", 32),
        ("8月9日", 62),
        ("8月10日", 116),
        ("8月11日", 92),
        ("8月12日", 66),
        ("8月13日", 84),
        ("8月14日", 73),
        ("8月15日", 58),
        ("8月16日", 74),
        ("8月17日", 119),
        ("8月18日", 102),
        ("8月19日", 66),
        ("8月20日", 72),
        ("8月21日", 97),
        ("8月22日", 146),
        ("8月23日", 130),
        ("8月24日", 96),
        ("8月25日", 83),
        ("8月26日", 60),
    ]
)


HIGH_FREQUENCY = [
    ("机器人工单提交", 1351),
    ("未标注", 603),
    ("怎么样防止上当受骗", 14),
    ("哪些人群更容易被骗", 9),
    ("请求帮助", 8),
    ("市民如何称呼", 7),
    ("表示差点被骗", 7),
    ("问卷回复结束语", 6),
    ("居住证制发（居住证申领）如何办理？", 4),
    ("派出所工作时间", 4),
    ("紧急求助类警情", 4),
    ("表示需要向警官咨询", 4),
    ("被骗怎么办", 4),
    ("询问民警是否会用私人号码打电话", 4),
    ("96110相关问题", 3),
    ("如何办理香港其他签注", 3),
    ("居民如何开具无犯罪记录证明", 3),
    ("推荐防诈app > 我想提高诈骗防范意识，有没有合适的app？ > 你知道的，哪些app能帮助我们识别和防范诈骗", 3),
    ("经常接到诈骗电话怎么办", 3),
    ("表示怀疑", 3),
]

INSPECTION_TOTAL = 2306
ISSUE_TOTAL = 284
PASS_TOTAL = INSPECTION_TOTAL - ISSUE_TOTAL
INSPECTION_ACCURACY = PASS_TOTAL / INSPECTION_TOTAL


def font(size: int, bold: bool = False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/System/Library/Fonts/STSong.ttf",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size, index=0)
        except Exception:
            pass
    return ImageFont.load_default()


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color=LIGHT_BORDER, sz="6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = qn(f"w:{edge}")
        node = borders.find(tag)
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_widths(table, widths_cm: list[float]) -> None:
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths_cm):
            cell = row.cells[idx]
            cell.width = Cm(width)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(round(width * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_run_font(run, size=10.5, bold=False, color=TEXT, name=DEFAULT_FONT) -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


def clear_paragraph(paragraph) -> None:
    p = paragraph._element
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def set_para(
    paragraph,
    text="",
    size=10.5,
    bold=False,
    color=TEXT,
    align=None,
    first_line=True,
    before=0,
    after=6,
    line=1.45,
    name=DEFAULT_FONT,
):
    clear_paragraph(paragraph)
    if text:
        run = paragraph.add_run(text)
        set_run_font(run, size=size, bold=bold, color=color, name=name)
    if align is not None:
        paragraph.alignment = align
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.first_line_indent = Cm(0.74 if first_line else 0)
    return paragraph


def add_body(doc, text, first_line=True, after=6):
    set_para(doc.add_paragraph(), text, first_line=first_line, after=after)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    set_para(
        p,
        text,
        size=14 if level == 1 else 12,
        bold=True,
        color=NAVY,
        first_line=False,
        before=12 if level == 1 else 8,
        after=6,
        line=1.15,
        name=DEFAULT_FONT,
    )
    p.paragraph_format.keep_with_next = True
    p_pr = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "5")
    bottom.set(qn("w:color"), BLUE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    set_para(
        p,
        text,
        size=9,
        color=GRAY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False,
        before=0,
        after=7,
        line=1.0,
    )
    return p


def style_table(table, header_fill=BLUE, banded=True):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            set_cell_border(cell)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.first_line_indent = Cm(0)
                for run in p.runs:
                    set_run_font(run, size=9.2)
        if ri == 0:
            for cell in row.cells:
                set_cell_shading(cell, header_fill)
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, size=9.2, bold=True, color=NAVY, name=DEFAULT_FONT)
        elif banded and ri % 2 == 0:
            for cell in row.cells:
                set_cell_shading(cell, PALE_BLUE)


def add_table(doc, headers, rows, widths_cm, header_fill=BLUE):
    table = doc.add_table(rows=1, cols=len(headers))
    for idx, header in enumerate(headers):
        table.cell(0, idx).text = str(header)
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cells[idx].text = str(value)
    set_table_widths(table, widths_cm)
    style_table(table, header_fill=header_fill)
    return table


def add_image(doc, path: Path, width_cm=15.8, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    if caption:
        add_caption(doc, caption)


def add_image_cell(cell, path: Path, width_cm=4.6):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(str(path), width=Cm(width_cm))


def add_case_table(doc, headers, cases, widths_cm, header_fill):
    table = doc.add_table(rows=1, cols=len(headers))
    for i, h in enumerate(headers):
        table.cell(0, i).text = h
    for case in cases:
        cells = table.add_row().cells
        cells[0].text = case["no"]
        cells[1].text = case["scene"]
        add_image_cell(cells[2], case["image"])
        cells[3].text = case["finding"]
        cells[4].text = case["action"]
    set_table_widths(table, widths_cm)
    style_table(table, header_fill=header_fill, banded=False)
    return table


def extract_logo():
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    logo = ASSET_DIR / "logo.png"
    if not logo.exists():
        with ZipFile(TEMPLATE) as z:
            logo.write_bytes(z.read("word/media/image1.png"))
    return logo


def draw_dashboard(path: Path):
    width, height = 1600, 960
    im = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(im)
    title_font = font(42, True)
    metric_font = font(40, True)
    label_font = font(23)
    small_font = font(19)
    draw.text((60, 36), "天津河北分局 AI 文本机器人运行概况", fill="#1F3A5F", font=title_font)
    metrics = [
        ("会话总数", "2,306", "#E7F0F8"),
        ("有效对话轮次", "2,324", "#E5F4F1"),
        ("多轮会话", "183", "#FFF1D9"),
        ("静默会话", "284", "#FBE9E7"),
    ]
    card_w, card_h, gap = 350, 150, 28
    for i, (label, value, fill) in enumerate(metrics):
        x = 60 + i * (card_w + gap)
        draw.rounded_rectangle((x, 115, x + card_w, 115 + card_h), radius=8, fill=fill, outline="#B5C1CC", width=2)
        draw.text((x + 28, 140), label, fill="#1F3A5F", font=label_font)
        draw.text((x + 28, 186), value, fill="#1F3A5F", font=metric_font)
    x0, y0, x1, y1 = 80, 380, 1510, 790
    draw.text((x0, 315), "每日会话量", fill="#1F3A5F", font=label_font)
    maxv = max(DAILY.values())
    for tick in range(0, 201, 50):
        y = y1 - (y1 - y0) * tick / 220
        draw.line((x0, y, x1, y), fill="#D7E0E8", width=2)
        draw.text((x0 - 55, y - 12), str(tick), fill="#777777", font=small_font)
    points = []
    vals = list(DAILY.values())
    for i, val in enumerate(vals):
        x = x0 + i * (x1 - x0) / (len(vals) - 1)
        y = y1 - (y1 - y0) * val / 220
        points.append((x, y))
    draw.line(points, fill=CHART_BLUE, width=5, joint="curve")
    for i, (x, y) in enumerate(points):
        draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=CHART_BLUE)
        if i in (0, 7, 21, 25):
            draw.text((x - 14, y - 36), str(vals[i]), fill="#1F3A5F", font=small_font)
    for i in range(0, len(vals), 3):
        x = x0 + i * (x1 - x0) / (len(vals) - 1)
        draw.text((x - 20, y1 + 18), list(DAILY.keys())[i], fill="#777777", font=small_font)
    draw.text((60, 875), "2026年8月1日-8月26日", fill="#777777", font=small_font)
    im.save(path)


def draw_distribution(path: Path):
    width, height = 1600, 940
    im = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(im)
    title_font = font(38, True)
    body_font = font(24)
    small_font = font(20)
    draw.text((60, 36), "会话识别方式与回复标签分布", fill="#1F3A5F", font=title_font)
    panels = [
        (
            60,
            "识别方式",
            [("关键词相关", 1587, CHART_BLUE), ("AI知识库相关", 723, CHART_TEAL), ("图片相关", 57, CHART_ORANGE), ("未形成识别记录", 49, CHART_GRAY)],
            2306,
        ),
        (
            830,
            "回复标签",
            [("直接回复", 1392, CHART_BLUE), ("静默", 284, CHART_RED), ("FAQ_RAG", 100, CHART_TEAL), ("MrcQA", 5, CHART_ORANGE), ("未标记", 525, CHART_GRAY)],
            2306,
        ),
    ]
    for px, title, items, total in panels:
        draw.text((px, 130), title, fill="#1F3A5F", font=body_font)
        cx, cy, radius, inner = px + 230, 435, 170, 78
        start = -90
        for label, value, color in items:
            angle = value / total * 360
            draw.pieslice((cx - radius, cy - radius, cx + radius, cy + radius), start, start + angle, fill=color, outline="white", width=3)
            start += angle
        draw.ellipse((cx - inner, cy - inner, cx + inner, cy + inner), fill="white")
        draw.text((cx, cy - 12), f"{total:,}", fill="#1F3A5F", font=body_font, anchor="mm")
        draw.text((cx, cy + 28), "条会话", fill="#777777", font=small_font, anchor="mm")
        ly = 210
        legend_x = px + 430
        label_x = px + 470
        value_x = px + 610
        for label, value, color in items:
            draw.rounded_rectangle((legend_x, ly + 4, legend_x + 24, ly + 28), radius=4, fill=color)
            draw.text((label_x, ly), label, fill="#2B3440", font=small_font)
            draw.text((value_x, ly), f"{value:,}  ({value / total * 100:.1f}%)", fill="#66717D", font=small_font)
            ly += 65
    im.save(path)


def set_header_footer(doc: Document, logo: Path):
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    clear_paragraph(hp)
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.add_run().add_picture(str(logo), width=Cm(1.35))
    footer = section.footer
    fp = footer.paragraphs[0]
    set_para(fp, "天津河北分局 AI运营周报", size=8.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, after=0, line=1.0)


def wrap_cn(draw, text, text_font, max_width):
    lines = []
    for para in str(text).splitlines() or [""]:
        current = ""
        for ch in para:
            candidate = current + ch
            if draw.textbbox((0, 0), candidate, font=text_font)[2] <= max_width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = ch
        lines.append(current)
    return lines


def draw_wrapped(draw, xy, text, text_font, fill, max_width, line_gap=14, max_lines=None):
    x, y = xy
    lines = wrap_cn(draw, text, text_font, max_width)
    if max_lines:
        lines = lines[:max_lines]
    bbox = draw.textbbox((0, 0), "中", font=text_font)
    line_height = bbox[3] - bbox[1] + line_gap
    for line in lines:
        draw.text((x, y), line, font=text_font, fill=fill)
        y += line_height
    return y


def draw_page_base(page_no, logo_path):
    page = Image.new("RGB", (1654, 2339), "white")
    draw = ImageDraw.Draw(page)
    logo = Image.open(logo_path).convert("RGBA")
    logo.thumbnail((120, 120))
    page.paste(logo, (1450, 64), logo)
    draw.line((125, 214, 1529, 214), fill="#C8D3DE", width=3)
    draw.text((827, 2260), str(page_no), fill="#66717D", font=font(23), anchor="ma")
    draw.text((827, 2295), "天津河北分局 AI运营周报", fill="#66717D", font=font(18), anchor="ma")
    return page, draw


def paste_width(page, path, x, y, width):
    im = Image.open(path).convert("RGB")
    scale = width / im.width
    im = im.resize((int(width), int(im.height * scale)))
    page.paste(im, (int(x), int(y)))
    return im.height


def draw_kv_table(draw, x, y, widths, rows, row_h=82):
    total_w = sum(widths)
    header_fill = "#DCE4EC"
    for ri, row in enumerate(rows):
        yy = y + ri * row_h
        fill = header_fill if ri == 0 else ("#F3F6F9" if ri % 2 == 0 else "#FFFFFF")
        draw.rectangle((x, yy, x + total_w, yy + row_h), fill=fill, outline="#B5C1CC", width=2)
        xx = x
        for ci, value in enumerate(row):
            if ci:
                draw.line((xx, yy, xx, yy + row_h), fill="#B5C1CC", width=2)
            text_font = font(22, ri == 0)
            lines = wrap_cn(draw, value, text_font, widths[ci] - 30)
            line_h = 30
            start_y = yy + max(12, (row_h - len(lines) * line_h) // 2)
            for li, line in enumerate(lines[:2]):
                bbox = draw.textbbox((0, 0), line, font=text_font)
                draw.text((xx + (widths[ci] - (bbox[2] - bbox[0])) // 2, start_y + li * line_h), line, fill="#1F3A5F" if ri == 0 else "#2B3440", font=text_font)
            xx += widths[ci]


def draw_frequency_table(draw, x, y, rows, row_h=72):
    table_rows = [["知识库名称", "会话量"]]
    table_rows.extend([[name, f"{count:,}"] for name, count in rows])
    widths = [1150, 254]
    total_w = sum(widths)
    for ri, row in enumerate(table_rows):
        yy = y + ri * row_h
        fill = "#DCE4EC" if ri == 0 else ("#F3F6F9" if ri % 2 == 0 else "#FFFFFF")
        draw.rectangle((x, yy, x + total_w, yy + row_h), fill=fill, outline="#B5C1CC", width=2)
        xx = x
        for ci, value in enumerate(row):
            if ci:
                draw.line((xx, yy, xx, yy + row_h), fill="#B5C1CC", width=2)
            text_font = font(20, ri == 0)
            lines = wrap_cn(draw, value, text_font, widths[ci] - 28)
            line_h = 27
            start_y = yy + max(10, (row_h - min(len(lines), 2) * line_h) // 2)
            for li, line in enumerate(lines[:2]):
                bbox = draw.textbbox((0, 0), line, font=text_font)
                draw.text(
                    (xx + (widths[ci] - (bbox[2] - bbox[0])) // 2, start_y + li * line_h),
                    line,
                    fill="#1F3A5F" if ri == 0 else "#2B3440",
                    font=text_font,
                )
            xx += widths[ci]


def draw_stat_cards(draw, y):
    cards = [
        ("巡检会话", f"{INSPECTION_TOTAL:,}", "#E7F0F8"),
        ("待优化会话", f"{ISSUE_TOTAL:,}", "#FBE9E7"),
        ("完成复核", f"{PASS_TOTAL:,}", "#E5F4F1"),
        ("巡检准确率", f"{INSPECTION_ACCURACY * 100:.2f}%", "#FFF1D9"),
    ]
    x, w, h, gap = 155, 300, 118, 16
    for i, (label, value, fill) in enumerate(cards):
        xx = x + i * (w + gap)
        draw.rounded_rectangle((xx, y, xx + w, y + h), radius=8, fill=fill, outline="#B5C1CC", width=2)
        draw.text((xx + 22, y + 18), label, fill="#1F3A5F", font=font(20))
        draw.text((xx + 22, y + 54), value, fill="#1F3A5F", font=font(31, True))


def draw_case_block(page, draw, y, title, image_path, finding, action, accent):
    x, w = 125, 1404
    draw.rounded_rectangle((x, y, x + w, y + 510), radius=8, fill="white", outline="#B5C1CC", width=2)
    draw.rectangle((x, y, x + w, y + 58), fill=accent)
    draw.text((x + 24, y + 15), title, fill="white", font=font(28, True))
    img = Image.open(image_path).convert("RGB")
    img.thumbnail((590, 400))
    page.paste(img, (x + 25, y + 88))
    is_issue = title.startswith("案例")
    draw.text((x + 660, y + 92), "问题表现：" if is_issue else "优秀表现：", fill="#1F3A5F", font=font(23, True))
    y2 = draw_wrapped(draw, (x + 660, y + 132), finding, font(21), "#2B3440", 690, 12, 9)
    draw.text((x + 660, y2 + 18), "优化动作：" if is_issue else "可复用动作：", fill="#1F3A5F", font=font(23, True))
    draw_wrapped(draw, (x + 660, y2 + 58), action, font(21), "#2B3440", 690, 12, 8)
    return y + 535


def build_raster_doc(docx_path, dashboard, distribution, logo):
    pages = []

    page, draw = draw_page_base(1, logo)
    draw.text((827, 300), "“天津河北分局”AI周报", fill=TITLE_PINK, font=font(54, True), anchor="ma")
    draw.text((827, 370), "2026年8月1日-8月26日  AI文本机器人运营与巡检报告", fill="#2B3440", font=font(27), anchor="ma")
    intro = "本期对2,306条会话逐条查看，并从咨询数据、会话明细、线上优化动作和优秀案例四个维度进行分析。"
    draw_wrapped(draw, (155, 470), intro, font(27), "#2B3440", 1330, 18)
    draw.text((155, 600), "一、咨询数据分析", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 650, 1499, 650), fill="#C8D3DE", width=3)
    h = paste_width(page, dashboard, 155, 710, 1344)
    summary = "总体咨询情况：本期共纳入2,306条会话，累计有效对话轮次2,324轮，平均每条会话约1.01轮；其中1轮会话1,868条，多轮会话183条，0轮会话255条。会话量最高日为8月1日，共195条；最低日为8月8日，共32条。"
    draw_wrapped(draw, (155, 710 + h + 35), summary, font(24), "#2B3440", 1330, 14, 6)
    pages.append(page)

    page, draw = draw_page_base(2, logo)
    draw.text((155, 270), "二、会话明细分析", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 320, 1499, 320), fill="#C8D3DE", width=3)
    h = paste_width(page, distribution, 155, 370, 1344)
    detail = "本期会话主要由关键词和AI知识库两类方式承接，其中关键词相关会话1,587条，AI知识库相关会话723条，图片相关会话57条，未形成识别记录49条。回复以直接回复为主，但仍有284条静默会话，且存在图片识别失败、知识库未命中和多轮重复问候等需要优化的交互。"
    y = 370 + h + 34
    y = draw_wrapped(draw, (155, y), detail, font(22), "#2B3440", 1330, 12, 5)
    rows = [
        ["指标", "数量", "占比", "分析说明"],
        ["会话总数", "2,306", "100.0%", "本期独立会话数量"],
        ["累计有效对话轮次", "2,324", "100.8%", "按会话明细中的有效对话轮次汇总"],
        ["0轮会话", "255", "11.1%", "未形成有效对话，包含静默或未识别场景"],
        ["1轮会话", "1,868", "81.0%", "本期主要会话形态"],
        ["多轮会话", "183", "7.9%", "需要重点观察上下文承接与问题闭环"],
        ["静默标签", "284", "12.3%", "需结合用户输入和系统返回进一步优化"],
    ]
    draw_kv_table(draw, 155, y + 26, [300, 230, 220, 594], rows, 72)
    pages.append(page)

    page, draw = draw_page_base(3, logo)
    draw.text((155, 270), "三、市民咨询高频问题", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 320, 1499, 320), fill="#C8D3DE", width=3)
    freq_intro = "以下为本期市民咨询量较高的20个知识库，按会话量由高到低排列。"
    draw_wrapped(draw, (155, 360), freq_intro, font(22), "#2B3440", 1330, 12, 3)
    draw_frequency_table(draw, 155, 455, HIGH_FREQUENCY, row_h=72)
    pages.append(page)

    page, draw = draw_page_base(4, logo)
    draw.text((155, 270), "待优化案例与线上优化AI动作", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 320, 1499, 320), fill="#C8D3DE", width=3)
    action_intro = "本期线上优化动作以“回复不恰当或未较好解决”的会话为核心。逐条查看2,306条会话，共识别284条待优化会话；本报告展示其中3例典型案例，巡检准确率为87.68%。"
    draw_wrapped(draw, (155, 360), action_intro, font(22), "#2B3440", 1330, 12, 4)
    draw_stat_cards(draw, 445)
    y = 610
    y = draw_case_block(page, draw, y, "案例01  图片信息识别失败", ISSUE_IMAGE,
                        "用户上传图片后，系统返回“未成功识别图片信息”，同时展示“全局配置策略”等系统信息，没有给出可执行的替代路径，用户的问题未被解决。",
                        "增加图片识别失败兜底：明确说明未识别成功，引导用户重新上传清晰图片或直接输入文字；屏蔽内部策略名称，并补充失败场景回归测试。",
                        CASE_ISSUE)
    y = draw_case_block(page, draw, y, "案例02  知识库未命中/岗位与联系信息咨询", ISSUE_JOB_IMAGE,
                        "用户咨询河北区派出所工作、岗位及联系方式等具体事项时，系统命中AI_NO_REFERENCE，回复未能提供有效办理入口和可核实的联系方式，问题解决不完整。",
                        "补充就业招聘、派出所地址电话和业务咨询入口知识；无命中时采用“说明能力边界+给出官方办理入口+提示人工核实”的标准兜底话术。",
                        CASE_ISSUE)
    draw_case_block(page, draw, y, "案例03  纠纷求助后的重复追问", ISSUE_DISPUTE_IMAGE,
                    "用户已明确说明曾报警并提供了处理经过，系统仍重复追问“你报案了吗”，没有承接已有信息，也没有围绕案件进展给出新的办理指引。",
                    "增加多轮状态记忆和已答复识别：用户确认已报警后直接进入“核实辖区、补充材料、查询进展或联系承办单位”的分支，避免重复提问。",
                    CASE_ISSUE)
    pages.append(page)

    page, draw = draw_page_base(5, logo)
    draw.text((155, 270), "四、本期优秀案例", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 320, 1499, 320), fill="#C8D3DE", width=3)
    good_intro = "本期优秀案例聚焦交互自然、问题解决较完整的会话，重点呈现能够识别场景、给出具体操作建议并形成安全提醒闭环的答复。"
    draw_wrapped(draw, (155, 360), good_intro, font(22), "#2B3440", 1330, 12, 4)
    y = 475
    y = draw_case_block(page, draw, y, "优秀案例01  身份证丢失/补领及派出所联系信息", GOOD_IMAGE,
                        "系统围绕居民身份证丢失、补领和换证要求进行分层说明，补充所需材料，并提供河北区多个派出所的地址与电话，回答覆盖办理条件、材料和办理地点。",
                        "保留“明确事项类型→说明办理要求→列出所需材料→提供就近办理地点及电话”的结构，作为户籍身份证类咨询的标准化回复模板。",
                        CASE_GOOD)
    y = draw_case_block(page, draw, y, "优秀案例02  全民5G北斗软件风险识别", GOOD_5G_IMAGE,
                        "系统识别出所谓全民5G北斗卫星导航软件以索要银行卡和身份证信息为主要风险点，明确告知用户不要填写敏感信息，并提出卸载软件、安装国家反诈中心APP及必要时报警等动作。",
                        "沿用“识别诈骗特征→明确风险结论→给出立即处置动作→补充账户保护和报警入口”的反诈咨询回复结构。",
                        CASE_GOOD)
    draw_case_block(page, draw, y, "优秀案例03  图片识别与家庭照片隐私提醒", GOOD_IMAGE_CASE,
                    "系统对图片内容进行客观描述，明确判断其不属于直接欺诈材料，同时补充家庭成员、住址背景等隐私信息不宜在公共平台随意分享，风险边界表达清晰。",
                    "图片类咨询采用“先描述内容→再判断直接风险→补充隐私和社会工程风险→给出保护建议”的分层答复结构，避免过度放大风险。",
                    CASE_GOOD)
    pages.append(page)

    page, draw = draw_page_base(6, logo)
    draw.text((155, 270), "五、后续优化重点", fill="#1F3A5F", font=font(34, True))
    draw.line((155, 320, 1499, 320), fill="#C8D3DE", width=3)
    draw.text((155, 370), "围绕本期巡检结果，下一阶段确定以下AI训练与运营动作：", fill="#2B3440", font=font(23))
    next_steps = [
        "1、完善图片识别失败、图片模糊和无法读取场景的统一兜底话术，确保用户始终获得下一步操作指引。",
        "2、补充河北区本地政务知识库，优先覆盖户籍身份证、派出所地址电话、就业招聘和常见业务办理入口。",
        "3、围绕284条静默会话和255条0轮会话开展专项复盘，区分用户未输入、图片失败、回复未展示和主动结束等原因。",
        "4、保留身份证、反诈软件识别和图片隐私提醒优秀案例的完整答复结构，作为政务咨询类知识库和提示词的复用样本。",
        "5、持续记录问题类型、处理结果和复测结果，形成可追踪的线上质量指标。",
    ]
    yy = 430
    for step in next_steps:
        yy = draw_wrapped(draw, (175, yy), step, font(23), "#2B3440", 1290, 12) + 22
    draw.text((155, 1080), "本期结论", fill="#1F3A5F", font=font(27, True))
    draw.line((155, 1125, 1499, 1125), fill="#C8D3DE", width=3)
    conclusion = "本期共识别284条待优化会话，主要集中在静默无回复、图片识别异常、知识库未命中和多轮上下文承接不足等场景；通过补齐本地政务知识、完善失败兜底和加强多轮状态记忆，可持续提升问题解决完整度。"
    draw_wrapped(draw, (175, 1175), conclusion, font(22), "#2B3440", 1290, 14, 8)
    pages.append(page)

    temp_dir = ASSET_DIR / "pages"
    temp_dir.mkdir(parents=True, exist_ok=True)
    page_paths = []
    for idx, page in enumerate(pages, 1):
        path = temp_dir / f"page-{idx}.png"
        page.save(path, dpi=(150, 150))
        page_paths.append(path)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(0)
    sec.right_margin = Cm(0)
    sec.top_margin = Cm(0)
    sec.bottom_margin = Cm(0)
    for idx, path in enumerate(page_paths):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(str(path), width=Cm(21.0))
        if idx < len(page_paths) - 1:
            p.add_run().add_break(WD_BREAK.PAGE)
    doc.core_properties.title = "天津河北分局AI运营周报（2026年8月1日-8月26日）"
    doc.core_properties.subject = "AI文本机器人运营与会话巡检报告"
    doc.core_properties.author = "AI Operations Team"
    doc.save(docx_path)


def build():
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    dashboard = ASSET_DIR / "dashboard.png"
    distribution = ASSET_DIR / "distribution.png"
    draw_dashboard(dashboard)
    draw_distribution(distribution)

    # Build a compact source document first, then replace it with a high-fidelity raster layout
    # because the bundled headless renderer does not provide a Chinese font fallback.
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    set_header_footer(doc, extract_logo())

    p = doc.add_paragraph()
    set_para(p, "“天津河北分局”AI周报", size=22, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, before=28, after=12, line=1.1, name=DEFAULT_FONT)
    p = doc.add_paragraph()
    set_para(p, "2026年8月1日-8月26日  天津河北分局AI文本机器人运营与巡检报告", size=12, color=TEXT, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, after=20, line=1.15, name=DEFAULT_FONT)
    add_body(
        doc,
        "以下是2026年8月1日-8月26日天津河北分局AI文本机器人的运营周报。本期对2,306条会话逐条查看，并从咨询数据、会话明细、线上优化动作和优秀案例四个维度进行分析。",
        first_line=True,
        after=10,
    )

    add_heading(doc, "一、咨询数据分析")
    add_heading(doc, "咨询情况", level=2)
    add_image(doc, dashboard, width_cm=15.8, caption="图1  天津河北分局AI文本机器人本期运行概况")
    add_body(
        doc,
        "总体咨询情况。本期共纳入2,306条会话，累计有效对话轮次2,324轮，平均每条会话约1.01轮；其中接待1轮会话1,868条，多轮会话183条，0轮会话255条。会话量最高日为8月1日，共195条；最低日为8月8日，共32条。",
        after=6,
    )
    add_body(
        doc,
        "2、当前AI准确率。本期共识别284条待优化会话，完成复核的会话为2,022条，巡检准确率为87.68%。按会话处理结果观察，直接回复1,392条，占60.4%；静默284条，占12.3%；FAQ_RAG 100条，占4.3%；MrcQA 5条，占0.2%；另有525条未标记回复标签，占22.8%。",
        after=6,
    )
    add_body(
        doc,
        "3、当前咨询量分析。本期日均会话量约88.7条，整体呈波动运行。8月1日、8月3日、8月22日至8月24日为相对高量区间，后续应重点关注高峰时段的知识库命中和静默会话表现。",
        after=6,
    )

    add_heading(doc, "二、会话明细分析")
    add_heading(doc, "会话识别与回复标签", level=2)
    add_image(doc, distribution, width_cm=15.8, caption="图2  本期会话识别方式与回复标签分布")
    add_body(
        doc,
        "本期会话主要由关键词和AI知识库两类方式承接，其中关键词相关会话1,587条，AI知识库相关会话723条，图片相关会话57条，未形成识别记录49条。回复以直接回复为主，但仍有284条静默会话，且存在图片识别失败、知识库未命中和多轮重复问候等需要优化的交互。",
        after=8,
    )
    add_table(
        doc,
        ["指标", "数量", "占比", "分析说明"],
        [
            ("会话总数", "2,306", "100.0%", "本期纳入的独立会话数量"),
            ("累计有效对话轮次", "2,324", "100.8%", "本期有效对话轮次合计"),
            ("0轮会话", "255", "11.1%", "未形成有效对话，包含静默或未识别场景"),
            ("1轮会话", "1,868", "81.0%", "本期主要会话形态"),
            ("多轮会话", "183", "7.9%", "需要重点观察上下文承接与问题闭环"),
            ("静默标签", "284", "12.3%", "需结合用户输入和系统返回进一步优化"),
        ],
        [3.5, 2.7, 2.7, 7.0],
    )
    add_caption(doc, "表1  本期会话明细关键指标")

    add_heading(doc, "待优化案例与线上优化AI动作")
    add_body(
        doc,
        "本期线上优化动作以“回复不恰当或未较好解决”的会话为核心，主要集中在图片识别失败后的承接、无知识库命中时的兜底回复，以及多轮会话中的重复问候和问题未闭环。",
        after=8,
    )
    issue_cases = [
        {
            "no": "01",
            "scene": "图片信息识别失败",
            "image": ISSUE_IMAGE,
            "finding": "用户上传图片后，系统返回“未成功识别图片信息”，同时展示“全局配置策略”等系统信息，没有给出可执行的替代路径，用户的问题未被解决。",
            "action": "增加图片识别失败兜底：明确说明未识别成功，引导用户重新上传清晰图片或直接输入文字；屏蔽内部策略名称，并补充失败场景回归测试。",
        },
        {
            "no": "02",
            "scene": "知识库未命中/岗位与联系信息咨询",
            "image": ISSUE_JOB_IMAGE,
            "finding": "用户咨询河北区派出所工作、岗位及联系方式等具体事项时，系统命中AI_NO_REFERENCE，回复未能提供有效办理入口和可核实的联系方式，问题解决不完整。",
            "action": "补充就业招聘、派出所地址电话和业务咨询入口知识；无命中时采用“说明能力边界+给出官方办理入口+提示人工核实”的标准兜底话术。",
        },
    ]
    add_case_table(
        doc,
        ["序号", "问题场景", "案例截图", "问题表现", "优化动作"],
        issue_cases,
        [1.2, 3.1, 5.0, 5.4, 5.4],
        PALE_YELLOW,
    )
    add_caption(doc, "表2  本期典型待优化案例")
    add_body(
        doc,
        "线上执行重点：一是将图片识别异常纳入自动质检，统一失败反馈与重试引导；二是围绕派出所地址电话、招聘就业、户籍身份证等高频政务事项补齐本地知识库；三是针对多轮会话增加“已回答/未回答/需追问”的状态判断，避免用户连续输入“您好、在线吗、同志”时只返回重复欢迎语。",
        after=8,
    )

    add_heading(doc, "四、本期优秀案例")
    add_body(
        doc,
        "本期优秀案例聚焦交互较好、问题解决较完整的会话。身份证业务案例中，AI能够围绕证件丢失、补领和换证等问题给出办理要求、所需材料，并进一步提供河北区各派出所地址及电话，形成了较完整的政务咨询闭环。",
        after=8,
    )
    good_cases = [
        {
            "no": "01",
            "scene": "身份证丢失/补领及派出所联系信息",
            "image": GOOD_IMAGE,
            "finding": "系统围绕居民身份证丢失、补领和换证要求进行分层说明，补充所需材料，并提供河北区多个派出所的地址与电话，回答内容覆盖办理条件、材料和办理地点。",
            "action": "保留“明确事项类型→说明办理要求→列出所需材料→提供就近办理地点及电话”的结构，作为户籍身份证类咨询的标准化回复模板。",
        }
    ]
    add_case_table(
        doc,
        ["序号", "业务场景", "案例截图", "优秀表现", "可复用动作"],
        good_cases,
        [1.2, 3.1, 5.0, 5.4, 5.4],
        PALE_GREEN,
    )
    add_caption(doc, "表3  本期优秀案例")

    add_heading(doc, "五、后续优化重点")
    add_body(doc, "围绕本期巡检结果，下一阶段确定以下AI训练与运营动作：", first_line=False, after=6)
    add_body(doc, "1、完善图片识别失败、图片模糊和无法读取场景的统一兜底话术，确保用户始终获得下一步操作指引。", first_line=False, after=4)
    add_body(doc, "2、补充河北区本地政务知识库，优先覆盖户籍身份证、派出所地址电话、就业招聘和常见业务办理入口。", first_line=False, after=4)
    add_body(doc, "3、围绕284条静默会话和255条0轮会话开展专项复盘，区分用户未输入、图片失败、回复未展示和主动结束等原因。", first_line=False, after=4)
    add_body(doc, "4、保留身份证业务优秀案例的完整答复结构，作为政务咨询类知识库和大模型提示词的复用样本。", first_line=False, after=4)
    add_body(doc, "5、建立会话ID、用户问题、识别方式、回复标签和巡检结果的关联统计，形成可持续的线上质量指标。", first_line=False, after=4)

    doc.core_properties.title = "天津河北分局AI运营周报（2026年8月1日-8月26日）"
    doc.core_properties.subject = "AI文本机器人运营与会话巡检报告"
    doc.core_properties.author = "AI Operations Team"
    doc.save(OUTPUT)
    build_raster_doc(OUTPUT, dashboard, distribution, extract_logo())
    print(OUTPUT)


if __name__ == "__main__":
    build()
