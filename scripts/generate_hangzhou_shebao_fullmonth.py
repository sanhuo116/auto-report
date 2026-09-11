#!/usr/bin/env python3
"""Build the full-month Hangzhou social-security inspection report."""

from collections import defaultdict
from copy import deepcopy
from pathlib import Path
import math
import re
import zipfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont


ROOT = Path("/Users/zhuyanzhu/Documents/ChatGPT/自动巡检平台")
RUNTIME = ROOT / "auto-report" / "runtime" / "hangzhou_shebao_20260801_31"
TEMPLATE = ROOT / "杭州社保智能接待运营巡检报告（2026年8月1日-8月27日）.docx"
OUTPUT = ROOT / "杭州社保智能接待运营巡检报告（2026年8月1日-8月31日）.docx"

CHAT_XLSX = RUNTIME / "杭州社保202608全月文本聊天记录.xlsx"

RECEPTION_DISPLAY = "约1.88万次"
CONSULTATION_DISPLAY = "约3.9万次"
EFFECTIVE_DISPLAY = "约1.83万次"
ACTIVITY_DISPLAY = "19.3%"
AVG_ROUNDS = "2.1轮"
REVIEW_TOTAL = 1000
BAD_TOTAL = 333
GOOD_TOTAL = REVIEW_TOTAL - BAD_TOTAL
ACCURACY = GOOD_TOTAL / REVIEW_TOTAL * 100
SESSION_TOTAL = 673

TOP20 = [
    ("人工客服", 8972),
    ("人社小灵光", 1134),
    ("问候语", 911),
    ("网报新增回退申请表下载", 660),
    ("业务办理", 598),
    ("个人权益记录查询打印", 584),
    ("个体劳动者参保登记、停保登记", 395),
    ("跨省转移接续", 233),
    ("职工参保登记表格下载", 193),
    ("隶属关系变更", 193),
    ("人工客服工作时间", 152),
    ("停保该怎么操作", 135),
    ("社保缴费基数范围", 135),
    ("当月新增回退", 133),
    ("养老保险重复缴费怎么办", 119),
    ("网上服务系统进入路径", 106),
    ("申请补缴城镇职工社会保险费", 100),
    ("社保参保", 90),
    ("灵活就业人员参保条件", 78),
    ("个人参保信息变更", 73),
]

BAD_CASES = [
    {
        "id": "SS202608020645341536655873",
        "title": "案例一：企业参保后咨询线下办理，追问未形成有效回复",
        "finding": "用户已明确说明企业完成社保登记后咨询后续办理方式，机器人先完成企业参保分类和线上办理说明，但在追问“线下可以办理吗”时未给出线下办理入口、受理条件或办理提醒，问题未闭环。",
        "action": "补充“企业参保后线下办理、窗口办理、首次职工参保”等关联问法；识别为企业参保后，直接说明网上办理、经办机构窗口和首次办理的适用路径，并补充办理材料及受理时间提示。",
    },
    {
        "id": "SS202608021247581688582331",
        "title": "案例二：养老保险缴费年限重复追问，未承接上下文",
        "finding": "用户在已获得跨省转移和养老保险中断影响说明后，连续三次询问“怎么查询养老保险缴费总年限”。机器人均未形成有效回复，也没有沿用前文的养老保险场景给出查询入口，导致同一诉求重复出现。",
        "action": "增加“养老保险缴费总年限、累计缴费年限、跨省转移后如何查询”等组合问法；在上下文中优先识别查询诉求，直接提供浙里办、支付宝或微信浙里办小程序的查询路径，必要时补充证明打印入口。",
    },
    {
        "id": "SS202608261435221267609000",
        "title": "案例三：单位职工停保主体判断不清，未明确办理路径",
        "finding": "用户先咨询“我要把社保断掉”，随后明确表示原单位未办理停保并反复询问能否自行办理。机器人连续多轮未形成有效回复，未区分单位职工与灵活就业人员的停保责任主体，最后转人工。",
        "action": "补充“原单位不给停保、离职后社保断缴、个人能否办理职工停保”等问法；首轮先判断参保身份和劳动关系，再说明单位职工停保由用人单位办理、个人可采取的投诉或咨询路径，并与灵活就业人员停保流程分开回答。",
    },
]

GOOD_CASES = [
    {
        "id": "SS202608020107031354486798",
        "title": "案例一：灵活就业参保停保，办理路径和边界说明完整",
        "finding": "机器人准确识别灵活就业人员参保及停保事项，同时提供浙江政务服务网、浙里办、社保经办机构和社银合作银行等办理路径；对医保办理边界、税务缴费和参保地选择作出提示，信息完整且具有可执行性。",
    },
    {
        "id": "SS202608221058111837978949",
        "title": "案例二：连续咨询多个事项，均给出具体办理入口",
        "finding": "用户连续咨询职工参保登记表格下载、个人权益记录查询和个体劳动者参保停保登记，机器人能够分别匹配对应事项，依次给出表格下载、浙里办查询及线上线下办理路径，连续咨询承接较完整。",
    },
    {
        "id": "SS202608281425171890255710",
        "title": "案例三：待遇领取地咨询，条件识别和结论表达清晰",
        "finding": "机器人围绕户籍地、养老保险关系、性别及缴费地逐步识别条件，并在用户补充信息后给出省内待遇领取地判断规则，能够结合多轮信息推进问题判断，回答结构清晰。",
    },
]


def image_font(size, bold=False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ]
    path = next((p for p in candidates if Path(p).exists()), None)
    if not path:
        return ImageFont.load_default()
    try:
        return ImageFont.truetype(path, size=size, index=0)
    except OSError:
        return ImageFont.load_default()


def set_run_font(run, size=12, bold=False, color=None):
    run.font.name = "宋体"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def clear_document_body(doc):
    body = doc._element.body
    sect_pr = body.sectPr
    for child in list(body):
        if child is not sect_pr:
            body.remove(child)


def add_text(doc, text, size=12, bold=False, align=None, before=0, after=6, color=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, color=color)
    return p


def add_rich_text(doc, prefix, suffix, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(prefix)
    set_run_font(r, size=size, bold=True)
    r = p.add_run(suffix)
    set_run_font(r, size=size)
    return p


def add_heading(doc, text, level=1, new_page=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8 if level == 1 else 5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = new_page
    r = p.add_run(text)
    set_run_font(r, size=16 if level == 1 else 13, bold=True)
    return p


def add_title(doc):
    p = add_text(
        doc,
        "杭州社保智能接待运营巡检报告",
        size=24,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        after=2,
        color=(226, 88, 111),
    )
    p.paragraph_format.space_before = Pt(18)
    add_text(
        doc,
        "（2026年8月1日-8月31日）",
        size=16,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        after=18,
    )


def set_cell(cell, value, header=False, align=WD_ALIGN_PARAGRAPH.LEFT, fill=None):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(value))
    set_run_font(r, size=10.5 if not header else 11, bold=header)
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:fill"), fill or ("D8EAF6" if header else "FFFFFF"))
    tc_mar = tcpr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tcpr.append(tc_mar)
    for side in ("top", "start", "bottom", "end"):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), "100")
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color="B7C9D6", sz="6"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), sz)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def add_metric_image(path):
    w, h = 1320, 610
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    title = image_font(30, True)
    label = image_font(22, True)
    value = image_font(31, True)
    small = image_font(17)
    d.text((35, 25), "杭州社保 2026年8月运行概览", fill=(45, 45, 45), font=title)
    cards = [
        ("接待量", RECEPTION_DISPLAY, (213, 226, 255), (67, 120, 224)),
        ("咨询量", CONSULTATION_DISPLAY, (217, 245, 242), (71, 183, 170)),
        ("有效接待量", EFFECTIVE_DISPLAY, (255, 244, 204), (241, 169, 82)),
        ("客户活跃度", ACTIVITY_DISPLAY, (253, 232, 231), (224, 116, 122)),
        ("平均对话总轮次", AVG_ROUNDS, (232, 239, 255), (111, 118, 210)),
        ("有效平均对话轮次", AVG_ROUNDS, (233, 245, 233), (83, 180, 149)),
        ("复核记录", f"{REVIEW_TOTAL:,}条", (238, 242, 247), (101, 126, 151)),
        ("复核准确率", f"{ACCURACY:.1f}%", (255, 244, 204), (224, 146, 70)),
    ]
    for i, (name, val, bg, accent) in enumerate(cards):
        row, col = divmod(i, 4)
        x = 35 + col * 320
        y = 90 + row * 230
        d.rounded_rectangle((x, y, x + 292, y + 185), radius=10, fill=bg, outline=(220, 225, 232), width=2)
        d.ellipse((x + 22, y + 47, x + 92, y + 117), fill=accent)
        d.text((x + 118, y + 25), name, fill=(50, 50, 50), font=label)
        d.text((x + 118, y + 80), val, fill=(25, 25, 25), font=value)
        d.text((x + 24, y + 142), "本期运行指标", fill=(105, 110, 120), font=small)
    im.save(path)


def add_frequency_image(path):
    w, h = 1320, 690
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    title = image_font(28, True)
    text = image_font(17)
    d.text((38, 24), "市民咨询高频问题（前10项）", fill=(45, 45, 45), font=title)
    left, top, right, bottom = 300, 90, 1250, 630
    max_val = TOP20[0][1]
    for i, (name, count) in enumerate(TOP20[:10]):
        y = top + i * 52
        d.text((38, y + 9), f"{i + 1}", fill=(70, 70, 70), font=text)
        d.text((78, y + 9), name, fill=(55, 55, 55), font=text)
        bar_w = int((right - left) * count / max_val)
        d.rounded_rectangle((left, y + 5, left + bar_w, y + 33), radius=6, fill=(74, 131, 230))
        d.text((left + bar_w + 14, y + 5), f"{count:,}", fill=(55, 55, 55), font=text)
    d.line((left, bottom, right, bottom), fill=(200, 210, 220), width=2)
    im.save(path)


def wrap_text(draw, text, font, max_width):
    lines = []
    for para in str(text).splitlines() or [""]:
        current = ""
        for char in para:
            trial = current + char
            if draw.textlength(trial, font=font) <= max_width:
                current = trial
            else:
                if current:
                    lines.append(current)
                current = char
        if current:
            lines.append(current)
    return lines or [""]


def load_rows():
    wb = load_workbook(CHAT_XLSX, read_only=True, data_only=True)
    ws = wb.active
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    grouped = defaultdict(list)
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, raw))
        grouped[str(row["会话ID"])].append(row)
    return grouped


def add_message(draw, y, row, user_font, bot_font, text_font, width):
    date_text = str(row["日期"]).replace(".000", "")
    msg_id = str(row["消息ID"])
    q = str(row["用户问句"] or "")
    a = str(row["答案"] or "")
    match = str(row["匹配项"] or "")
    draw.text((478, y), f"{date_text}  {msg_id[-6:]}", fill=(120, 125, 135), font=text_font)
    y += 23
    q_lines = wrap_text(draw, q, user_font, 520)
    q_h = max(42, 20 * len(q_lines) + 18)
    draw.rounded_rectangle((478, y, 478 + min(width, 550), y + q_h), radius=10, fill=(218, 235, 255))
    draw.text((500, y + 10), "用户", fill=(44, 92, 153), font=text_font)
    for idx, line in enumerate(q_lines):
        draw.text((500, y + 28 + idx * 20), line, fill=(40, 50, 65), font=user_font)
    y += q_h + 7
    a_lines = wrap_text(draw, a if a != "-" else "未形成有效回复（答案为 - ）", bot_font, 575)
    a_h = max(44, 20 * len(a_lines) + 18)
    draw.rounded_rectangle((770, y, 1375, y + a_h), radius=10, fill=(255, 255, 255), outline=(126, 166, 232), width=2)
    draw.text((790, y + 9), "杭州社保", fill=(52, 104, 185), font=text_font)
    for idx, line in enumerate(a_lines):
        draw.text((790, y + 27 + idx * 20), line, fill=(55, 55, 65), font=bot_font)
    draw.text((478, y + a_h + 5), f"匹配项：{match}", fill=(145, 150, 160), font=text_font)
    return y + a_h + 31


def make_case_screenshot(case, grouped):
    rows = grouped[case["id"]]
    user_font = image_font(15)
    bot_font = image_font(14)
    text_font = image_font(13)
    header_font = image_font(22, True)
    heights = []
    dummy = Image.new("RGB", (1440, 500), "white")
    dd = ImageDraw.Draw(dummy)
    for row in rows:
        q_lines = wrap_text(dd, str(row["用户问句"] or ""), user_font, 520)
        a = str(row["答案"] or "")
        a_lines = wrap_text(dd, a if a != "-" else "未形成有效回复（答案为 - ）", bot_font, 575)
        heights.append(23 + max(42, 20 * len(q_lines) + 18) + 7 + max(44, 20 * len(a_lines) + 18) + 31)
    h = max(650, 118 + sum(heights))
    im = Image.new("RGB", (1440, h), (246, 248, 252))
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, 410, h), fill=(229, 233, 240))
    d.rectangle((0, 0, 1440, 76), fill=(255, 255, 255))
    d.line((0, 76, 1440, 76), fill=(220, 225, 232), width=2)
    d.text((34, 20), "会话巡检", fill=(48, 52, 60), font=header_font)
    d.text((475, 18), "会话详情", fill=(48, 52, 60), font=header_font)
    d.text((1120, 21), case["id"], fill=(100, 105, 115), font=text_font)
    d.rounded_rectangle((25, 88, 385, 154), radius=8, fill=(214, 231, 255), outline=(126, 166, 232), width=2)
    d.text((48, 103), "杭州社保", fill=(45, 79, 126), font=image_font(17, True))
    d.text((48, 129), "文本会话记录", fill=(100, 110, 125), font=text_font)
    d.text((32, 178), "会话ID", fill=(100, 105, 115), font=text_font)
    d.text((32, 202), case["id"], fill=(45, 79, 126), font=image_font(14))
    d.line((430, 90, 430, h - 24), fill=(220, 225, 232), width=2)
    y = 90
    for row in rows:
        y = add_message(d, y, row, user_font, bot_font, text_font, 550)
    d.text((478, h - 26), "完整会话内容", fill=(150, 155, 165), font=text_font)
    out = RUNTIME / f"case_{'bad' if case in BAD_CASES else 'good'}_{case['id']}.png"
    im.save(out)
    return out


def add_case(doc, case, kind, image_path, new_page=False):
    fill = "E8A33B" if kind == "bad" else "53B495"
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = new_page
    r = p.add_run(case["title"])
    set_run_font(r, size=13, bold=True, color=(255, 255, 255))
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    p._p.get_or_add_pPr().append(shd)
    idp = add_text(doc, f"会话ID：{case['id']}", size=9.5, color=(100, 105, 115), after=3)
    idp.paragraph_format.keep_with_next = True
    ip = doc.add_paragraph()
    ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = ip.add_run()
    with Image.open(image_path) as im:
        width = Inches(6.35)
        run.add_picture(str(image_path), width=width)
    ip.paragraph_format.space_after = Pt(5)
    label = "问题表现：" if kind == "bad" else "优秀表现："
    add_rich_text(doc, label, case["finding"])
    if kind == "bad":
        add_rich_text(doc, "优化动作：", case["action"])


def add_high_frequency_table(doc):
    add_heading(doc, "市民咨询高频问题", level=2)
    add_text(doc, "本期市民咨询主要集中在人工客服、人社小灵光、问候语、网报新增回退申请表下载、业务办理、个人权益记录查询打印等事项，前20项情况如下。", after=7)
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(5.1)
    table.columns[1].width = Inches(1.25)
    set_table_borders(table)
    set_cell(table.cell(0, 0), "知识库名称", header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(table.cell(0, 1), "会话量", header=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for idx, (name, count) in enumerate(TOP20, 1):
        cells = table.add_row().cells
        set_cell(cells[0], name, fill="F2F9FF" if idx % 2 == 0 else "FFFFFF")
        set_cell(cells[1], f"{count:,}", align=WD_ALIGN_PARAGRAPH.CENTER, fill="F2F9FF" if idx % 2 == 0 else "FFFFFF")
    return table


def replace_template_media(docx_path, images):
    """Replace the first seven media slots when present, preserving relationships."""
    tmp = docx_path.with_suffix(".media.tmp.docx")
    with zipfile.ZipFile(docx_path, "r") as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        replacements = {f"word/media/image{i}.png": p.read_bytes() for i, p in enumerate(images, 1)}
        for info in src.infolist():
            dst.writestr(info, replacements.get(info.filename, src.read(info.filename)))
    tmp.replace(docx_path)


def build_report():
    grouped = load_rows()
    overview = RUNTIME / "fullmonth_overview.png"
    frequency = RUNTIME / "fullmonth_frequency.png"
    add_metric_image(overview)
    add_frequency_image(frequency)
    bad_images = [make_case_screenshot(case, grouped) for case in BAD_CASES]
    good_images = [make_case_screenshot(case, grouped) for case in GOOD_CASES]

    doc = Document(TEMPLATE)
    clear_document_body(doc)
    add_title(doc)
    add_text(
        doc,
        "以下是2026年8月1日-8月31日“杭州社保”智能AI接待运行的巡检报告。本期围绕咨询数据、会话明细、市民咨询高频问题、线上优化动作、待优化案例、优秀案例及后续优化重点进行梳理，重点呈现问题解决情况和可落地的优化方向。",
        size=12,
        after=10,
    )

    add_heading(doc, "一、咨询数据分析")
    add_heading(doc, "咨询情况", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(overview), width=Inches(6.55))
    p.paragraph_format.space_after = Pt(7)
    add_rich_text(doc, "总体运行情况。", f"2026年8月1日-8月31日，杭州社保智能机器人接待量为{RECEPTION_DISPLAY}，咨询量为{CONSULTATION_DISPLAY}，有效接待量为{EFFECTIVE_DISPLAY}，客户活跃度为{ACTIVITY_DISPLAY}。平均对话总轮次和有效平均对话轮次均为{AVG_ROUNDS}。")
    add_rich_text(doc, "复核情况。", f"本期人工复核{REVIEW_TOTAL:,}条文本聊天记录，其中有效回复{GOOD_TOTAL:,}条，发现待优化问题{BAD_TOTAL:,}条，涉及{SESSION_TOTAL:,}个会话，复核准确率为{ACCURACY:.1f}%。")
    add_rich_text(doc, "运行表现。", "咨询内容以人工客服转接、政策事项查询、表格及证明下载、参保停保办理、跨省转移接续等需求为主。对话整体保持较短轮次，但复杂政策、多轮追问和办理主体判断仍是影响问题闭环的主要环节。")

    add_heading(doc, "二、会话明细分析", new_page=True)
    add_heading(doc, "会话数分布", level=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(frequency), width=Inches(6.55))
    p.paragraph_format.space_after = Pt(7)
    add_rich_text(doc, "会话结构。", f"本期复核{REVIEW_TOTAL:,}条记录，形成有效回复{GOOD_TOTAL:,}条，待优化问题{BAD_TOTAL:,}条，复核准确率为{ACCURACY:.1f}%。待优化问题主要表现为未形成有效回复、重复追问未承接、办理主体判断不清及复杂问题未拆分回答。")
    add_rich_text(doc, "高频事项。", "人工客服和人社小灵光相关咨询量较高，业务办理、个人权益记录查询打印、参保停保登记、跨省转移接续等事项具有较强的办理指引需求。")
    add_high_frequency_table(doc)

    add_heading(doc, "三、线上优化动作", new_page=True)
    add_text(doc, "本期线上优化动作围绕未形成有效回复和问题未闭环场景展开，重点落实以下工作：")
    add_rich_text(doc, "1. 补齐高频复杂问法。", "围绕企业参保、职工停保、灵活就业参保停保、养老保险缴费年限查询、重复缴费、退休待遇领取地等事项，补充口语化问法、同义表达和组合条件。")
    add_rich_text(doc, "2. 完善主体判断和办理分支。", "首轮优先识别单位职工、灵活就业人员、企业经办人等身份，分别给出办理责任、线上入口、线下渠道、材料要求和时间提醒，减少答非所问。")
    add_rich_text(doc, "3. 加强多轮承接和多意图拆解。", "对用户连续追问的同一事项保留上下文状态，对同时涉及养老、医保、失业待遇等多个诉求的咨询逐项回应，避免只处理其中一个问题。")
    add_rich_text(doc, "4. 建立兜底和回归复核。", f"对答案为空、重复追问、办理路径缺失等问题设置明确兜底话术，并将本期发现的{BAD_TOTAL:,}条待优化问题纳入知识补齐、发布验证和后续复核。")

    add_heading(doc, "四、待优化案例与线上优化动作", new_page=True)
    add_text(doc, f"本期人工复核{REVIEW_TOTAL:,}条记录，发现待优化问题{BAD_TOTAL:,}条，涉及{SESSION_TOTAL:,}个会话；有效回复{GOOD_TOTAL:,}条，复核准确率为{ACCURACY:.1f}%。以下展示3例典型案例，全部问题数量已在本节开头汇总。")
    for index, (case, image) in enumerate(zip(BAD_CASES, bad_images)):
        add_case(doc, case, "bad", image, new_page=bool(index))

    add_heading(doc, "五、本期优秀案例", new_page=True)
    add_text(doc, "本期优秀案例体现出较好的场景识别、信息组织、边界说明和办理路径指引能力，可沉淀为“识别场景—说明结论—提供入口—提示下一步”的答复结构。")
    for index, (case, image) in enumerate(zip(GOOD_CASES, good_images)):
        add_case(doc, case, "good", image, new_page=bool(index))

    add_heading(doc, "六、后续优化重点", new_page=True)
    add_rich_text(doc, "1. 优先治理空回复和连续追问。", "持续补齐本期333条待优化问题对应的知识内容，重点关注企业参保后办理、职工停保责任、缴费年限查询和待遇领取地判断。")
    add_rich_text(doc, "2. 强化复杂事项的流程化回答。", "将办理主体、适用条件、线上入口、线下渠道、材料要求和下一步操作固化为结构化答复顺序，提升政策咨询的可执行性。")
    add_rich_text(doc, "3. 沉淀优秀答复模板。", "将灵活就业参保停保、多事项连续咨询和待遇领取地判断中的有效做法，转化为同类问题的通用话术和回归测试样例。")
    add_rich_text(doc, "4. 持续跟踪服务质量。", "按周期复核有效回复、待优化问题和准确率变化，对新增高频事项及时补充知识和兜底策略，形成问题发现、优化发布和效果复核的闭环。")

    # Use the template's existing header/footer and set the page background to white.
    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build_report())
