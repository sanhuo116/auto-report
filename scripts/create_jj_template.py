from __future__ import annotations

import os
import re
import tempfile
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


SOURCE = Path("/Users/zhuyanzhu/Desktop/晋江公安百应安盾外呼报告（7.1-7.31）.docx")
OUTPUT = Path(
    "/Users/zhuyanzhu/Documents/ChatGPT/自动巡检平台/"
    "晋江公安百应安盾外呼报告（排版优化模板）.docx"
)

NAVY = "17365D"
BLUE = "D9EAF7"
PALE_BLUE = "EEF5FB"
PALE_YELLOW = "FFF4CC"
PALE_GREEN = "EAF5EA"
GRAY = "666666"
LIGHT_BORDER = "B7C9D6"
TEXT = "222222"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, **kwargs) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge not in kwargs:
            continue
        edge_data = kwargs.get(edge)
        tag = "w:{}".format(edge)
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        for key in ["val", "sz", "space", "color"]:
            if key in edge_data:
                element.set(qn("w:{}".format(key)), str(edge_data[key]))


def set_cell_margins(cell, top=80, start=100, bottom=80, end=100) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_row_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_table_widths(table, widths_cm: list[float]) -> None:
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths_cm):
            row.cells[idx].width = Cm(width)
            tc_pr = row.cells[idx]._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(width * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_run_font(run, size=10.5, bold=False, color=TEXT, name="Hiragino Sans GB") -> None:
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


def set_para(paragraph, text="", size=10.5, bold=False, color=TEXT,
             align=None, first_line=True, before=0, after=6, line=1.5,
             name="Hiragino Sans GB"):
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
    if first_line:
        fmt.first_line_indent = Cm(0.74)
    else:
        fmt.first_line_indent = Cm(0)
    return paragraph


def add_body(doc, text, first_line=True, after=6) -> None:
    p = doc.add_paragraph()
    set_para(p, text, first_line=first_line, after=after)


def add_heading(doc, text, level=1) -> None:
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
    )
    p.paragraph_format.keep_with_next = True
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "5")
    bottom.set(qn("w:color"), BLUE)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_caption(doc, text) -> None:
    p = doc.add_paragraph()
    set_para(
        p, text, size=9, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False, before=0, after=8, line=1.0
    )


def style_table(table, header_rows=1, title_rows=0, banded=True) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for ri, row in enumerate(table.rows):
        set_row_cant_split(row)
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            set_cell_border(
                cell,
                top={"val": "single", "sz": 6, "color": LIGHT_BORDER},
                bottom={"val": "single", "sz": 6, "color": LIGHT_BORDER},
                left={"val": "single", "sz": 6, "color": LIGHT_BORDER},
                right={"val": "single", "sz": 6, "color": LIGHT_BORDER},
            )
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.first_line_indent = Cm(0)
                for run in p.runs:
                    set_run_font(run, size=9.2, color=TEXT)
        if ri < title_rows:
            for cell in row.cells:
                set_cell_shading(cell, NAVY)
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, size=10, bold=True, color="FFFFFF")
        elif ri < title_rows + header_rows:
            for cell in row.cells:
                set_cell_shading(cell, BLUE)
                for p in cell.paragraphs:
                    for run in p.runs:
                        set_run_font(run, size=9.2, bold=True, color=NAVY)
        elif banded and ri % 2 == 0:
            for cell in row.cells:
                set_cell_shading(cell, PALE_BLUE)


def add_picture_from_zip(doc, archive: ZipFile, media_name: str, width_in: float) -> None:
    with tempfile.TemporaryDirectory(prefix="jj_template_img_") as td:
        path = Path(td) / Path(media_name).name
        path.write_bytes(archive.read(media_name))
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(str(path), width=Inches(width_in))


def add_placeholder_table(doc, headers, rows, widths_cm, header_fill=BLUE, title=None) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.rows[0].cells[0].text = headers[0]
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = header
    for row_data in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row_data):
            cells[idx].text = value
    set_table_widths(table, widths_cm)
    style_table(table, header_rows=1, title_rows=0, banded=True)
    for cell in table.rows[0].cells:
        set_cell_shading(cell, header_fill)
    if title:
        p = doc.add_paragraph()
        set_para(p, title, size=10.5, bold=True, color=NAVY,
                 first_line=False, before=4, after=4, line=1.15)
    return table


def add_header_footer(doc: Document, logo_path: Path) -> None:
    section = doc.sections[0]
    header = section.header
    hp = header.paragraphs[0]
    clear_paragraph(hp)
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(0)
    hp.add_run().add_picture(str(logo_path), width=Cm(3.0))
    footer = section.footer
    fp = footer.paragraphs[0]
    set_para(
        fp,
        "晋江公安百应安盾96110外呼报告｜排版优化模板",
        size=8.5,
        color=GRAY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False,
        before=0,
        after=0,
        line=1.0,
    )


def add_cover_title(doc: Document) -> None:
    p = doc.add_paragraph()
    set_para(
        p,
        "晋江公安百应安盾96110测试报告",
        size=20,
        bold=True,
        color=NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False,
        before=20,
        after=6,
        line=1.0,
    )
    p.paragraph_format.keep_with_next = True
    p = doc.add_paragraph()
    set_para(
        p,
        "晋江公安百应安盾96110外呼（{{统计周期}}）",
        size=12,
        color=GRAY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line=False,
        before=0,
        after=14,
        line=1.0,
    )
    p.paragraph_format.keep_with_next = True


def build() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.1)
    section.right_margin = Cm(2.1)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Hiragino Sans GB"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Hiragino Sans GB")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    with ZipFile(SOURCE) as archive:
        td = Path(tempfile.mkdtemp(prefix="jj_template_assets_"))
        logo = td / "image1.png"
        logo.write_bytes(archive.read("word/media/image1.png"))
        add_header_footer(doc, logo)

        add_cover_title(doc)
        add_body(
            doc,
            "{{统计周期}}整体运行{{运行情况结论}}，具体情况汇报如下：",
            first_line=False,
            after=10,
        )

        add_heading(doc, "一、AI外呼诈骗类型分析")
        add_body(
            doc,
            "本期外呼以{{主要诈骗类型}}为主，占比{{主要类型占比}}；其次为{{第二类型}}，占比{{第二类型占比}}。{{其他类型分析}}",
        )
        add_picture_from_zip(doc, archive, "word/media/image2.png", 5.95)
        add_caption(doc, "图1  AI外呼诈骗类型分布（示例图，请替换为当期图表）")

        add_heading(doc, "二、外呼情况分析")
        add_body(
            doc,
            "统计本期外呼情况，共外呼{{外呼总量}}人次，接通{{已接听}}人次，排除空号、关机、停机及拦截{{空关停拦截}}个，实际有效外呼{{实际外呼}}个市民，接通率为{{接通率}}。",
        )
        add_body(doc, "具体情况如下：", first_line=False, after=4)
        summary = doc.add_table(rows=3, cols=9)
        summary.cell(0, 0).merge(summary.cell(0, 8))
        summary.cell(0, 0).text = "晋江安盾96110外呼汇总表"
        headers = ["日期", "外呼总量", "空关停/拦截", "实际外呼", "已接听", "未接", "其他", "拒接", "无法接通"]
        values = ["{{统计周期}}", "{{外呼总量}}", "{{空关停拦截}}", "{{实际外呼}}", "{{已接听}}", "{{未接}}", "{{其他}}", "{{拒接}}", "{{无法接通}}"]
        for i, text in enumerate(headers):
            summary.cell(1, i).text = text
        for i, text in enumerate(values):
            summary.cell(2, i).text = text
        set_table_widths(summary, [1.8, 1.7, 1.9, 1.8, 1.8, 1.5, 1.5, 1.5, 1.7])
        style_table(summary, header_rows=1, title_rows=1, banded=False)
        add_picture_from_zip(doc, archive, "word/media/image3.png", 5.95)
        add_caption(doc, "图2  AI外呼结果分析（示例图，请替换为当期图表）")

        add_heading(doc, "三、通话时长分析")
        add_picture_from_zip(doc, archive, "word/media/image4.png", 5.75)
        add_caption(doc, "图3  AI外呼接通时长占比分布（示例图，请替换为当期图表）")
        add_body(
            doc,
            "本期平均通话时长为{{平均通话时长}}，其中{{重点时长区间}}通话占比最高，为{{重点时长占比}}。通常情况下，有效通话时长越长，代表AI与群众的沟通有效性越强，宣防影响力也在逐步加深。",
        )

        doc.add_page_break()
        add_heading(doc, "系统通话分析", level=2)
        add_body(
            doc,
            "在实际外呼过程中，针对成功采集到的完整通话数据进行分析，发现如下情况：",
            first_line=False,
        )
        profile = doc.add_table(rows=6, cols=5)
        profile.cell(0, 0).merge(profile.cell(0, 4))
        profile.cell(0, 0).text = "群众画像分析"
        profile_headers = ["线索类别", "类别说明", "总线索数量（个）", "确有实情数量（个）", "AI标记准确率（%）"]
        for i, text in enumerate(profile_headers):
            profile.cell(1, i).text = text
        sample_rows = [
            ["有效预警", "已被骗", "{{数量}}", "{{数量}}", "{{准确率}}"],
            ["有效预警", "未被骗", "{{数量}}", "{{数量}}", "{{准确率}}"],
            ["有效预警", "已被骗-未报警", "{{数量}}", "{{数量}}", "{{准确率}}"],
            ["有效预警", "已被骗-已报警", "{{数量}}", "{{数量}}", "{{准确率}}"],
        ]
        for ri, row_data in enumerate(sample_rows, start=2):
            for ci, value in enumerate(row_data):
                profile.cell(ri, ci).text = value
        set_table_widths(profile, [2.0, 3.0, 3.0, 3.0, 3.0])
        style_table(profile, header_rows=1, title_rows=1, banded=True)
        add_body(
            doc,
            "{{巡检结论}}通过对后台已采集语料进行批量巡检分析发现，主流程分支命中整体{{命中评价}}，但仍存在{{主要问题}}。",
        )
        add_body(
            doc,
            "典型原话示例：{{用户原话}}",
            first_line=False,
            after=2,
        )
        add_body(doc, "解决措施：", first_line=False, after=2)
        add_body(doc, "{{优化措施}}", first_line=True, after=2)
        add_body(doc, "当前状态：{{已完成优化/待验证/持续跟踪}}", first_line=False, after=8)

        doc.add_page_break()
        add_heading(doc, "待优化案例", level=2)
        add_body(
            doc,
            "本节用于沉淀巡检中发现的识别、分支命中、回复内容或流程衔接问题。建议每条案例保留用户原话、AI实际表现、问题定位和后续优化动作，便于复盘闭环。",
            first_line=False,
            after=8,
        )
        pending_rows = [
            [
                "01",
                "否定+肯定混合表达",
                "用户原话：没有没有，那我这里我有设置拦截的，有下载反诈APP。\n\n【此处插入通话截图或录音截图】",
                "AI可能只识别其中一类语意，出现误触或分支判断不完整。",
                "在否定分支关键词后补充否定语意规避；增加同类语料测试并复核命中结果。",
                "{{已优化/待验证}}",
            ],
            [
                "{{编号}}",
                "{{问题场景}}",
                "{{用户原话}}\n\n【此处插入截图】",
                "{{AI实际表现}}",
                "{{问题分析与优化方案}}",
                "{{状态}}",
            ],
            [
                "{{编号}}",
                "{{问题场景}}",
                "{{用户原话}}\n\n【此处插入截图】",
                "{{AI实际表现}}",
                "{{问题分析与优化方案}}",
                "{{状态}}",
            ],
        ]
        add_placeholder_table(
            doc,
            ["序号", "场景/问题", "用户原话及图片", "AI表现", "问题分析与优化方案", "状态"],
            pending_rows,
            [0.8, 1.8, 5.0, 3.0, 4.0, 1.6],
            header_fill=PALE_YELLOW,
        )
        add_caption(doc, "表3  待优化案例清单（可按实际案例增删行；图片直接嵌入对应单元格）")

        add_heading(doc, "优秀案例", level=2)
        add_body(
            doc,
            "本节用于沉淀识别准确、回复自然、宣防效果较好或具有复用价值的通话案例。优秀案例建议同时记录可复用话术和形成效果，作为后续机器人训练与质检标杆。",
            first_line=False,
            after=8,
        )
        excellent_rows = [
            [
                "01",
                "{{诈骗场景/用户状态}}",
                "{{用户关键原话}}\n\n【此处插入通话截图或录音截图】",
                "识别准确，分支衔接自然；{{亮点表现}}",
                "{{结果或价值}}",
                "{{可复用话术/配置}}",
            ],
            [
                "{{编号}}",
                "{{诈骗场景/用户状态}}",
                "{{用户关键原话}}\n\n【此处插入截图】",
                "{{AI处理亮点}}",
                "{{结果或价值}}",
                "{{可复用话术/配置}}",
            ],
            [
                "{{编号}}",
                "{{诈骗场景/用户状态}}",
                "{{用户关键原话}}\n\n【此处插入截图】",
                "{{AI处理亮点}}",
                "{{结果或价值}}",
                "{{可复用话术/配置}}",
            ],
        ]
        add_placeholder_table(
            doc,
            ["序号", "场景", "用户原话及图片", "AI处理亮点", "结果/价值", "可复用话术或配置"],
            excellent_rows,
            [0.8, 1.8, 5.0, 3.0, 2.5, 3.1],
            header_fill=PALE_GREEN,
        )
        add_caption(doc, "表4  优秀案例清单（可按实际案例增删行；图片直接嵌入对应单元格）")

        add_heading(doc, "五、后续的AI智能度训练优化")
        add_body(
            doc,
            "围绕巡检结果，持续优化升级现有bot外呼机器人，主要从以下方面展开：",
            first_line=False,
        )
        add_body(doc, "1、补充部分诈骗场景的特定分支关键词，提升分支命中准确度。")
        add_body(doc, "2、细化知识库，例如将“市民已被骗”进一步区分为“以前被骗”“近期被骗”等场景，匹配差异化回复话术。")
        add_body(doc, "3、精简冗长话术内容，并结合最新诈骗手法持续更新知识库与外呼策略。")

    doc.core_properties.title = "晋江公安百应安盾96110外呼报告（排版优化模板）"
    doc.core_properties.subject = "外呼报告模板"
    doc.core_properties.author = "百应"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
