from __future__ import annotations

import json
import math
import re
import tempfile
import textwrap
from collections import Counter
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from openpyxl import load_workbook

from create_jj_template import (
    BLUE,
    GRAY,
    LIGHT_BORDER,
    NAVY,
    PALE_BLUE,
    PALE_GREEN,
    PALE_YELLOW,
    TEXT,
    add_caption,
    add_header_footer,
    add_heading,
    add_body,
    set_cell_border,
    set_cell_margins,
    set_cell_shading,
    set_para,
    set_repeat_table_header,
    set_row_cant_split,
    set_run_font,
    set_table_widths,
    style_table,
)


INPUT = Path("/Users/zhuyanzhu/Downloads/导出AI通话记录_EX260818111534122374.xlsx")
TEMPLATE = Path(
    "/Users/zhuyanzhu/Documents/ChatGPT/自动巡检平台/"
    "晋江公安百应安盾外呼报告（排版优化模板）.docx"
)
OUTPUT = Path(
    "/Users/zhuyanzhu/Documents/ChatGPT/自动巡检平台/"
    "鄞州公安反诈AI外呼巡检报告（2026年7月18日-7月29日）.docx"
)

FONT_CANDIDATES = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
]


def font(size: int, bold: bool = False):
    for path in FONT_CANDIDATES:
        try:
            return ImageFont.truetype(path, size=size, index=0)
        except Exception:
            continue
    return ImageFont.load_default()


def parse_seconds(value) -> int:
    match = re.search(r"(\d+)", str(value or ""))
    return int(match.group(1)) if match else 0


def parse_tags(value) -> list[str]:
    if not value:
        return []
    try:
        result = json.loads(value)
        return result if isinstance(result, list) else []
    except Exception:
        return re.findall(r"[^[\]\",]+", str(value))


def dedupe_records(path: Path) -> tuple[list[dict], int]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    rows = list(worksheet.iter_rows(values_only=True))
    headers = [str(x) if x is not None else "" for x in rows[0]]
    records = []
    seen = set()
    duplicate_rows = 0
    for row in rows[1:]:
        record = dict(zip(headers, row))
        record_id = str(record.get("通话记录id") or "")
        if record_id in seen:
            duplicate_rows += 1
            continue
        seen.add(record_id)
        records.append(record)
    return records, duplicate_rows


def wrap_text(draw, text: str, text_font, max_width: int) -> list[str]:
    lines: list[str] = []
    for paragraph in str(text or "").splitlines() or [""]:
        current = ""
        for char in paragraph:
            candidate = current + char
            if draw.textbbox((0, 0), candidate, font=text_font)[2] <= max_width:
                current = candidate
            else:
                if current:
                    lines.append(current)
                current = char
        if current:
            lines.append(current)
        if not paragraph:
            lines.append("")
    return lines


def draw_donut(title: str, items: list[tuple[str, int]], output: Path, colors: list[str]) -> None:
    width, height = 1600, 900
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = font(34, True)
    body_font = font(25)
    small_font = font(21)
    draw.text((width // 2, 44), title, fill="#17365D", font=title_font, anchor="ma")

    total = sum(value for _, value in items) or 1
    cx, cy, radius = 430, 475, 260
    start = -90
    for idx, (_, value) in enumerate(items):
        angle = 360 * value / total
        draw.pieslice(
            (cx - radius, cy - radius, cx + radius, cy + radius),
            start=start,
            end=start + angle,
            fill=colors[idx % len(colors)],
            outline="white",
            width=3,
        )
        start += angle
    inner = 125
    draw.ellipse(
        (cx - inner, cy - inner, cx + inner, cy + inner),
        fill="white",
    )
    draw.text((cx, cy - 18), f"{total:,}", fill="#17365D", font=title_font, anchor="mm")
    draw.text((cx, cy + 32), "条记录", fill="#666666", font=small_font, anchor="mm")

    legend_x, legend_y = 820, 150
    for idx, (label, value) in enumerate(items):
        y = legend_y + idx * 76
        color = colors[idx % len(colors)]
        draw.rounded_rectangle((legend_x, y + 7, legend_x + 24, y + 31), radius=5, fill=color)
        percent = value / total * 100
        label = textwrap.shorten(label, width=22, placeholder="…")
        draw.text((legend_x + 42, y), label, fill="#222222", font=body_font)
        draw.text((1420, y), f"{value:,}  ({percent:.1f}%)", fill="#666666", font=small_font, anchor="ra")

    image.save(output)


def draw_case_card(
    title: str,
    script: str,
    record_id: str,
    duration: str,
    tags: list[str],
    transcript: str,
    output: Path,
    accent: str,
) -> None:
    width, height = 1200, 820
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = font(30, True)
    meta_font = font(20)
    body_font = font(22)
    draw.rectangle((0, 0, width, 84), fill=accent)
    draw.text((35, 24), title, fill="white", font=title_font)
    meta = f"话术：{script}    通话ID：{record_id}    时长：{duration}"
    draw.text((35, 112), meta, fill="#17365D", font=meta_font)
    if tags:
        draw.text((35, 148), "标签：" + "、".join(tags[:6]), fill="#666666", font=meta_font)
    body = transcript[:1700]
    y = 205
    for line in wrap_text(draw, body, body_font, width - 70)[:26]:
        draw.text((35, y), line, fill="#222222", font=body_font)
        y += 30
        if y > height - 55:
            break
    draw.line((35, height - 34, width - 35, height - 34), fill="#D9EAF7", width=2)
    draw.text((35, height - 25), "依据：本期AI外呼通话记录；同一通话记录ID按一条通话计入", fill="#777777", font=meta_font)
    image.save(output)


def add_image_to_cell(cell, image_path: Path, width_cm: float = 4.5) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_picture(str(image_path), width=Cm(width_cm))


def add_case_table(
    doc: Document,
    headers: list[str],
    cases: list[dict],
    widths: list[float],
    fill: str,
    caption: str,
) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    for idx, header in enumerate(headers):
        table.cell(0, idx).text = header
    for case in cases:
        cells = table.add_row().cells
        cells[0].text = case["no"]
        cells[1].text = case["scene"]
        add_image_to_cell(cells[2], case["image"], width_cm=4.5)
        cells[3].text = case["finding"]
        cells[4].text = case["action"]
    set_table_widths(table, widths)
    style_table(table, header_rows=1, title_rows=0, banded=True)
    for cell in table.rows[0].cells:
        set_cell_shading(cell, fill)
    for row in table.rows:
        set_row_cant_split(row)
    add_caption(doc, caption)


def call_type_label(name: str) -> str:
    return (
        name.replace("三季度-鄞州劝阻-", "")
        .replace("鄞州劝阻-", "")
        .replace("V1.0", "")
        .replace("结构A1", "")
        .strip(" -")
    )


def main() -> None:
    records, duplicate_rows = dedupe_records(INPUT)
    total = len(records)
    status_counts = Counter(str(d.get("通话状态") or "") for d in records)
    script_counts = Counter(str(d.get("话术名称") or "") for d in records)
    answered = [d for d in records if d.get("通话状态") == "已接听"]
    answered_seconds = [parse_seconds(d.get("通话时长")) for d in answered]
    transcript_records = [d for d in records if d.get("通话记录")]
    labeled_records = [d for d in records if d.get("通话标签")]

    starts = [datetime.fromisoformat(str(d["通话开始时间"])) for d in records if d.get("通话开始时间")]
    ends = [datetime.fromisoformat(str(d["通话结束时间"])) for d in records if d.get("通话结束时间")]
    start_time, end_time = min(starts), max(ends)
    period = f"{start_time.year}年{start_time.month}月{start_time.day}日-{end_time.month}月{end_time.day}日"
    period_short = f"{start_time.month}.{start_time.day}-{end_time.month}.{end_time.day}"

    unavailable_statuses = {"停机", "关机", "空号", "天盾拦截"}
    unavailable = sum(status_counts[s] for s in unavailable_statuses)
    effective = total - unavailable
    full_connect_rate = status_counts["已接听"] / total if total else 0
    effective_connect_rate = status_counts["已接听"] / effective if effective else 0
    avg_seconds = sum(answered_seconds) / len(answered_seconds) if answered_seconds else 0
    median_seconds = sorted(answered_seconds)[len(answered_seconds) // 2] if answered_seconds else 0
    over_60 = sum(x >= 60 for x in answered_seconds)

    tag_counts = Counter()
    for record in records:
        for tag in parse_tags(record.get("通话标签")):
            tag_counts[tag] += 1

    with tempfile.TemporaryDirectory(prefix="jj_call_report_assets_") as asset_dir:
        assets = Path(asset_dir)
        palette = ["#3F72D8", "#40B7D7", "#F2B84B", "#E26D5A", "#61B15A", "#8E6CCB", "#7B8794", "#E68A3F"]
        type_items = [(call_type_label(name), count) for name, count in script_counts.most_common()]
        status_order = ["已接听", "未接", "拒接", "无法接通", "占线", "停机", "关机", "空号", "天盾拦截"]
        status_items = [(name, status_counts.get(name, 0)) for name in status_order if status_counts.get(name, 0)]
        duration_buckets = [
            ("0-10秒", sum(x <= 10 for x in answered_seconds)),
            ("11-30秒", sum(11 <= x <= 30 for x in answered_seconds)),
            ("31-60秒", sum(31 <= x <= 60 for x in answered_seconds)),
            ("61-120秒", sum(61 <= x <= 120 for x in answered_seconds)),
            ("121秒以上", sum(x >= 121 for x in answered_seconds)),
        ]
        chart_type = assets / "chart_type.png"
        chart_status = assets / "chart_status.png"
        chart_duration = assets / "chart_duration.png"
        draw_donut("AI外呼话术类型分布", type_items, chart_type, palette)
        draw_donut("AI外呼通话状态分布", status_items, chart_status, palette)
        draw_donut("已接听通话时长分布", duration_buckets, chart_duration, palette)

        by_id = {str(d["通话记录id"]): d for d in records}
        pending_candidates = {
            "condition": max(
                records,
                key=lambda d: (str(d.get("通话记录") or "").count("条件判断-条件判断节点"), len(str(d.get("通话记录") or ""))),
            ),
            "misunderstanding": max(
                [d for d in records if "你是觉得说我" in str(d.get("通话记录") or "")] or records,
                key=lambda d: len(str(d.get("通话记录") or "")),
            ),
            "assistant": max(
                [
                    d for d in records
                    if "您好。在吗" in str(d.get("通话记录") or "")
                    and any(x in str(d.get("通话记录") or "") for x in ["小助理", "没听懂", "信号不好"])
                ] or records,
                key=lambda d: str(d.get("通话记录") or "").count("您好。在吗"),
            ),
        }
        good_candidates = {
            "刷单返利": next(
                d for d in records
                if d.get("通话记录id") == "2877154264410"
            ),
            "陌生群聊": next(
                d for d in records
                if d.get("通话记录id") == "2877214726910"
            ),
            "投资理财": next(
                d for d in records
                if d.get("通话记录id") == "2876943708110"
            ),
        }

        pending_cases = [
            {
                "no": "01",
                "scene": "条件节点文本泄漏\n刷单返利/贷款关联场景",
                "record": pending_candidates["condition"],
                "finding": "通话中多次直接输出“条件判断-条件判断节点”，且“转钱”后紧接“没有啊”时仍继续触发多个条件分支，用户信息没有被稳定归因。",
                "action": "清理流程节点调试文本；对“转钱/没转钱”等相邻正负语意增加上下文判断；以真实语料回归测试分支跳转。",
                "accent": "#C88A00",
            },
            {
                "no": "02",
                "scene": "用户明确否认后仍持续追问\n通用劝阻场景",
                "record": pending_candidates["misunderstanding"],
                "finding": "用户明确说明未刷单、未下载、已挂断贷款电话，并进一步询问“到底是提醒还是在取证”；AI仍重复泛化风险话术，没有先澄清来电目的。",
                "action": "增加“用户明确否认/质疑来意”分支：先解释预警来源与来电目的，再给出简短提醒；用户无风险时及时收束。",
                "accent": "#C88A00",
            },
            {
                "no": "03",
                "scene": "智能助理/信号异常未及时退出\n刷单返利场景",
                "record": pending_candidates["assistant"],
                "finding": "对话出现“机主正在忙”“小助理”“我没听懂”“信号不好”等明显非真人或通信异常信号，AI仍多次重复问题和“您好。在吗”，造成角色混淆。",
                "action": "增加智能助理、自动应答、连续听不清和信号异常的识别阈值；连续命中后礼貌结束，避免继续消耗通话时长。",
                "accent": "#C88A00",
            },
        ]
        excellent_cases = [
            {
                "no": "01",
                "scene": "刷单返利\n用户主动删除拉黑",
                "record": good_candidates["刷单返利"],
                "finding": "用户明确表示发现异常后已删除并拉黑对方；AI围绕“不要转账、不要缴费、不要提供验证码”等关键动作进行提醒，用户多次确认“清楚了、谢谢”。",
                "action": "保留“确认处置结果→补充高发套路→核实是否转账→给出96110/派出所出口”的简洁链路，作为刷单返利场景标杆话术。",
                "accent": "#4E9F62",
            },
            {
                "no": "02",
                "scene": "陌生群聊\n用户主动退群并删除",
                "record": good_candidates["陌生群聊"],
                "finding": "用户说明被陌生人拉群后感到不靠谱，已退出群聊并删除好友；AI准确承接用户行为，补充“不转账、不缴费、不提供验证码”等风险点，形成闭环。",
                "action": "复用“认可用户自主识诈行为+补充下一步动作+确认无资金损失”的回应结构，适合有明确处置动作的预警对象。",
                "accent": "#4E9F62",
            },
            {
                "no": "03",
                "scene": "投资理财\n用户明确未参与",
                "record": good_candidates["投资理财"],
                "finding": "用户连续否认陌生投资、扫码下载和转账行为，AI没有转入无关分支，针对虚假投资的高收益诱导和代操作风险做了集中提醒，用户明确表示“知道了”。",
                "action": "保留“分层核实关键动作→给出典型风险→确认用户理解→及时结束”的节奏，避免对低风险用户过度延长通话。",
                "accent": "#4E9F62",
            },
        ]

        for case in pending_cases + excellent_cases:
            record = case["record"]
            card_path = assets / f"case_{case['no']}_{len(case['scene'])}.png"
            draw_case_card(
                "待优化案例" if case in pending_cases else "优秀案例",
                call_type_label(str(record.get("话术名称") or "")),
                str(record.get("通话记录id") or ""),
                str(record.get("通话时长") or ""),
                parse_tags(record.get("通话标签")),
                str(record.get("通话记录") or ""),
                card_path,
                case["accent"],
            )
            case["image"] = card_path

        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.1)
        section.right_margin = Cm(2.1)
        section.header_distance = Cm(0.8)
        section.footer_distance = Cm(0.8)
        normal = doc.styles["Normal"]
        normal.font.name = "Hiragino Sans GB"
        normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Hiragino Sans GB")
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = RGBColor.from_string(TEXT)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_after = Pt(6)

        with ZipFile(TEMPLATE) as template_archive:
            logo_path = assets / "logo.png"
            logo_path.write_bytes(template_archive.read("word/media/image1.png"))
        add_header_footer(doc, logo_path)
        set_para(
            doc.sections[0].footer.paragraphs[0],
            "鄞州公安分局反诈AI外呼巡检报告",
            size=8.5,
            color=GRAY,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            first_line=False,
            before=0,
            after=0,
            line=1.0,
        )

        p = doc.add_paragraph()
        set_para(p, "鄞州公安分局反诈AI外呼巡检报告", size=20, bold=True,
                 color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                 before=20, after=6, line=1.0)
        p = doc.add_paragraph()
        set_para(p, f"外呼周期：{period}", size=12, color=GRAY,
                 align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False,
                 before=0, after=14, line=1.0)
        add_body(
            doc,
            f"{period}鄞州公安分局反诈AI外呼共纳入{total:,}条独立通话记录；同一“通话记录id”的重复项按一条通话计入。本报告所有统计、图表与案例均基于这{total:,}条独立通话记录。",
            first_line=False,
            after=10,
        )

        add_heading(doc, "一、AI外呼诈骗类型分析")
        top_scripts = script_counts.most_common(4)
        top_text = "、".join(f"{call_type_label(k)}{v:,}次（{v / total * 100:.1f}%）" for k, v in top_scripts)
        add_body(
            doc,
            f"本期共配置{len(script_counts)}类外呼话术，累计外呼{total:,}人次。其中{top_text}；其余话术合计{total - sum(v for _, v in top_scripts):,}人次。外呼任务以通用劝阻、刷单返利和投资理财三类为主。",
        )
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(chart_type), width=Inches(5.95))
        add_caption(doc, f"图1  AI外呼话术类型分布（本期{total:,}条独立通话记录）")

        type_table = doc.add_table(rows=1, cols=4)
        for idx, header in enumerate(["话术类型", "外呼量（次）", "占比（%）", "已接听（次）"]):
            type_table.cell(0, idx).text = header
        for name, count in script_counts.most_common():
            row = type_table.add_row().cells
            row[0].text = call_type_label(name)
            row[1].text = f"{count:,}"
            row[2].text = f"{count / total * 100:.1f}%"
            row[3].text = f"{sum(1 for d in records if d.get('话术名称') == name and d.get('通话状态') == '已接听'):,}"
        set_table_widths(type_table, [6.2, 3.0, 3.0, 3.8])
        style_table(type_table, header_rows=1, title_rows=0, banded=True)
        add_caption(doc, "表1  AI外呼话术类型统计")

        add_heading(doc, "二、外呼情况分析")
        add_body(
            doc,
            f"本期共外呼{total:,}人次，已接听{status_counts['已接听']:,}人次，未接{status_counts['未接']:,}人次，拒接{status_counts['拒接']:,}人次，无法接通{status_counts['无法接通']:,}人次，占线{status_counts['占线']:,}人次。剔除停机、关机、空号及天盾拦截共{unavailable:,}条后，实际有效外呼{effective:,}条；按有效外呼口径计算接通率为{effective_connect_rate * 100:.1f}%，按全量外呼口径计算为{full_connect_rate * 100:.1f}%。",
        )
        add_body(doc, "具体情况如下：", first_line=False, after=4)
        summary = doc.add_table(rows=3, cols=9)
        summary.cell(0, 0).merge(summary.cell(0, 8))
        summary.cell(0, 0).text = "鄞州公安反诈AI外呼汇总表"
        headers = ["日期", "外呼总量", "空关停/拦截", "实际外呼", "已接听", "未接", "其他", "拒接", "无法接通"]
        values = [period_short, f"{total:,}", f"{unavailable:,}", f"{effective:,}", f"{status_counts['已接听']:,}", f"{status_counts['未接']:,}", f"{status_counts['占线']:,}", f"{status_counts['拒接']:,}", f"{status_counts['无法接通']:,}"]
        for idx, value in enumerate(headers):
            summary.cell(1, idx).text = value
        for idx, value in enumerate(values):
            summary.cell(2, idx).text = value
        set_table_widths(summary, [1.8, 1.7, 1.9, 1.8, 1.8, 1.5, 1.5, 1.5, 1.7])
        style_table(summary, header_rows=1, title_rows=1, banded=False)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(chart_status), width=Inches(5.95))
        add_caption(doc, "图2  AI外呼通话状态分布")

        add_heading(doc, "三、通话时长分析")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(chart_duration), width=Inches(5.75))
        add_caption(doc, "图3  已接听通话时长分布")
        add_body(
            doc,
            f"本期已接听通话平均时长为{avg_seconds:.1f}秒，中位数为{median_seconds:.0f}秒，其中60秒以上通话{over_60:,}条，占已接听通话的{over_60 / len(answered) * 100:.1f}%。61-120秒区间占比最高，为{duration_buckets[3][1] / len(answered) * 100:.1f}%；数据表明接通后能够完成完整宣防的通话占比较高，短时挂断和无有效对话记录是下一阶段的明确优化重点。",
        )

        doc.add_page_break()
        add_heading(doc, "系统通话分析", level=2)
        add_body(
            doc,
            f"本期独立通话记录中，已接听{len(answered):,}条，其中有通话文本{len(transcript_records):,}条，占已接听通话的{len(transcript_records) / len(answered) * 100:.1f}%；通话标签非空{len(labeled_records):,}条。标签统计结果为：宣传成功{tag_counts['宣传成功']:,}条、非诈骗{tag_counts['非诈骗']:,}条、预警准确{tag_counts['预警准确']:,}条、疑似存在风险{tag_counts['疑似存在风险']:,}条、已被骗{tag_counts['已被骗']:,}条。本报告按标签名称进行数量统计，不将标签次数表述为人工准确率。",
        )
        profile = doc.add_table(rows=1, cols=5)
        profile.cell(0, 0).merge(profile.cell(0, 4))
        profile.cell(0, 0).text = "群众画像与系统标签分析"
        profile_headers = ["标签类别", "标签说明", "出现次数（次）", "占总通话量（%）", "分析结论"]
        header_cells = profile.add_row().cells
        for idx, value in enumerate(profile_headers):
            header_cells[idx].text = value
        profile_rows = [
            ["结果标签", "宣传成功", f"{tag_counts['宣传成功']:,}", f"{tag_counts['宣传成功'] / total * 100:.1f}%", "已完成有效宣防"],
            ["画像标签", "非诈骗", f"{tag_counts['非诈骗']:,}", f"{tag_counts['非诈骗'] / total * 100:.1f}%", "用户明确否认相关诈骗"],
            ["预警标签", "预警准确", f"{tag_counts['预警准确']:,}", f"{tag_counts['预警准确'] / total * 100:.1f}%", "系统标记为预警准确"],
            ["风险标签", "疑似存在风险", f"{tag_counts['疑似存在风险']:,}", f"{tag_counts['疑似存在风险'] / total * 100:.1f}%", "纳入重点风险案例"],
            ["风险标签", "已被骗", f"{tag_counts['已被骗']:,}", f"{tag_counts['已被骗'] / total * 100:.1f}%", "纳入重点处置案例"],
        ]
        for row_values in profile_rows:
            cells = profile.add_row().cells
            for idx, value in enumerate(row_values):
                cells[idx].text = value
        set_table_widths(profile, [2.1, 2.5, 2.4, 2.6, 5.0])
        style_table(profile, header_rows=1, title_rows=1, banded=True)
        add_caption(doc, "表2  通话标签与群众画像分析")
        add_body(
            doc,
            f"本次巡检明确发现，{sum('条件判断-条件判断节点' in str(d.get('通话记录') or '') for d in records):,}条通话文本中出现“条件判断-条件判断节点”内部节点文本，{sum('您好。在吗' in str(d.get('通话记录') or '') for d in records):,}条出现“您好。在吗”重复唤醒。流程调试文本、连续无效应答和异常对话退出机制是本期明确的三项优化重点。",
        )

        doc.add_page_break()
        add_heading(doc, "待优化案例", level=2)
        add_body(
            doc,
            "以下列出本期通话中已明确识别的三类典型问题：流程节点泄漏、用户明确否认后的重复追问、智能助理及通信异常未及时退出。",
            first_line=False,
            after=8,
        )
        add_case_table(
            doc,
            ["序号", "场景", "通话内容卡片", "发现与影响", "优化动作"],
            pending_cases,
            [0.8, 2.0, 5.0, 4.0, 4.4],
            PALE_YELLOW,
            "表3  待优化案例清单",
        )

        add_heading(doc, "优秀案例", level=2)
        add_body(
            doc,
            "以下列出本期通话中已形成完整宣传闭环的三条优秀案例，均具备明确的用户处置动作、匹配的AI回应和清晰的风险提醒结果。",
            first_line=False,
            after=8,
        )
        add_case_table(
            doc,
            ["序号", "场景", "通话内容卡片", "优秀表现", "标准化动作"],
            excellent_cases,
            [0.8, 2.0, 5.0, 4.0, 4.4],
            PALE_GREEN,
            "表4  优秀案例清单",
        )

        add_heading(doc, "五、后续的AI智能度训练优化")
        add_body(doc, "围绕本次巡检结果，本期确定以下AI外呼优化动作：", first_line=False)
        add_body(doc, "1、清理并屏蔽“条件判断-条件判断节点”等内部调试文本，建立流程节点泄漏的自动质检规则。")
        add_body(doc, "2、完善用户明确否认、质疑来意、已自主识诈、已删除拉黑等高频分支，减少无效重复追问。")
        add_body(doc, "3、增加智能助理、语音信箱、连续听不清、信号异常和用户挂断意图识别，达到阈值后及时结束。")
        add_body(doc, "4、建立“宣传成功、预警准确、疑似存在风险、已被骗”等标签与处置结果的关联统计，形成准确率和闭环处置指标。")

        doc.core_properties.title = f"鄞州公安反诈AI外呼巡检报告（{period}）"
        doc.core_properties.subject = "AI外呼巡检报告"
        doc.core_properties.author = "百应"
        doc.save(OUTPUT)
        print(OUTPUT)


if __name__ == "__main__":
    main()
