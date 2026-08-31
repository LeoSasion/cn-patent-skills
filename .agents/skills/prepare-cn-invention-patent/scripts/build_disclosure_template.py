from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from docx_package import normalize_docx


CONTROL_SPECS = (
    ("PATENT_TITLE", "发明专利名称", "[填写准确、简洁的发明名称]", True),
    ("CONTACT_NAME", "技术联系人", "[姓名]", True),
    ("CONTACT_PHONE", "电话", "[电话]", True),
    ("CONTACT_FAX", "传真", "[传真，可空]", True),
    ("CONTACT_EMAIL", "电子邮箱", "[电子邮箱]", True),
    ("TERMS", "术语与缩写", "[逐项写明统一中文名称、英文缩写和定义；没有可留空]", False),
    ("SECTION_1_PROBLEM", "一、实际要解决的技术问题", "[说明现有技术在什么条件下发生什么技术失败；不要只写业务目标]", False),
    ("SECTION_2_BACKGROUND", "二、背景与最接近方案", "[说明现有流程、最接近方案及其局限；详细检索矩阵另附]", False),
    ("SECTION_3_DEFECTS_PURPOSE", "三、缺点与发明目的", "[逐项对应：现有缺点 → 本发明要达到的具体技术目的]", False),
    ("SECTION_4_SOLUTION", "四、详细技术方案", "[定义对象和前提，给出有序步骤、模块输入输出、状态、交互及异常分支，并逐图说明]", False),
    ("SECTION_5_KEY_POINTS", "五、关键点与拟保护组合", "[先写不可拆分的核心区别组合，再写可作为从属限定的细化特征]", False),
    ("SECTION_6_ADVANTAGES", "六、技术效果", "[按“技术手段—作用机制—技术效果”说明，避免营销口号和未经验证的数字]", False),
    ("SECTION_7_ALTERNATIVES", "七、可替代实施方式", "[列出保持核心关系不变的算法、存储、部署、步骤或容错替代]", False),
    ("SECTION_8_SUPPORTING_MATERIALS", "八、辅助资料与待补信息", "[列出代码版本、测试、附图、检索材料、公开披露和仍需确认的事项]", False),
)

QUESTION_GUIDANCE = {
    "SECTION_1_PROBLEM": "聚焦可由当前方案和证据支持的技术问题。",
    "SECTION_2_BACKGROUND": "正文只写代理人员理解方案所需的背景；检索结论单独留痕。",
    "SECTION_3_DEFECTS_PURPOSE": "每个目的都应能回到一个具体缺点，并由后续技术手段实现。",
    "SECTION_4_SOLUTION": "建议按总体链路、模块关系、正常路径、条件分支和异常恢复分层撰写。",
    "SECTION_5_KEY_POINTS": "不要把常规模块、营销功能或未经验证的效果单独当作保护点。",
    "SECTION_6_ADVANTAGES": "只写由已说明技术手段产生、且可被当前事实支持的效果。",
    "SECTION_7_ALTERNATIVES": "替代方案应保留核心技术关系，同时提供合理回退位置。",
    "SECTION_8_SUPPORTING_MATERIALS": "身份原件、证照、秘密和客户数据不放入本文件。",
}

ACCENT = "1F4E5F"
LIGHT_ACCENT = "EAF1F3"
LIGHT_GRAY = "F6F7F8"
BORDER = "AEB8BC"
TEXT_GRAY = RGBColor(89, 89, 89)


def set_font(run, east_asia: str, latin: str, size: float, bold: bool = False, color=None) -> None:
    run.font.name = latin
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.25

    for name, east_asia, size, color in (
        ("Title", "黑体", 22, RGBColor(31, 78, 95)),
        ("Subtitle", "宋体", 10.5, TEXT_GRAY),
        ("Heading 1", "黑体", 15, RGBColor(31, 78, 95)),
        ("Heading 2", "黑体", 12.5, RGBColor(31, 78, 95)),
    ):
        style = document.styles[name]
        style.font.name = "Arial"
        style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia)
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = name.startswith("Heading") or name == "Title"
        style.paragraph_format.keep_with_next = True
    document.styles["Heading 1"].paragraph_format.space_before = Pt(8)
    document.styles["Heading 1"].paragraph_format.space_after = Pt(6)
    document.styles["Heading 2"].paragraph_format.space_before = Pt(6)
    document.styles["Heading 2"].paragraph_format.space_after = Pt(4)

    styles = document.styles
    for name in ("FormLabel", "FormGuidance", "FormAnswer"):
        if name not in styles:
            styles.add_style(name, 1)
    label = styles["FormLabel"]
    label.font.name = "Arial"
    label._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "黑体")
    label.font.size = Pt(10)
    label.font.bold = True
    label.font.color.rgb = RGBColor(31, 78, 95)
    label.paragraph_format.space_after = Pt(2)
    label.paragraph_format.keep_with_next = True

    guidance = styles["FormGuidance"]
    guidance.font.name = "Arial"
    guidance._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    guidance.font.size = Pt(9.5)
    guidance.font.color.rgb = TEXT_GRAY
    guidance.paragraph_format.space_after = Pt(4)
    guidance.paragraph_format.keep_with_next = True

    answer = styles["FormAnswer"]
    answer.font.name = "Arial"
    answer._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    answer.font.size = Pt(10.5)
    answer.paragraph_format.space_after = Pt(4)
    answer.paragraph_format.line_spacing = 1.25


def shade_cell(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 100, start: int = 120, bottom: int = 100, end: int = 120) -> None:
    properties = cell._tc.get_or_add_tcPr()
    margins = properties.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.append(margins)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color: str = BORDER, size: str = "6") -> None:
    properties = table._tbl.tblPr
    borders = properties.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def wrap_paragraph_in_control(paragraph, tag: str, alias: str, placeholder: str, control_id: int, single_line: bool, blank_lines: int = 0) -> None:
    paragraph.style = "FormAnswer"
    paragraph.text = ""
    run = paragraph.add_run(placeholder)
    set_font(run, "宋体", "Arial", 10, color=TEXT_GRAY)

    p = paragraph._p
    parent = p.getparent()
    index = parent.index(p)
    parent.remove(p)

    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    alias_node = OxmlElement("w:alias")
    alias_node.set(qn("w:val"), alias)
    properties.append(alias_node)
    tag_node = OxmlElement("w:tag")
    tag_node.set(qn("w:val"), tag)
    properties.append(tag_node)
    id_node = OxmlElement("w:id")
    id_node.set(qn("w:val"), str(control_id))
    properties.append(id_node)
    if single_line:
        properties.append(OxmlElement("w:text"))
    sdt.append(properties)

    content = OxmlElement("w:sdtContent")
    content.append(p)
    for _ in range(blank_lines):
        extra = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        style = OxmlElement("w:pStyle")
        style.set(qn("w:val"), "FormAnswer")
        ppr.append(style)
        extra.append(ppr)
        content.append(extra)
    sdt.append(content)
    parent.insert(index, sdt)


def add_labeled_control(cell, label: str, tag: str, alias: str, placeholder: str, control_id: int) -> None:
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)
    label_paragraph = cell.paragraphs[0]
    label_paragraph.style = "FormLabel"
    label_run = label_paragraph.add_run(label)
    set_font(label_run, "黑体", "Arial", 9.5, bold=True, color=RGBColor(31, 78, 95))
    answer = cell.add_paragraph()
    wrap_paragraph_in_control(answer, tag, alias, placeholder, control_id, True)


def add_response_box(document: Document, tag: str, alias: str, placeholder: str, control_id: int, blank_lines: int) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table)
    cell = table.cell(0, 0)
    shade_cell(cell, LIGHT_GRAY)
    set_cell_margins(cell, top=140, start=160, bottom=140, end=160)
    paragraph = cell.paragraphs[0]
    wrap_paragraph_in_control(paragraph, tag, alias, placeholder, control_id, False, blank_lines)
    document.add_paragraph().paragraph_format.space_after = Pt(1)


def add_notice(document: Document) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="C9D8DD", size="4")
    cell = table.cell(0, 0)
    shade_cell(cell, LIGHT_ACCENT)
    set_cell_margins(cell, top=120, start=160, bottom=120, end=160)
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(
        "用途说明：本表用于内部技术交底和申请准备，不是正式申请文件。只填写可追溯事实；未确认身份与联系方式应留空。"
    )
    set_font(run, "宋体", "Arial", 9.5, color=RGBColor(31, 78, 95))


def add_footer(document: Document) -> None:
    footer = document.sections[0].footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    value = OxmlElement("w:t")
    value.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instruction, separate, value, end):
        field_run = OxmlElement("w:r")
        field_run.append(node)
        paragraph._p.append(field_run)


def build_document() -> Document:
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.9)
    configure_styles(document)

    properties = document.core_properties
    properties.title = "发明专利技术交底书空白模板"
    properties.subject = "中国发明专利申请准备的内部技术交底模板"
    properties.creator = ""
    properties.last_modified_by = ""
    properties.keywords = ""
    properties.comments = ""

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("发明专利技术交底书")
    subtitle = document.add_paragraph(style="Subtitle")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("内部技术材料 · 非正式申请文件")
    add_notice(document)

    document.add_heading("基本信息", level=1)
    table = document.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table)
    specs = {tag: (alias, placeholder) for tag, alias, placeholder, _ in CONTROL_SPECS}
    add_labeled_control(table.cell(0, 0), "发明专利名称", "PATENT_TITLE", *specs["PATENT_TITLE"], 1001)
    add_labeled_control(table.cell(0, 1), "技术联系人", "CONTACT_NAME", *specs["CONTACT_NAME"], 1002)
    add_labeled_control(table.cell(1, 0), "电话", "CONTACT_PHONE", *specs["CONTACT_PHONE"], 1003)
    add_labeled_control(table.cell(1, 1), "传真", "CONTACT_FAX", *specs["CONTACT_FAX"], 1004)
    add_labeled_control(table.cell(2, 0), "电子邮箱", "CONTACT_EMAIL", *specs["CONTACT_EMAIL"], 1005)
    table.cell(2, 1).text = ""
    shade_cell(table.cell(2, 1), "FFFFFF")

    document.add_heading("术语与缩写", level=2)
    guidance = document.add_paragraph(style="FormGuidance")
    guidance.add_run("一个对象只使用一个中文主称谓；英文缩写首次出现时给出中文全称。")
    terms = next(item for item in CONTROL_SPECS if item[0] == "TERMS")
    add_response_box(document, terms[0], terms[1], terms[2], 1006, 1)

    question_specs = [item for item in CONTROL_SPECS if item[0].startswith("SECTION_")]
    for index, (tag, alias, placeholder, _) in enumerate(question_specs, start=1):
        if index == 1 or index in (3, 5, 7):
            document.add_page_break()
        document.add_heading(alias, level=1)
        paragraph = document.add_paragraph(style="FormGuidance")
        paragraph.add_run(QUESTION_GUIDANCE[tag])
        blank_lines = 3 if tag == "SECTION_4_SOLUTION" else 2
        add_response_box(document, tag, alias, placeholder, 1100 + index, blank_lines)

    add_footer(document)
    return document


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="确定性生成 A4 发明专利技术交底书空白模板。")
    parser.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[1] / "assets" / "技术交底书_空白模板.docx"),
    )
    args = parser.parse_args()
    output = Path(args.output).resolve()
    if output.suffix.casefold() != ".docx":
        parser.error("输出文件必须使用 .docx 扩展名")

    document = build_document()
    with tempfile.TemporaryDirectory(prefix="patent-template-") as directory:
        raw = Path(directory) / "raw.docx"
        document.save(raw)
        normalize_docx(raw, output)
    print(f"WROTE {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
