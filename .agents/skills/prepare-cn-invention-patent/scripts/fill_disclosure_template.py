from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from pathlib import Path

from lxml import etree

from docx_package import load_parts, sanitize_parts, write_parts


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}
REQUIRED_TAGS = (
    "PATENT_TITLE",
    "CONTACT_NAME",
    "CONTACT_PHONE",
    "CONTACT_FAX",
    "CONTACT_EMAIL",
    "TERMS",
    "SECTION_1_PROBLEM",
    "SECTION_2_BACKGROUND",
    "SECTION_3_DEFECTS_PURPOSE",
    "SECTION_4_SOLUTION",
    "SECTION_5_KEY_POINTS",
    "SECTION_6_ADVANTAGES",
    "SECTION_7_ALTERNATIVES",
    "SECTION_8_SUPPORTING_MATERIALS",
)
SINGLE_LINE_TAGS = {
    "PATENT_TITLE",
    "CONTACT_NAME",
    "CONTACT_PHONE",
    "CONTACT_FAX",
    "CONTACT_EMAIL",
}


def controls_by_tag(root: etree._Element) -> dict[str, etree._Element]:
    controls: dict[str, etree._Element] = {}
    duplicates: list[str] = []
    for control in root.xpath(".//w:sdt", namespaces=NS):
        tags = control.xpath("./w:sdtPr/w:tag/@w:val", namespaces=NS)
        if not tags:
            continue
        tag = str(tags[0])
        if tag in controls:
            duplicates.append(tag)
        controls[tag] = control
    if duplicates:
        raise ValueError(f"模板含重复内容控件标签：{', '.join(sorted(set(duplicates)))}")
    return controls


def paragraph_with_text(template: etree._Element | None, text: str) -> etree._Element:
    paragraph = etree.Element(f"{{{W_NS}}}p")
    if template is not None:
        ppr = template.find(f"{{{W_NS}}}pPr")
        if ppr is not None:
            paragraph.append(deepcopy(ppr))
    if text:
        run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
        node = etree.SubElement(run, f"{{{W_NS}}}t")
        if text[:1].isspace() or text[-1:].isspace():
            node.set(f"{{{XML_NS}}}space", "preserve")
        node.text = text
    return paragraph


def replace_control(control: etree._Element, value: str) -> None:
    content_nodes = control.xpath("./w:sdtContent", namespaces=NS)
    if len(content_nodes) != 1:
        raise ValueError("内容控件缺少唯一的 w:sdtContent")
    content = content_nodes[0]
    template_paragraph = content.find(f"{{{W_NS}}}p")
    for child in list(content):
        content.remove(child)
    lines = value.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines:
        lines = [""]
    for line in lines:
        content.append(paragraph_with_text(template_paragraph, line))
    for placeholder in control.xpath("./w:sdtPr/w:showingPlcHdr", namespaces=NS):
        placeholder.getparent().remove(placeholder)


def normalize_value(tag: str, value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        text = str(value)
    else:
        raise ValueError(f"字段值必须是字符串、数字、布尔值或 null：{tag}")
    if tag in SINGLE_LINE_TAGS and ("\n" in text or "\r" in text):
        raise ValueError(f"单行字段不能包含换行：{tag}")
    return text


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="按稳定内容控件标签填写内置技术交底书模板。")
    parser.add_argument(
        "--template",
        default=str(Path(__file__).resolve().parents[1] / "assets" / "技术交底书_空白模板.docx"),
    )
    parser.add_argument("--values", help="UTF-8 JSON 对象")
    parser.add_argument("--output", help="输出 DOCX；不得与模板相同")
    parser.add_argument("--require-all", action="store_true", help="要求 JSON 包含全部标签")
    parser.add_argument("--list-tags", action="store_true", help="只列出模板中的标签")
    args = parser.parse_args()

    template = Path(args.template).resolve()
    if not template.is_file() or template.suffix.casefold() != ".docx":
        parser.error(f"模板不存在或不是 DOCX：{template}")
    parts = load_parts(template)
    if "word/document.xml" not in parts:
        parser.error("模板缺少 word/document.xml")
    root = etree.fromstring(parts["word/document.xml"])
    controls = controls_by_tag(root)
    actual_tags = tuple(controls)
    missing_template_tags = sorted(set(REQUIRED_TAGS) - set(actual_tags))
    unexpected_template_tags = sorted(set(actual_tags) - set(REQUIRED_TAGS))
    if missing_template_tags or unexpected_template_tags:
        parser.error(
            "模板标签合同不一致："
            f"missing={missing_template_tags} unexpected={unexpected_template_tags}"
        )

    if args.list_tags:
        print(json.dumps(list(actual_tags), ensure_ascii=False, indent=2))
        return 0
    if not args.values or not args.output:
        parser.error("填写模板时必须同时提供 --values 和 --output")

    values_path = Path(args.values).resolve()
    output = Path(args.output).resolve()
    if output.suffix.casefold() != ".docx":
        parser.error("输出文件必须使用 .docx 扩展名")
    if output == template:
        parser.error("输出文件不得覆盖模板本身")
    try:
        values = json.loads(values_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        parser.error(f"无法读取字段 JSON：{exc}")
    if not isinstance(values, dict):
        parser.error("字段 JSON 顶层必须是对象")
    unknown = sorted(set(values) - set(REQUIRED_TAGS))
    if unknown:
        parser.error(f"字段 JSON 含未知标签：{unknown}")
    missing_values = sorted(set(REQUIRED_TAGS) - set(values))
    if args.require_all and missing_values:
        parser.error(f"字段 JSON 缺少标签：{missing_values}")

    try:
        for tag, raw_value in values.items():
            replace_control(controls[tag], normalize_value(tag, raw_value))
    except ValueError as exc:
        parser.error(str(exc))
    parts["word/document.xml"] = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    write_parts(sanitize_parts(parts), output)
    print(f"WROTE {output}")
    if missing_values:
        print(f"UNFILLED {', '.join(missing_values)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
