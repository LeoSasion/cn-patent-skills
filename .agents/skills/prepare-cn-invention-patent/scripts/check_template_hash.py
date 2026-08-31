from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree


BUNDLED_SHA256 = "d407d40f679acbe1cf02760699709c92dc085b38be55edf3f4f380e36ecbea91"
REQUIRED_TAGS = {
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
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
NS = {"w": W_NS, "r": REL_NS, "cp": CP_NS, "dc": DC_NS}
EXPECTED_PAGE = {"width_twips": 11906, "height_twips": 16838}
EXPECTED_MARGINS = {"top": 1134, "bottom": 1134, "left": 1247, "right": 1247}
FORBIDDEN_PREFIXES = ("word/embeddings/", "word/comments", "customXml/")
FORBIDDEN_NAMES = {
    "docProps/custom.xml",
    "word/vbaProject.bin",
    "word/people.xml",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def parse_xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name))


def inspect_template(path: Path) -> tuple[dict[str, object], list[str]]:
    errors: list[str] = []
    report: dict[str, object] = {"path": str(path), "sha256": sha256(path)}
    try:
        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                errors.append(f"DOCX 压缩成员损坏：{bad_member}")
            names = set(archive.namelist())
            for required in ("word/document.xml", "docProps/core.xml", "_rels/.rels"):
                if required not in names:
                    errors.append(f"DOCX 缺少必要部件：{required}")
            for name in sorted(names):
                if name in FORBIDDEN_NAMES or name.startswith(FORBIDDEN_PREFIXES):
                    errors.append(f"DOCX 含禁止部件：{name}")
            if errors:
                return report, errors

            document = parse_xml(archive, "word/document.xml")
            tags = document.xpath(".//w:sdt/w:sdtPr/w:tag/@w:val", namespaces=NS)
            tag_strings = [str(value) for value in tags]
            duplicates = sorted({tag for tag in tag_strings if tag_strings.count(tag) > 1})
            missing = sorted(REQUIRED_TAGS - set(tag_strings))
            unexpected = sorted(set(tag_strings) - REQUIRED_TAGS)
            report["content_control_tags"] = tag_strings
            if duplicates:
                errors.append(f"内容控件标签重复：{duplicates}")
            if missing or unexpected:
                errors.append(f"内容控件标签不一致：missing={missing} unexpected={unexpected}")

            sections = document.xpath(".//w:sectPr", namespaces=NS)
            if len(sections) != 1:
                errors.append(f"模板必须只有一个节，实际 {len(sections)}")
            else:
                page_sizes = sections[0].xpath("./w:pgSz", namespaces=NS)
                margins = sections[0].xpath("./w:pgMar", namespaces=NS)
                if len(page_sizes) != 1 or len(margins) != 1:
                    errors.append("模板缺少唯一的页面尺寸或页边距设置")
                else:
                    width = int(page_sizes[0].get(f"{{{W_NS}}}w", "0"))
                    height = int(page_sizes[0].get(f"{{{W_NS}}}h", "0"))
                    actual_margins = {
                        key: int(margins[0].get(f"{{{W_NS}}}{key}", "0"))
                        for key in EXPECTED_MARGINS
                    }
                    report["page_twips"] = {"width": width, "height": height}
                    report["margins_twips"] = actual_margins
                    if abs(width - EXPECTED_PAGE["width_twips"]) > 2 or abs(height - EXPECTED_PAGE["height_twips"]) > 2:
                        errors.append(f"页面不是预期 A4：{width} × {height} twips")
                    for key, expected in EXPECTED_MARGINS.items():
                        if abs(actual_margins[key] - expected) > 2:
                            errors.append(f"页边距不一致：{key}={actual_margins[key]} expected={expected}")

            external: list[str] = []
            for name in sorted(value for value in names if value.endswith(".rels")):
                relationships = parse_xml(archive, name)
                for relationship in relationships.xpath("./r:Relationship[@TargetMode='External']", namespaces=NS):
                    external.append(f"{name}:{relationship.get('Target', '')}")
            report["external_relationships"] = external
            if external:
                errors.append(f"DOCX 含外部关系：{external}")

            core = parse_xml(archive, "docProps/core.xml")
            creator = "".join(core.xpath("./dc:creator/text()", namespaces=NS)).strip()
            last_modified_by = "".join(core.xpath("./cp:lastModifiedBy/text()", namespaces=NS)).strip()
            report["core_metadata"] = {
                "creator": creator,
                "last_modified_by": last_modified_by,
            }
            if creator or last_modified_by:
                errors.append("模板含个人作者或最近修改者元数据")

            hidden = document.xpath(".//w:vanish", namespaces=NS)
            revisions = document.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo", namespaces=NS)
            report["hidden_nodes"] = len(hidden)
            report["revision_nodes"] = len(revisions)
            if hidden:
                errors.append(f"模板含隐藏文字节点：{len(hidden)}")
            if revisions:
                errors.append(f"模板含修订节点：{len(revisions)}")
    except (OSError, zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as exc:
        errors.append(f"DOCX 无法解析：{exc}")
    return report, errors


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="核验内置技术交底书模板的哈希、结构、页面和隐私边界。")
    parser.add_argument("template", nargs="?", help="待核验模板；省略时核验内置模板")
    parser.add_argument("--structure-only", action="store_true", help="只核验结构，不要求内置哈希")
    parser.add_argument("--json", action="store_true", help="输出 JSON 报告")
    args = parser.parse_args()

    bundled = Path(__file__).resolve().parents[1] / "assets" / "技术交底书_空白模板.docx"
    template = Path(args.template).resolve() if args.template else bundled
    if not template.is_file():
        parser.error(f"模板不存在：{template}")

    report, errors = inspect_template(template)
    actual = str(report.get("sha256", ""))
    if not args.structure_only and actual != BUNDLED_SHA256:
        errors.append(f"模板哈希不一致：expected={BUNDLED_SHA256} actual={actual}")
    report["expected_sha256"] = None if args.structure_only else BUNDLED_SHA256
    report["errors"] = errors
    report["passed"] = not errors
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    elif errors:
        for error in errors:
            print(f"ERROR {error}")
    else:
        print(f"OK {actual} {template}")
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
