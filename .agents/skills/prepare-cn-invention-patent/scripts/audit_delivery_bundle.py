from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree as ET

from make_output_paths import DRAWING_PACKAGE_DOCS, PACKAGE_DOCS, REQUIRED_PACKAGE_DOCS


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}

A4_WIDTH = 595.28
A4_HEIGHT = 841.89
REVIEW_LABEL = "申请前校对稿（非电子提交格式）"
TEXT_EXTENSIONS = {".md", ".txt"}
RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp"}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def compact_text(value: str) -> str:
    value = re.sub(r"(?m)^\s{0,3}#{1,6}\s*", "", value)
    value = re.sub(r"[`*_>\s]+", "", value)
    return value.replace("．", ".")


def manifest_path(value: object, base: Path) -> Path | None:
    text = str(value or "").strip()
    if not text:
        return None
    candidate = Path(text)
    return (candidate if candidate.is_absolute() else base / candidate).resolve()


def manifest_paths(
    value: object,
    field: str,
    base: Path,
    errors: list[str],
) -> list[Path]:
    if not isinstance(value, list):
        errors.append(f"provenance 字段必须是数组：{field}")
        return []
    result: list[Path] = []
    for item in value:
        path = manifest_path(item, base)
        if path is None:
            errors.append(f"provenance 字段含空路径：{field}")
        else:
            result.append(path)
    return result


def run_git(repo: Path, *args: str) -> tuple[int, list[str]]:
    completed = subprocess.run(
        ["git", *args],
        cwd=repo,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    return completed.returncode, lines


def audit_manifest(
    root: Path,
    manifest: Path,
    errors: list[str],
    warnings: list[str],
    required: bool = True,
    required_paths: set[Path] | None = None,
) -> dict[str, Any]:
    result: dict[str, Any] = {"path": str(manifest), "checked": 0, "mismatches": []}
    if not manifest.is_file():
        (errors if required else warnings).append(f"未发现 SHA-256 清单：{manifest}")
        result["present"] = False
        return result

    result["present"] = True
    seen: set[str] = set()
    verified: set[Path] = set()
    pattern = re.compile(r"^([0-9A-Fa-f]{64})\s+\*?(.+?)\s*$")
    for number, line in enumerate(manifest.read_text(encoding="utf-8-sig").splitlines(), start=1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        match = pattern.match(line)
        if not match:
            errors.append(f"SHA-256 清单第 {number} 行格式错误")
            continue
        expected, relative_text = match.groups()
        normalized = relative_text.replace("\\", "/")
        if normalized in seen:
            errors.append(f"SHA-256 清单重复路径：{normalized}")
            continue
        seen.add(normalized)
        candidate = (root / Path(*normalized.split("/"))).resolve()
        if candidate == manifest.resolve():
            errors.append("SHA-256 清单不得包含自身")
            continue
        if not inside(candidate, root):
            errors.append(f"SHA-256 路径超出交付目录：{normalized}")
            continue
        if not candidate.is_file():
            errors.append(f"SHA-256 清单文件不存在：{normalized}")
            continue
        actual = sha256(candidate)
        result["checked"] += 1
        if actual.lower() != expected.lower():
            result["mismatches"].append(normalized)
            errors.append(f"SHA-256 不一致：{normalized}")
        else:
            verified.add(candidate)
    if result["checked"] == 0:
        errors.append("SHA-256 清单没有可核验条目")
    for missing in sorted((required_paths or set()) - verified):
        errors.append(f"SHA-256 清单未覆盖交付文件：{missing}")
    return result


def audit_provenance(
    path: Path,
    delivery_root: Path,
    errors: list[str],
    warnings: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    record: dict[str, Any] = {"path": str(path), "checked_hashes": 0}
    expectations: dict[str, Any] = {}
    if not path.is_file():
        errors.append(f"provenance 文件不存在：{path}")
        return record, expectations
    if not inside(path, delivery_root):
        errors.append(f"provenance 文件必须位于交付目录内：{path}")

    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        errors.append(f"provenance 无法解析：{path}: {exc}")
        return record, expectations
    if not isinstance(data, dict):
        errors.append("provenance 顶层必须是对象")
        return record, expectations

    base = delivery_root
    workspace = manifest_path(data.get("workspace_root"), base)
    if workspace is None or not workspace.is_dir():
        errors.append("provenance.workspace_root 必须指向存在的目录")
    elif not inside(delivery_root, workspace):
        errors.append("交付目录必须位于 provenance.workspace_root 内")

    artifact_type = str(data.get("artifact_type", "")).strip()
    allowed_artifact_types = {
        "application_docx",
        "review_pdf",
        "application_sources",
        "xml_preparation",
        "analysis_bundle",
        "section_bundle",
    }
    if artifact_type not in allowed_artifact_types:
        errors.append(
            "provenance.artifact_type 无效：" + ", ".join(sorted(allowed_artifact_types))
        )
    topic = str(data.get("invention_topic", "")).strip()
    if not topic:
        errors.append("provenance.invention_topic 不能为空")
    review_label = str(data.get("review_label", "")).strip()
    output_pdf = manifest_path(data.get("output_pdf"), base)
    if artifact_type == "review_pdf" or (artifact_type == "application_docx" and output_pdf is not None):
        if review_label != REVIEW_LABEL:
            errors.append(f"校对 PDF 必须使用精确标记：{REVIEW_LABEL}")
        if output_pdf is None or not output_pdf.is_file():
            errors.append("review_pdf 必须声明存在的 output_pdf")
        elif not inside(output_pdf, delivery_root):
            errors.append("output_pdf 必须位于交付目录内")

    requires_canonical = artifact_type in {
        "application_docx", "review_pdf", "application_sources", "xml_preparation"
    }
    canonical = data.get("canonical_application_sources")
    if not isinstance(canonical, dict):
        if requires_canonical:
            errors.append("canonical_application_sources 必须是对象")
        canonical = {}
    abstract = manifest_path(canonical.get("abstract"), base)
    claims = manifest_path(canonical.get("claims"), base)
    description = manifest_path(canonical.get("description"), base)
    drawings = manifest_paths(canonical.get("drawings", []), "canonical_application_sources.drawings", base, errors)
    abstract_figure = manifest_path(canonical.get("abstract_figure"), base)

    canonical_text = {
        "abstract": abstract,
        "claims": claims,
        "description": description,
    }
    if artifact_type == "application_docx":
        expected_names = {
            "abstract": "04_说明书摘要.md",
            "claims": "02_权利要求书.md",
            "description": "03_说明书.md",
        }
        for role, name in expected_names.items():
            if canonical_text[role] != (delivery_root / name).resolve():
                errors.append(f"分件规范源 {role} 必须指向 {name}")
    for role, source in canonical_text.items():
        if source is None or not source.is_file():
            if not requires_canonical:
                continue
            errors.append(f"规范申请源不存在：{role}")
        elif source.suffix.casefold() not in TEXT_EXTENSIONS:
            errors.append(f"规范申请源必须是 Markdown 或文本：{source}")
    for drawing in drawings:
        if not drawing.is_file() or drawing.suffix.casefold() not in RASTER_EXTENSIONS:
            errors.append(f"规范附图不存在或格式不支持：{drawing}")
    if drawings:
        if abstract_figure is None or abstract_figure not in drawings:
            errors.append("有附图时 abstract_figure 必须指向 drawings 中的一幅图")
    elif requires_canonical and abstract_figure is not None:
        errors.append("没有附图时 abstract_figure 必须为空")

    implementation = manifest_paths(
        data.get("implementation_evidence", []),
        "implementation_evidence",
        base,
        errors,
    )
    if not implementation:
        errors.append("implementation_evidence 至少需要一个当前事实来源")
    layout = manifest_path(data.get("layout_authority"), base)
    if requires_canonical and (layout is None or not layout.is_file()):
        errors.append("layout_authority 必须指向存在的版式合同")

    pipeline = data.get("build_pipeline")
    if not isinstance(pipeline, dict):
        if requires_canonical:
            errors.append("build_pipeline 必须是对象")
        pipeline = {}
    builders = manifest_paths(pipeline.get("document_builders", []), "build_pipeline.document_builders", base, errors)
    exporters = manifest_paths(pipeline.get("pdf_exporters", []), "build_pipeline.pdf_exporters", base, errors)
    orchestrators = manifest_paths(pipeline.get("orchestrators", []), "build_pipeline.orchestrators", base, errors)
    if artifact_type == "review_pdf" and not exporters:
        warnings.append("provenance 未记录 PDF 导出脚本；若使用平台托管导出，请在交付说明中注明")

    disclosure = manifest_paths(data.get("disclosure_materials", []), "disclosure_materials", base, errors)
    raw_publications = manifest_paths(
        data.get("raw_publication_references", []),
        "raw_publication_references",
        base,
        errors,
    )
    roles: dict[str, list[Path]] = {
        "canonical_text": [item for item in canonical_text.values() if item is not None],
        "canonical_drawings": drawings,
        "implementation_evidence": implementation,
        "layout_authority": [layout] if layout is not None else [],
        "document_builders": builders,
        "pdf_exporters": exporters,
        "orchestrators": orchestrators,
        "disclosure_materials": disclosure,
        "raw_publication_references": raw_publications,
    }
    owners: dict[str, str] = {}
    for role, sources in roles.items():
        for source in sources:
            key = str(source).casefold()
            previous = owners.setdefault(key, role)
            if previous != role:
                errors.append(f"来源角色冲突：{source} 同时属于 {previous} 和 {role}")
            if not source.is_file():
                errors.append(f"来源文件不存在：{source}")
            elif role in {"canonical_text", "canonical_drawings", "disclosure_materials"} and not inside(source, delivery_root):
                errors.append(f"交付产物来源超出案件版本目录：{source}")
            elif role == "implementation_evidence" and workspace is not None and not inside(source, workspace):
                errors.append(f"实现证据超出 workspace_root：{source}")

    official_urls = data.get("official_rule_authorities", [])
    if not isinstance(official_urls, list) or not official_urls:
        if requires_canonical:
            errors.append("provenance 未记录当前 CNIPA 官方规则来源")
    else:
        for value in official_urls:
            parsed = urlparse(str(value).strip())
            host = (parsed.hostname or "").casefold()
            if parsed.scheme != "https" or not (host == "cnipa.gov.cn" or host.endswith(".cnipa.gov.cn")):
                errors.append(f"非 CNIPA HTTPS 官方规则来源：{value}")

    if abstract is not None and abstract.is_file():
        abstract_source = abstract.read_text(encoding="utf-8-sig").strip()
        abstract_match = re.search(
            r"(?ms)^\s*##\s*摘要\s*$\s*(.+?)(?=^\s*##\s|\Z)",
            abstract_source,
        )
        abstract_text = (
            abstract_match.group(1).strip() if abstract_match else abstract_source
        )
        if not abstract_text:
            errors.append("规范摘要为空")
        if len(compact_text(abstract_text)) > 300:
            warnings.append("规范摘要超过 300 字；提交前按当期 CNIPA 校验规则复核")
    else:
        abstract_text = ""

    claim_blocks: list[str] = []
    if claims is not None and claims.is_file():
        claim_text = claims.read_text(encoding="utf-8-sig")
        matches = list(
            re.finditer(
                r"(?ms)^\s*(\d+)[.、]\s*(.+?)(?=^\s*\d+[.、]\s*|\Z)",
                claim_text,
            )
        )
        if not matches:
            errors.append("规范权利要求未识别出连续编号条目")
        for expected_number, match in enumerate(matches, start=1):
            actual_number = int(match.group(1))
            if actual_number != expected_number:
                errors.append(
                    f"规范权利要求编号不连续：期望 {expected_number}，实际 {actual_number}"
                )
            claim_blocks.append(match.group(2).strip())

    if description is not None and description.is_file():
        description_text = description.read_text(encoding="utf-8-sig")
        normalized_description = compact_text(description_text)
        required_headings = ["技术领域", "背景技术", "发明内容", "具体实施方式"]
        if drawings:
            required_headings.append("附图说明")
        for heading in required_headings:
            if heading not in normalized_description:
                errors.append(f"规范说明书缺少章节：{heading}")
    else:
        description_text = ""

    source_hashes = data.get("source_hashes")
    normalized_hashes: dict[str, str] = {}
    if not isinstance(source_hashes, dict):
        errors.append("source_hashes 必须是对象")
    else:
        for key, value in source_hashes.items():
            source = manifest_path(key, base)
            if source is None:
                errors.append("source_hashes 含空路径")
            else:
                normalized_hashes[str(source).casefold()] = str(value).strip().casefold()

    required_hashed = (
        [item for item in canonical_text.values() if item is not None]
        + drawings
        + implementation
        + ([layout] if layout is not None else [])
        + builders
        + exporters
        + orchestrators
    )
    required_keys = {str(source).casefold() for source in required_hashed}
    for source in required_hashed:
        if not source.is_file():
            continue
        expected_hash = normalized_hashes.get(str(source).casefold())
        if not expected_hash:
            errors.append(f"source_hashes 缺少来源：{source}")
        elif expected_hash != sha256(source):
            errors.append(f"来源 SHA-256 不一致：{source}")
        else:
            record["checked_hashes"] += 1
    for extra in sorted(set(normalized_hashes) - required_keys):
        errors.append(f"source_hashes 含未分配路径：{extra}")

    identity = data.get("confirmed_identity")
    applicant = ""
    inventors: list[str] = []
    if isinstance(identity, dict):
        applicant = str(identity.get("applicant", "")).strip()
        raw_inventors = identity.get("inventors", [])
        if isinstance(raw_inventors, list):
            inventors = [str(value).strip() for value in raw_inventors if str(value).strip()]
        else:
            errors.append("confirmed_identity.inventors 必须是数组")
    elif identity is not None:
        errors.append("confirmed_identity 必须是对象")

    record.update(
        {
            "artifact_type": artifact_type,
            "invention_topic": topic,
            "output_pdf": str(output_pdf) if output_pdf else None,
            "canonical_drawings": len(drawings),
            "canonical_claims": len(claim_blocks),
            "assigned_sources": sum(len(items) for items in roles.values()),
        }
    )
    expectations.update(
        {
            "artifact_type": artifact_type,
            "topic": topic,
            "review_label": review_label,
            "output_pdf": output_pdf,
            "abstract_text": abstract_text,
            "claims": claim_blocks,
            "description_text": description_text,
            "drawing_count": len(drawings),
            "drawings": drawings,
            "abstract_figure": abstract_figure,
            "applicant": applicant,
            "inventors": inventors,
        }
    )
    return record, expectations


def audit_docx(
    path: Path, errors: list[str], warnings: list[str], *, expected_drawings: list[Path] | None = None
) -> dict[str, Any]:
    record: dict[str, Any] = {"path": str(path), "sha256": sha256(path)}
    try:
        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                errors.append(f"DOCX 压缩成员损坏：{path}::{bad_member}")
            names = set(archive.namelist())
            if "word/document.xml" not in names:
                errors.append(f"DOCX 缺少 word/document.xml：{path}")
                return record
            document = ET.fromstring(archive.read("word/document.xml"))
            inline = len(document.findall(".//wp:inline", NS))
            anchors = len(document.findall(".//wp:anchor", NS))
            revisions = sum(
                len(document.findall(f".//w:{name}", NS))
                for name in ("ins", "del", "moveFrom", "moveTo")
            )
            alt_texts = []
            for item in document.findall(".//wp:docPr", NS):
                value = (item.attrib.get("descr") or item.attrib.get("title") or "").strip()
                alt_texts.append(value)
            media = [name for name in names if name.startswith("word/media/") and not name.endswith("/")]
            if expected_drawings:
                embedded = {hashlib.sha256(archive.read(name)).hexdigest() for name in media}
                expected = {sha256(drawing) for drawing in expected_drawings if drawing.is_file()}
                if not (inline or anchors) or embedded != expected:
                    errors.append(f"附图 Word 内嵌媒体与规范附图不一致：{path.name}")
            comments = any(name.startswith("word/comments") for name in names)
            custom_properties = "docProps/custom.xml" in names
            forbidden_parts = sorted(
                name
                for name in names
                if name in {"word/vbaProject.bin", "word/people.xml"}
                or name.startswith("word/embeddings/")
                or name.startswith("word/activeX/")
                or name.startswith("customXml/")
            )
            hidden_nodes = sum(
                len(document.findall(f".//w:{name}", NS))
                for name in ("vanish", "webHidden")
            )
            external_relationships: list[str] = []
            for relationship_name in sorted(name for name in names if name.endswith(".rels")):
                try:
                    relationships = ET.fromstring(archive.read(relationship_name))
                except ET.ParseError as exc:
                    errors.append(f"DOCX 关系部件无法解析：{path}::{relationship_name}: {exc}")
                    continue
                for relationship in relationships.findall("./r:Relationship", NS):
                    if relationship.attrib.get("TargetMode") == "External":
                        external_relationships.append(
                            f"{relationship_name}:{relationship.attrib.get('Target', '')}"
                        )

            core_metadata = {"creator": "", "last_modified_by": ""}
            if "docProps/core.xml" in names:
                core = ET.fromstring(archive.read("docProps/core.xml"))
                creator = core.find("dc:creator", NS)
                modified = core.find("cp:lastModifiedBy", NS)
                core_metadata = {
                    "creator": (creator.text or "").strip() if creator is not None else "",
                    "last_modified_by": (modified.text or "").strip() if modified is not None else "",
                }
            extended_metadata = {"company": "", "manager": ""}
            if "docProps/app.xml" in names:
                app = ET.fromstring(archive.read("docProps/app.xml"))
                company = app.find("ep:Company", NS)
                manager = app.find("ep:Manager", NS)
                extended_metadata = {
                    "company": (company.text or "").strip() if company is not None else "",
                    "manager": (manager.text or "").strip() if manager is not None else "",
                }
            record.update(
                {
                    "inline_images": inline,
                    "anchored_images": anchors,
                    "media_files": len(media),
                    "alt_texts": alt_texts,
                    "comments": comments,
                    "tracked_revision_nodes": revisions,
                    "hidden_nodes": hidden_nodes,
                    "custom_properties": custom_properties,
                    "forbidden_parts": forbidden_parts,
                    "external_relationships": external_relationships,
                    "core_metadata": core_metadata,
                    "extended_metadata": extended_metadata,
                }
            )
            if comments:
                errors.append(f"DOCX 含批注：{path}")
            if revisions:
                errors.append(f"DOCX 含未处理修订：{path}")
            if hidden_nodes:
                errors.append(f"DOCX 含隐藏文字节点：{path} ({hidden_nodes})")
            if custom_properties:
                errors.append(f"DOCX 含自定义属性：{path}")
            if forbidden_parts:
                errors.append(f"DOCX 含宏、嵌入对象或自定义 XML：{path}: {forbidden_parts}")
            if external_relationships:
                errors.append(f"DOCX 含外部关系或链接：{path}: {external_relationships}")
            if any(core_metadata.values()) or any(extended_metadata.values()):
                errors.append(
                    "DOCX 含作者、最近修改者、公司或管理者元数据："
                    f"{path}: core={core_metadata} extended={extended_metadata}"
                )
            if anchors:
                warnings.append(f"DOCX 含 {anchors} 个浮动图形，请确认不会漂移：{path}")
            if alt_texts and any(not value for value in alt_texts):
                warnings.append(f"DOCX 存在空图片替代文字：{path}")
    except (zipfile.BadZipFile, ET.ParseError, OSError) as exc:
        errors.append(f"DOCX 无法解析：{path}: {exc}")
    return record


def audit_drawio(path: Path, errors: list[str]) -> dict[str, Any]:
    record: dict[str, Any] = {"path": str(path), "sha256": sha256(path)}
    try:
        root = ET.parse(path).getroot()
        cells = root.findall(".//mxCell")
        record["vertices"] = sum(cell.get("vertex") == "1" for cell in cells)
        record["edges"] = sum(cell.get("edge") == "1" for cell in cells)
    except (ET.ParseError, OSError) as exc:
        errors.append(f"Draw.io XML 无法解析：{path}: {exc}")
    return record


def audit_pdf(
    path: Path,
    errors: list[str],
    warnings: list[str],
    expectations: dict[str, Any] | None = None,
) -> dict[str, Any]:
    record: dict[str, Any] = {"path": str(path), "sha256": sha256(path)}
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError:
        errors.append(f"未安装 pypdf，无法完成 PDF 页数、加密和内容检查：{path}")
        return record
    try:
        reader = PdfReader(str(path))
        record["encrypted"] = bool(reader.is_encrypted)
        record["pages"] = len(reader.pages)
        if reader.is_encrypted:
            errors.append(f"PDF 已加密，无法可靠审计：{path}")
        if not reader.pages:
            errors.append(f"PDF 没有页面：{path}")
            return record

        page_texts: list[str] = []
        page_sizes: list[list[float]] = []
        for index, page in enumerate(reader.pages, start=1):
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            page_sizes.append([round(width, 2), round(height, 2)])
            try:
                page_texts.append(page.extract_text() or "")
            except Exception as exc:
                page_texts.append("")
                warnings.append(f"PDF 第 {index} 页文本提取失败：{path}: {exc}")
        record["page_sizes_pt"] = page_sizes
        metadata = reader.metadata or {}
        record["metadata"] = {
            "title": str(metadata.get("/Title", "") or ""),
            "subject": str(metadata.get("/Subject", "") or ""),
            "author": str(metadata.get("/Author", "") or ""),
        }

        if expectations and expectations.get("artifact_type") == "application_docx":
            cover = compact_text(page_texts[0]) if page_texts else ""
            if cover.count(compact_text(REVIEW_LABEL)) != 1:
                errors.append("校对 PDF 首页必须且只能出现一次精确状态标记")

        if expectations and expectations.get("artifact_type") == "review_pdf":
            for index, (width, height) in enumerate(page_sizes, start=1):
                if abs(width - A4_WIDTH) > 3 or abs(height - A4_HEIGHT) > 3:
                    errors.append(
                        f"校对 PDF 第 {index} 页不是 A4：{width:.2f} × {height:.2f} pt"
                    )

            compact_pages = [compact_text(value) for value in page_texts]
            compact_pdf = "".join(compact_pages)
            cover = compact_pages[0] if compact_pages else ""
            label = compact_text(str(expectations.get("review_label", "")))
            if not label or cover.count(label) != 1:
                errors.append("校对 PDF 首页必须且只能出现一次精确状态标记")
            topic = compact_text(str(expectations.get("topic", "")))
            if topic and topic not in cover:
                errors.append("校对 PDF 首页缺少发明名称")
            for identity in [
                str(expectations.get("applicant", "")),
                *[str(value) for value in expectations.get("inventors", [])],
            ]:
                if identity and compact_text(identity) not in cover:
                    errors.append(f"校对 PDF 首页缺少已确认身份：{identity}")

            abstract_text = compact_text(str(expectations.get("abstract_text", "")))
            if abstract_text and abstract_text not in compact_pdf:
                errors.append("校对 PDF 未完整包含规范摘要")

            previous = -1
            for number, claim in enumerate(expectations.get("claims", []), start=1):
                needle = compact_text(f"{number}.{claim}")
                position = compact_pdf.find(needle)
                if position < 0:
                    errors.append(f"校对 PDF 未完整包含规范权利要求 {number}")
                elif position <= previous:
                    errors.append(f"校对 PDF 权利要求顺序错误：{number}")
                else:
                    previous = position

            claim_page = next(
                (index for index, value in enumerate(compact_pages) if "权利要求书" in value),
                None,
            )
            description_page = next(
                (
                    index
                    for index, value in enumerate(compact_pages)
                    if "技术领域" in value and "背景技术" in value
                ),
                None,
            )
            drawing_page = next(
                (
                    index
                    for index, value in enumerate(compact_pages)
                    if "说明书附图" in value
                    and (description_page is None or index > description_page)
                ),
                None,
            )
            if claim_page is None:
                errors.append("校对 PDF 缺少权利要求书部分")
            if description_page is None:
                errors.append("校对 PDF 缺少可识别的说明书部分")
            if claim_page is not None and description_page is not None and claim_page >= description_page:
                errors.append("校对 PDF 中权利要求书必须先于说明书")
            if expectations.get("drawing_count"):
                if drawing_page is None:
                    errors.append("有规范附图但校对 PDF 缺少说明书附图部分")
                elif description_page is not None and drawing_page <= description_page:
                    errors.append("校对 PDF 中说明书附图必须位于说明书之后")
    except Exception as exc:  # pypdf may raise several parser-specific exceptions
        errors.append(f"PDF 无法解析：{path}: {exc}")
    return record


def audit_git(
    root: Path,
    repo: Path,
    require_isolated: bool,
    errors: list[str],
    warnings: list[str],
) -> dict[str, Any]:
    record: dict[str, Any] = {"repo": str(repo), "root": str(root)}
    if not repo.is_dir() or not (repo / ".git").exists():
        message = f"Git 仓库无效：{repo}"
        (errors if require_isolated else warnings).append(message)
        return record
    if not inside(root, repo):
        message = f"交付目录不在指定仓库内：{root}"
        (errors if require_isolated else warnings).append(message)
        return record
    relative = root.relative_to(repo).as_posix()
    _, tracked = run_git(repo, "ls-files", "--", relative)
    _, staged = run_git(repo, "diff", "--cached", "--name-only", "--", relative)
    _, unignored = run_git(repo, "status", "--short", "--untracked-files=all", "--", relative)
    _, ignore_rule = run_git(repo, "check-ignore", "-v", "--", relative)
    record.update(
        {
            "relative": relative,
            "tracked": tracked,
            "staged": staged,
            "unignored": unignored,
            "ignore_rule": ignore_rule,
        }
    )
    if require_isolated:
        if tracked:
            errors.append(f"保密目录包含 {len(tracked)} 个 Git 已跟踪文件")
        if staged:
            errors.append(f"保密目录包含 {len(staged)} 个 Git 已暂存文件")
        if unignored:
            errors.append(f"保密目录包含 {len(unignored)} 个未忽略改动或文件")
    return record


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="以失败关闭方式审计中国发明专利交付目录。")
    parser.add_argument("--root", required=True, help="案件版本目录")
    parser.add_argument("--repo", help="所属 Git 仓库根目录")
    parser.add_argument("--manifest", help="SHA-256 清单；默认 <root>/02-文件校验.sha256")
    parser.add_argument("--provenance", help="provenance.json；full 配置默认读取 <root>/99-内部校验/provenance.json")
    parser.add_argument("--json-out", help="可选 JSON 审计回执")
    parser.add_argument(
        "--profile",
        choices=("full", "analysis", "section"),
        default="full",
        help="审计配置；默认 full",
    )
    parser.add_argument(
        "--allow-partial",
        action="store_true",
        help="仅对 analysis/section 放宽标准根文件和来源记录要求",
    )
    parser.add_argument("--require-git-isolated", action="store_true")
    args = parser.parse_args()

    if args.profile == "full" and args.allow_partial:
        parser.error("full 配置不允许使用 --allow-partial")

    root = Path(args.root).resolve()
    if not root.is_dir():
        parser.error(f"交付目录不存在：{root}")

    errors: list[str] = []
    warnings: list[str] = []
    root_files = ("00-优化说明.md", "01-交付文件清单.md", "02-文件校验.sha256")
    substantive_files: list[Path] = []
    for candidate in sorted(path for path in root.rglob("*") if path.is_file()):
        relative = candidate.relative_to(root)
        if len(relative.parts) == 1 and relative.name in root_files:
            continue
        if relative.parts and relative.parts[0] == "99-内部校验":
            continue
        substantive_files.append(candidate)
    if not substantive_files:
        errors.append("交付目录没有任何实质申请、分析、交底书或附图文件")
    for candidate in substantive_files:
        if candidate.stat().st_size == 0:
            errors.append(f"实质交付文件为空：{candidate}")

    manifest = Path(args.manifest).resolve() if args.manifest else root / "02-文件校验.sha256"
    result: dict[str, Any] = {
        "root": str(root),
        "profile": args.profile,
        "allow_partial": args.allow_partial,
        "substantive_files": len(substantive_files),
        "docx": [],
        "drawio": [],
        "pdf": [],
    }
    provenance_expectations: dict[str, Any] = {}
    provenance_path: Path | None = None
    if args.provenance:
        provenance_path = Path(args.provenance).resolve()
    elif args.profile == "full":
        provenance_path = root / "99-内部校验" / "provenance.json"
    if provenance_path is not None:
        provenance_record, provenance_expectations = audit_provenance(
            provenance_path,
            root,
            errors,
            warnings,
        )
        result["provenance"] = provenance_record
        expected_types = {
            "full": {"application_docx", "review_pdf", "application_sources", "xml_preparation"},
            "analysis": {"analysis_bundle"},
            "section": {"section_bundle"},
        }[args.profile]
        artifact_type = str(provenance_expectations.get("artifact_type", ""))
        if artifact_type and artifact_type not in expected_types:
            errors.append(
                f"provenance.artifact_type 与 {args.profile} 配置不匹配：{artifact_type}"
            )
        if artifact_type == "xml_preparation":
            warnings.append("XML 文件只做来源和文件存在性审计；本回执不代表通过 CNIPA 官方校验")
    elif not args.allow_partial:
        warnings.append("analysis/section 配置未提供 provenance；来源深审未执行")

    separate_docx = provenance_expectations.get("artifact_type") == "application_docx"
    drawing_count = provenance_expectations.get("drawing_count", 0)
    if separate_docx:
        required_docs = set(REQUIRED_PACKAGE_DOCS)
        if drawing_count:
            required_docs.update(DRAWING_PACKAGE_DOCS)
        for stem in PACKAGE_DOCS:
            files = [root / f"{stem}.{extension}" for extension in ("md", "docx")]
            if stem in required_docs or any(path.exists() for path in files):
                for path in files:
                    if not path.is_file() or not inside(path, root):
                        errors.append(f"分件交付缺少有效文件：{path.name}")
    else:
        for required in root_files:
            if not (root / required).is_file():
                (warnings if args.allow_partial else errors).append(f"缺少标准根文件：{required}")
    result["manifest"] = audit_manifest(
        root,
        manifest,
        errors,
        warnings,
        required=not args.allow_partial,
        required_paths={path.resolve() for path in substantive_files} if separate_docx else None,
    )

    for path in sorted(root.rglob("*.docx")):
        expected_drawings = None
        if separate_docx and drawing_count and path.parent == root:
            if path.stem == DRAWING_PACKAGE_DOCS[0]:
                expected_drawings = provenance_expectations["drawings"]
            elif path.stem == DRAWING_PACKAGE_DOCS[1] and provenance_expectations["abstract_figure"]:
                expected_drawings = [provenance_expectations["abstract_figure"]]
        result["docx"].append(audit_docx(path, errors, warnings, expected_drawings=expected_drawings))
    for path in sorted(root.rglob("*.drawio")):
        result["drawio"].append(audit_drawio(path, errors))
    for path in sorted(root.rglob("*.pdf")):
        expected_pdf = provenance_expectations.get("output_pdf")
        deep_expectations = (
            provenance_expectations
            if isinstance(expected_pdf, Path) and path.resolve() == expected_pdf
            else None
        )
        result["pdf"].append(audit_pdf(path, errors, warnings, deep_expectations))

    if args.profile == "full" and not result["docx"]:
        errors.append("full 配置至少需要一份 DOCX 技术交底或申请材料")

    expected_pdf = provenance_expectations.get("output_pdf")
    if isinstance(expected_pdf, Path) and not any(
        Path(item["path"]).resolve() == expected_pdf for item in result["pdf"]
    ):
        errors.append(f"provenance 声明的 output_pdf 未进入交付目录审计：{expected_pdf}")

    if args.repo:
        result["git"] = audit_git(
            root,
            Path(args.repo).resolve(),
            args.require_git_isolated,
            errors,
            warnings,
        )
    elif args.require_git_isolated:
        errors.append("使用 --require-git-isolated 时必须提供 --repo")

    result["errors"] = errors
    result["warnings"] = warnings
    result["passed"] = not errors

    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        destination = Path(args.json_out).resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(output + "\n", encoding="utf-8")
    print(output)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
