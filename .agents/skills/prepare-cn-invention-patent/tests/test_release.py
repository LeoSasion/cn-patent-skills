from __future__ import annotations

import base64
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from lxml import etree


SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL_ROOT / "scripts"
ASSET = SKILL_ROOT / "assets" / "技术交底书_空白模板.docx"
sys.path.insert(0, str(SCRIPTS))

import audit_delivery_bundle as audit  # noqa: E402
import check_template_hash as template_check  # noqa: E402
from docx_package import load_parts, write_parts  # noqa: E402
from fill_disclosure_template import REQUIRED_TAGS  # noqa: E402
from make_output_paths import DRAWING_PACKAGE_DOCS, REQUIRED_PACKAGE_DOCS  # noqa: E402


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class SeparateWordTests(unittest.TestCase):
    def run_script(self, script: str, *args: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(SCRIPTS / script), *args],
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=False,
        )

    def test_split_docx_package_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workspace = Path(directory)
            root = workspace / "bundle"
            internal = root / "99-内部校验"
            internal.mkdir(parents=True)
            diagram = root / "图1.png"
            diagram.write_bytes(base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jJ1kAAAAASUVORK5CYII="
            ))
            for name in (*REQUIRED_PACKAGE_DOCS, *DRAWING_PACKAGE_DOCS):
                markdown = root / f"{name}.md"
                body = {
                    "02_权利要求书": "1. 一种采样方法。\n\n2. 根据权利要求1所述的采样方法。\n",
                    "03_说明书": "## 技术领域\n采样。\n## 背景技术\n时序。\n## 发明内容\n校验。\n## 附图说明\n见图1。\n## 具体实施方式\n先读取时序。\n",
                    "05_说明书附图": "![图1 采样流程](图1.png)\n",
                    "06_摘要附图": "![图1 采样流程](图1.png)\n",
                }.get(name, "一种采样方法的模拟测试。\n")
                markdown.write_text(f"# {name.split('_')[-1]}\n\n{body}", encoding="utf-8")
                result = self.run_script(
                    "render_patent_docx.py", "-i", str(markdown), "-o", str(markdown.with_suffix(".docx"))
                )
                self.assertEqual(0, result.returncode, result.stderr)

            claims = Document(root / "02_权利要求书.docx")
            section = claims.sections[0]
            self.assertAlmostEqual(210, section.page_width.mm, places=1)
            self.assertAlmostEqual(297, section.page_height.mm, places=1)
            self.assertEqual("2. 根据权利要求1所述的采样方法。", claims.paragraphs[-1].text)
            self.assertEqual(1, len(Document(root / "06_摘要附图.docx").inline_shapes))

            evidence = workspace / "evidence.md"
            evidence.write_text("模拟证据。\n", encoding="utf-8")
            layout = SKILL_ROOT / "references" / "document-contracts.md"
            builder = SCRIPTS / "render_patent_docx.py"
            canonical = {
                "abstract": "04_说明书摘要.md",
                "claims": "02_权利要求书.md",
                "description": "03_说明书.md",
            }
            source_files = [root / path for path in canonical.values()]
            provenance = {
                "artifact_type": "application_docx",
                "workspace_root": str(workspace),
                "invention_topic": "一种采样方法",
                "canonical_application_sources": {
                    **canonical, "drawings": [diagram.name], "abstract_figure": diagram.name
                },
                "implementation_evidence": [str(evidence)],
                "layout_authority": str(layout),
                "official_rule_authorities": ["https://www.cnipa.gov.cn/art/2023/12/21/art_98_189197.html"],
                "build_pipeline": {"document_builders": [str(builder)]},
                "source_hashes": {
                    str(path): digest(path) for path in [*source_files, diagram, evidence, layout, builder]
                },
            }
            (internal / "provenance.json").write_text(
                json.dumps(provenance, ensure_ascii=False), encoding="utf-8"
            )
            manifest = root / "02-文件校验.sha256"
            entries = {path.name: f"{digest(path)}  {path.name}" for path in root.iterdir() if path.is_file()}
            manifest.write_text("\n".join(entries.values()) + "\n", encoding="utf-8")
            command = ("audit_delivery_bundle.py", "--root", str(root))
            result = self.run_script(*command)
            self.assertEqual(0, result.returncode, result.stdout + result.stderr)

            lost = "06_摘要附图.docx"
            entries.pop(lost)
            manifest.write_text("\n".join(entries.values()) + "\n", encoding="utf-8")
            result = self.run_script(*command)
            self.assertEqual(1, result.returncode, result.stdout)
            self.assertIn("SHA-256 清单未覆盖交付文件", result.stdout)

            (root / lost).unlink()
            result = self.run_script(*command)
            self.assertEqual(1, result.returncode, result.stdout)
            self.assertIn(f"分件交付缺少有效文件：{lost}", result.stdout)

    def test_docx_rejects_unrendered_diagram_and_out_of_scope_image(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            markdown = root / "输入.md"
            output = root / "输出.docx"
            for source, error in [
                ("```mermaid\nflowchart TB\n```\n", "Mermaid 缺少"),
                ("![图](../outside.png)\n", "图片路径超出"),
            ]:
                markdown.write_text(source, encoding="utf-8")
                result = self.run_script("render_patent_docx.py", "-i", str(markdown), "-o", str(output))
                self.assertEqual(1, result.returncode, result.stderr)
                self.assertIn(error, result.stderr)
                self.assertFalse(output.exists())


class TemplateTests(unittest.TestCase):
    def test_bundled_template_contract(self) -> None:
        report, errors = template_check.inspect_template(ASSET)
        self.assertEqual([], errors)
        self.assertEqual(template_check.BUNDLED_SHA256, report["sha256"])
        self.assertEqual(set(REQUIRED_TAGS), set(report["content_control_tags"]))

    def test_builder_is_deterministic(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            first = Path(directory) / "first.docx"
            second = Path(directory) / "second.docx"
            for output in (first, second):
                completed = subprocess.run(
                    [sys.executable, str(SCRIPTS / "build_disclosure_template.py"), "--output", str(output)],
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                )
                self.assertEqual(0, completed.returncode, completed.stderr)
            self.assertEqual(digest(first), digest(second))

    def test_filler_replaces_controls_and_clears_identity(self) -> None:
        values = {tag: f"测试-{tag}" for tag in REQUIRED_TAGS}
        values.update(
            {
                "CONTACT_NAME": None,
                "CONTACT_PHONE": None,
                "CONTACT_FAX": None,
                "CONTACT_EMAIL": None,
                "SECTION_4_SOLUTION": "步骤一。\n步骤二。",
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            values_path = root / "values.json"
            output = root / "filled.docx"
            values_path.write_text(json.dumps(values, ensure_ascii=False), encoding="utf-8")
            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS / "fill_disclosure_template.py"),
                    "--template",
                    str(ASSET),
                    "--values",
                    str(values_path),
                    "--output",
                    str(output),
                    "--require-all",
                ],
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            self.assertEqual(0, completed.returncode, completed.stderr)
            document = etree.fromstring(load_parts(output)["word/document.xml"])
            visible = "".join(document.itertext())
            self.assertIn("步骤一。", visible)
            self.assertIn("步骤二。", visible)
            self.assertNotIn("[姓名]", visible)
            errors: list[str] = []
            warnings: list[str] = []
            audit.audit_docx(output, errors, warnings)
            self.assertEqual([], errors)


class DeliveryAuditTests(unittest.TestCase):
    def run_audit(self, *arguments: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(SCRIPTS / "audit_delivery_bundle.py"), *arguments],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

    def test_empty_directory_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            completed = self.run_audit("--root", directory)
            self.assertEqual(1, completed.returncode)
            report = json.loads(completed.stdout)
            self.assertFalse(report["passed"])
            self.assertTrue(any("没有任何实质" in item for item in report["errors"]))

    def test_partial_analysis_requires_explicit_flag(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            analysis_dir = root / "参考分析"
            analysis_dir.mkdir()
            (analysis_dir / "事实.md").write_text("# 事实\n有证据的分析。\n", encoding="utf-8")
            strict = self.run_audit("--root", str(root), "--profile", "analysis")
            self.assertEqual(1, strict.returncode)
            partial = self.run_audit(
                "--root",
                str(root),
                "--profile",
                "analysis",
                "--allow-partial",
            )
            self.assertEqual(0, partial.returncode, partial.stdout + partial.stderr)
            report = json.loads(partial.stdout)
            self.assertTrue(report["passed"])
            self.assertTrue(report["warnings"])

    def test_docx_private_metadata_and_external_parts_fail(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "unsafe.docx"
            parts = load_parts(ASSET)
            core = etree.fromstring(parts["docProps/core.xml"])
            namespace = "http://purl.org/dc/elements/1.1/"
            creator = core.find(f"{{{namespace}}}creator")
            self.assertIsNotNone(creator)
            creator.text = "Private Person"
            parts["docProps/core.xml"] = etree.tostring(
                core, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            parts["docProps/custom.xml"] = b"<?xml version='1.0' encoding='UTF-8'?><Properties/>"
            write_parts(parts, output)
            errors: list[str] = []
            warnings: list[str] = []
            audit.audit_docx(output, errors, warnings)
            self.assertTrue(any("自定义属性" in item for item in errors))
            self.assertTrue(any("作者" in item for item in errors))

    def test_minimal_full_application_sources_bundle_passes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workspace = Path(directory)
            root = workspace / "bundle"
            application = root / "申请文件"
            disclosure = root / "技术交底书"
            internal = root / "99-内部校验"
            application.mkdir(parents=True)
            disclosure.mkdir(parents=True)
            internal.mkdir(parents=True)
            (root / "00-优化说明.md").write_text("# 说明\n测试包。\n", encoding="utf-8")
            (root / "01-交付文件清单.md").write_text("# 清单\n测试包。\n", encoding="utf-8")
            abstract = application / "说明书摘要.md"
            claims = application / "权利要求书.md"
            description = application / "说明书.md"
            abstract.write_text("一种经验证的方法。", encoding="utf-8")
            claims.write_text("1. 一种方法，包括步骤 A。\n", encoding="utf-8")
            description.write_text(
                "# 技术领域\n内容。\n# 背景技术\n内容。\n# 发明内容\n内容。\n# 具体实施方式\n内容。\n",
                encoding="utf-8",
            )
            disclosure_docx = disclosure / "技术交底书.docx"
            shutil.copyfile(ASSET, disclosure_docx)
            evidence = workspace / "implementation.py"
            layout = workspace / "layout.md"
            evidence.write_text("def implemented(): return True\n", encoding="utf-8")
            layout.write_text("# 版式合同\n", encoding="utf-8")
            sources = [abstract, claims, description, evidence, layout]
            data = {
                "artifact_type": "application_sources",
                "workspace_root": str(workspace),
                "invention_topic": "测试发明",
                "canonical_application_sources": {
                    "abstract": str(abstract.relative_to(root)),
                    "claims": str(claims.relative_to(root)),
                    "description": str(description.relative_to(root)),
                    "drawings": [],
                    "abstract_figure": "",
                },
                "implementation_evidence": [str(evidence)],
                "layout_authority": str(layout),
                "official_rule_authorities": [
                    "https://www.cnipa.gov.cn/art/2025/11/12/art_75_202551.html"
                ],
                "build_pipeline": {
                    "document_builders": [],
                    "pdf_exporters": [],
                    "orchestrators": [],
                },
                "disclosure_materials": [str(disclosure_docx.relative_to(root))],
                "raw_publication_references": [],
                "source_hashes": {str(path): digest(path) for path in sources},
                "confirmed_identity": {"applicant": "", "inventors": []},
            }
            provenance = internal / "provenance.json"
            provenance.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            manifest = root / "02-文件校验.sha256"
            entries = []
            for path in sorted(item for item in root.rglob("*") if item.is_file() and item != manifest):
                entries.append(f"{digest(path)}  {path.relative_to(root).as_posix()}")
            manifest.write_text("\n".join(entries) + "\n", encoding="utf-8")
            completed = self.run_audit("--root", str(root))
            self.assertEqual(0, completed.returncode, completed.stdout + completed.stderr)
            self.assertTrue(json.loads(completed.stdout)["passed"])

    def test_no_drawings_does_not_require_drawing_description(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workspace = Path(directory)
            root = workspace / "bundle"
            application = root / "申请文件"
            internal = root / "99-内部校验"
            application.mkdir(parents=True)
            internal.mkdir(parents=True)
            abstract = application / "说明书摘要.md"
            claims = application / "权利要求书.md"
            description = application / "说明书.md"
            evidence = workspace / "implementation.py"
            layout = workspace / "layout.md"
            abstract.write_text("一种经验证的方法。", encoding="utf-8")
            claims.write_text("1. 一种方法，包括步骤 A。\n", encoding="utf-8")
            description.write_text(
                "# 技术领域\n内容。\n# 背景技术\n内容。\n# 发明内容\n内容。\n# 具体实施方式\n内容。\n",
                encoding="utf-8",
            )
            evidence.write_text("def implemented(): return True\n", encoding="utf-8")
            layout.write_text("# 版式合同\n", encoding="utf-8")
            sources = [abstract, claims, description, evidence, layout]
            data = {
                "artifact_type": "application_sources",
                "workspace_root": str(workspace),
                "invention_topic": "测试发明",
                "canonical_application_sources": {
                    "abstract": str(abstract.relative_to(root)),
                    "claims": str(claims.relative_to(root)),
                    "description": str(description.relative_to(root)),
                    "drawings": [],
                    "abstract_figure": "",
                },
                "implementation_evidence": [str(evidence)],
                "layout_authority": str(layout),
                "official_rule_authorities": [
                    "https://www.cnipa.gov.cn/art/2025/11/12/art_75_202551.html"
                ],
                "build_pipeline": {
                    "document_builders": [],
                    "pdf_exporters": [],
                    "orchestrators": [],
                },
                "disclosure_materials": [],
                "raw_publication_references": [],
                "source_hashes": {str(path): digest(path) for path in sources},
                "confirmed_identity": {"applicant": "", "inventors": []},
            }
            provenance = internal / "provenance.json"
            provenance.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            errors: list[str] = []
            warnings: list[str] = []
            audit.audit_provenance(provenance, root, errors, warnings)
            self.assertFalse(any("附图说明" in item for item in errors), errors)
            self.assertEqual([], errors)

    def test_legacy_xml_source_type_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            internal = root / "99-内部校验"
            internal.mkdir(parents=True)
            provenance = internal / "provenance.json"
            provenance.write_text(
                json.dumps(
                    {
                        "artifact_type": "xml_source",
                        "workspace_root": str(root),
                        "invention_topic": "测试发明",
                        "implementation_evidence": [],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            errors: list[str] = []
            warnings: list[str] = []
            audit.audit_provenance(provenance, root, errors, warnings)
            self.assertTrue(any("xml_preparation" in item for item in errors))


if __name__ == "__main__":
    unittest.main()
