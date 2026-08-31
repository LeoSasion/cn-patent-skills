from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
FIXED_W3CDTF = "2000-01-01T00:00:00Z"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)


def load_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        bad_member = archive.testzip()
        if bad_member:
            raise ValueError(f"DOCX 压缩成员损坏：{bad_member}")
        return {
            name: archive.read(name)
            for name in archive.namelist()
            if not name.endswith("/")
        }


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def sanitize_parts(parts: dict[str, bytes]) -> dict[str, bytes]:
    """Return a privacy-scrubbed, reproducible OOXML part mapping."""

    cleaned = dict(parts)
    cleaned.pop("docProps/custom.xml", None)
    for name in list(cleaned):
        if name.startswith("customXml/"):
            del cleaned[name]

    content_types = cleaned.get("[Content_Types].xml")
    if content_types:
        root = etree.fromstring(content_types)
        for node in root.xpath(
            "./ct:Override[@PartName='/docProps/custom.xml' or starts-with(@PartName, '/customXml/')]",
            namespaces={"ct": CT_NS},
        ):
            node.getparent().remove(node)
        cleaned["[Content_Types].xml"] = _serialize(root)

    root_relationships = cleaned.get("_rels/.rels")
    if root_relationships:
        root = etree.fromstring(root_relationships)
        for node in root.xpath(
            "./r:Relationship[contains(@Type, '/custom-properties')]",
            namespaces={"r": REL_NS},
        ):
            node.getparent().remove(node)
        cleaned["_rels/.rels"] = _serialize(root)

    for name, payload in list(cleaned.items()):
        if not name.endswith(".rels"):
            continue
        root = etree.fromstring(payload)
        changed = False
        for node in root.xpath(
            "./r:Relationship[contains(@Type, '/customXml') or contains(@Target, 'customXml')]",
            namespaces={"r": REL_NS},
        ):
            node.getparent().remove(node)
            changed = True
        if changed:
            cleaned[name] = _serialize(root)

    core = cleaned.get("docProps/core.xml")
    if core:
        root = etree.fromstring(core)
        for xpath in ("./dc:creator", "./cp:lastModifiedBy"):
            nodes = root.xpath(xpath, namespaces={"dc": DC_NS, "cp": CP_NS})
            for node in nodes:
                node.text = ""
        for xpath in ("./dcterms:created", "./dcterms:modified"):
            nodes = root.xpath(xpath, namespaces={"dcterms": DCTERMS_NS})
            for node in nodes:
                node.text = FIXED_W3CDTF
        cleaned["docProps/core.xml"] = _serialize(root)

    app = cleaned.get("docProps/app.xml")
    if app:
        root = etree.fromstring(app)
        for local_name in ("Company", "Manager"):
            nodes = root.xpath(f"./ep:{local_name}", namespaces={"ep": EP_NS})
            for node in nodes:
                node.text = ""
        cleaned["docProps/app.xml"] = _serialize(root)

    rsid_prefix = f"{{{W_NS}}}rsid"
    for name, payload in list(cleaned.items()):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        try:
            root = etree.fromstring(payload)
        except etree.XMLSyntaxError:
            continue
        changed = False
        for element in root.iter():
            for attribute in list(element.attrib):
                if attribute.startswith(rsid_prefix):
                    del element.attrib[attribute]
                    changed = True
        for node in root.xpath(".//w:rsids", namespaces={"w": W_NS}):
            node.getparent().remove(node)
            changed = True
        if changed:
            cleaned[name] = _serialize(root)

    return cleaned


def write_parts(parts: dict[str, bytes], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.tmp")
    if temporary.exists():
        temporary.unlink()
    try:
        with zipfile.ZipFile(
            temporary,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
        ) as archive:
            for name in sorted(parts):
                info = zipfile.ZipInfo(name, FIXED_ZIP_TIME)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0
                archive.writestr(info, parts[name])
        temporary.replace(output)
    finally:
        if temporary.exists():
            temporary.unlink()


def normalize_docx(source: Path, output: Path) -> None:
    write_parts(sanitize_parts(load_parts(source)), output)
