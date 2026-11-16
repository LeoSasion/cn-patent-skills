#!/usr/bin/env python3
# Adapted from HuangXinzhe/cn-patent-drafting. See its MIT notice in ../licenses/.
"""Generate timestamped output paths for Chinese patent case files."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path


FORBIDDEN = r'[\\/:*?"<>|\r\n\t]'
PLACEHOLDERS = ("[待填写]", "【待填写】", "待填写", "[TODO]", "TODO")
REQUIRED_PACKAGE_DOCS = (
    "01_技术交底书",
    "02_权利要求书",
    "03_说明书",
    "04_说明书摘要",
    "07_初步查新笔记",
    "08_发明专利请求书信息确认表",
    "提交文件清单",
)
DRAWING_PACKAGE_DOCS = ("05_说明书附图", "06_摘要附图")
PACKAGE_DOCS = (
    "00_专利草案总览",
    *REQUIRED_PACKAGE_DOCS,
    *DRAWING_PACKAGE_DOCS,
    "09_查新与差异化分析",
)


def normalize_case_name(name: str, max_len: int = 80) -> str:
    text = name.strip()
    for marker in PLACEHOLDERS:
        text = text.replace(marker, "")
    text = re.sub(FORBIDDEN, "", text)
    text = re.sub(r"\s+", "", text)
    text = text.strip("._- ")
    if not text:
        text = "中国发明专利"
    if len(text) > max_len:
        text = text[:max_len].rstrip("._- ")
    return text


def build_paths(case_name: str, case_dir: str, extensions: list[str]) -> dict[str, object]:
    safe_name = normalize_case_name(case_name)
    timestamp = datetime.now().astimezone().strftime("%Y%m%d%H%M%S")
    stem = f"{safe_name}_{timestamp}"
    base = Path(case_dir).expanduser()
    files = {ext: str(base / f"{stem}.{ext}") for ext in extensions}
    return {
        "case_name": case_name,
        "safe_case_name": safe_name,
        "timestamp": timestamp,
        "stem": stem,
        "case_dir": str(base),
        "files": files,
    }


def build_package(case_name: str, output_root: str, extensions: list[str]) -> dict[str, object]:
    safe_name = normalize_case_name(case_name)
    timestamp = datetime.now().astimezone().strftime("%Y%m%d%H%M%S")
    case_dir = Path(output_root).expanduser() / f"{safe_name}_{timestamp}"
    package_files = {
        doc: {ext: str(case_dir / f"{doc}.{ext}") for ext in extensions}
        for doc in PACKAGE_DOCS
    }
    return {
        "case_name": case_name,
        "safe_case_name": safe_name,
        "timestamp": timestamp,
        "case_dir": str(case_dir),
        "package_files": package_files,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case-name", required=True, help="Chinese case/invention name.")
    parser.add_argument("--case-dir", default="outputs/case", help="Output case directory.")
    parser.add_argument("--output-root", default="outputs", help="Package output root.")
    parser.add_argument(
        "--package",
        action="store_true",
        help="Generate a complete patent package path map.",
    )
    parser.add_argument(
        "--ext",
        action="append",
        default=["md"],
        help="Output extension without dot. Can be repeated. Default: md.",
    )
    parser.add_argument("--also-docx", action="store_true", help="Also include a docx path.")
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = parse_args()
    extensions = list(dict.fromkeys(ext.lstrip(".").lower() for ext in args.ext))
    if any(not re.fullmatch(r"[a-z0-9]+", ext) for ext in extensions):
        raise SystemExit("--ext 只能使用字母或数字组成的扩展名")
    if args.also_docx and "docx" not in extensions:
        extensions.append("docx")
    if args.package:
        result = build_package(args.case_name, args.output_root, extensions)
    else:
        result = build_paths(args.case_name, args.case_dir, extensions)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
