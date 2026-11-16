#!/usr/bin/env python3
# Adapted from HuangXinzhe/cn-patent-drafting. See its MIT notice in ../licenses/.
"""Append a Markdown revision-log entry for a patent case."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path


KIND_LABELS = {
    "merge": "合并迭代",
    "correct": "纠正迭代",
    "rewrite": "重写迭代",
    "search": "查新补充",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case-dir", required=True, help="Directory containing case outputs.")
    parser.add_argument(
        "--kind",
        required=True,
        choices=sorted(KIND_LABELS),
        help="Revision type.",
    )
    parser.add_argument("--user", required=True, help="Short summary of the user request.")
    parser.add_argument("--summary", required=True, help="Short summary of the change.")
    parser.add_argument(
        "--artifacts",
        default="",
        help="Comma-separated output artifact filenames or paths.",
    )
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args = parse_args()
    case_dir = Path(args.case_dir).expanduser()
    case_dir.mkdir(parents=True, exist_ok=True)
    log_path = case_dir / "交底书修订对话记录.md"
    if log_path.is_symlink() or not log_path.resolve().is_relative_to(case_dir.resolve()):
        raise SystemExit("修订记录必须保存在案件目录内")

    local_now = datetime.now().astimezone()
    utc_now = datetime.now(timezone.utc)
    artifacts = [item.strip() for item in args.artifacts.split(",") if item.strip()]
    artifact_text = "\n".join(f"- {item}" for item in artifacts) if artifacts else "- 未记录"

    entry = (
        f"## {local_now.strftime('%Y-%m-%d %H:%M:%S %z')}\n\n"
        f"- UTC: {utc_now.strftime('%Y-%m-%d %H:%M:%S %z')}\n"
        f"- 类型: {KIND_LABELS[args.kind]}\n"
        f"- 用户说明摘要: {args.user}\n"
        f"- 本轮交付文件:\n{artifact_text}\n"
        f"- 合并/纠正摘要摘录: {args.summary}\n\n"
    )
    with log_path.open("a", encoding="utf-8") as fh:
        if fh.tell() == 0:
            fh.write("# 交底书修订对话记录\n\n")
        fh.write(entry)

    print(str(log_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
