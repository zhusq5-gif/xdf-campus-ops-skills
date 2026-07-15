#!/usr/bin/env python3
"""Fail when repository files contain likely secrets or direct identifiers."""

from __future__ import annotations

import os
import re
import sys
import zipfile
from pathlib import Path


PATTERNS = {
    "中国大陆手机号": re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"),
    "身份证号": re.compile(r"(?<!\d)\d{17}[0-9Xx](?!\d)"),
    "疑似密钥": re.compile(r"(?i)(api[_-]?key|access[_-]?token|secret)\s*[:=]\s*['\"][A-Za-z0-9_\-]{12,}"),
    "机器人 Webhook": re.compile(r"(?i)https://[^\s'\"]+/(?:hook|webhook)/[A-Za-z0-9_\-]{8,}"),
}
IGNORED_DIRS = {".git", "data", "outputs", "node_modules", "__pycache__", ".venv"}
TEXT_SUFFIXES = {".md", ".txt", ".json", ".yaml", ".yml", ".py", ".mjs", ".js", ".toml", ".ini", ".cfg"}


def text_chunks(path: Path):
    if path.suffix.lower() == ".xlsx":
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    yield f"{path}:{name}", archive.read(name).decode("utf-8", errors="ignore")
    elif path.suffix.lower() in TEXT_SUFFIXES or path.name in {".gitignore"}:
        yield str(path), path.read_text(encoding="utf-8", errors="ignore")


def scan(root: Path) -> list[str]:
    findings: list[str] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [name for name in dirnames if name not in IGNORED_DIRS]
        for filename in filenames:
            path = Path(dirpath) / filename
            for label, text in text_chunks(path):
                for pattern_name, pattern in PATTERNS.items():
                    match = pattern.search(text)
                    if match:
                        findings.append(f"{label}: {pattern_name}（位置 {match.start()}）")
    return findings


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".").resolve()
    findings = scan(root)
    if findings:
        print("敏感信息扫描失败：")
        for finding in findings:
            print(f"- {finding}")
        return 1
    print("Sensitive data scan passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
