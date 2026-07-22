#!/usr/bin/env python3
"""Manage local, reversible rule overlays without editing the initial config."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from normalize_schedule import (
    base_config_sha256,
    deep_merge,
    flatten_config,
    validate_config,
    validate_rule_overrides,
)


OVERLAY_SCHEMA_VERSION = "1.0.0"


def read_base(path: Path) -> dict[str, Any]:
    config = json.loads(path.read_text(encoding="utf-8"))
    return validate_config(config)


def read_overlay(path: Path, base_hash: str) -> dict[str, Any]:
    if not path.exists():
        return {
            "schema_version": OVERLAY_SCHEMA_VERSION,
            "base_config_sha256": base_hash,
            "overrides": {},
            "history": [],
        }
    overlay = json.loads(path.read_text(encoding="utf-8"))
    if overlay.get("schema_version") != OVERLAY_SCHEMA_VERSION:
        raise ValueError("规则覆盖文件 schema_version 必须为 1.0.0")
    if overlay.get("base_config_sha256") != base_hash:
        raise ValueError("初始规则已变化，请先审核旧覆盖，不得静默迁移")
    validate_rule_overrides(overlay.get("overrides", {}))
    if not isinstance(overlay.get("history", []), list):
        raise ValueError("规则覆盖 history 必须是列表")
    return overlay


def write_overlay(path: Path, overlay: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(overlay, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def apply_patch(base_path: Path, overlay_path: Path, patch_path: Path, request: str) -> dict[str, Any]:
    base = read_base(base_path)
    base_hash = base_config_sha256(base_path)
    overlay = read_overlay(overlay_path, base_hash)
    patch = validate_rule_overrides(json.loads(patch_path.read_text(encoding="utf-8")))
    overrides = deep_merge(overlay["overrides"], patch)
    validate_config(deep_merge(base, overrides))
    overlay["overrides"] = overrides
    overlay["history"].append({"action": "apply", "request": request.strip(), "patch": patch})
    write_overlay(overlay_path, overlay)
    return overlay


def restore_initial(base_path: Path, overlay_path: Path, request: str) -> dict[str, Any]:
    read_base(base_path)
    base_hash = base_config_sha256(base_path)
    overlay = read_overlay(overlay_path, base_hash)
    overlay["overrides"] = {}
    overlay["history"].append({"action": "restore_initial", "request": request.strip()})
    write_overlay(overlay_path, overlay)
    return overlay


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="管理教师课表 Skill 的本地可回退规则")
    parser.add_argument("--base", type=Path, required=True, help="只读的初始规则 JSON")
    parser.add_argument("--overlay", type=Path, required=True, help="位于 data/local/ 的本地覆盖 JSON")
    subparsers = parser.add_subparsers(dest="command", required=True)
    apply_parser = subparsers.add_parser("apply", help="应用 AI 已转换并经使用者确认的 JSON 规则补丁")
    apply_parser.add_argument("--patch", type=Path, required=True, help="只包含允许规则路径的 JSON 补丁")
    apply_parser.add_argument("--request", required=True, help="使用者原始要求，仅保留在 Git 忽略的本地覆盖历史")
    restore_parser = subparsers.add_parser("restore", help="清空本地覆盖，回到初始规则")
    restore_parser.add_argument("--request", default="恢复初始规则")
    subparsers.add_parser("show", help="只读查看当前本地覆盖摘要")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        base_hash = base_config_sha256(args.base)
        if args.command == "apply":
            overlay = apply_patch(args.base, args.overlay, args.patch, args.request)
        elif args.command == "restore":
            overlay = restore_initial(args.base, args.overlay, args.request)
        else:
            read_base(args.base)
            overlay = read_overlay(args.overlay, base_hash)
        result = {
            "status": "ok",
            "mode": "local_override" if overlay["overrides"] else "initial",
            "base_config_sha256": base_hash,
            "override_count": len(flatten_config(overlay["overrides"])),
            "history_count": len(overlay["history"]),
        }
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
