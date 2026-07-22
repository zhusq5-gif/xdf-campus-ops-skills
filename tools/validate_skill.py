#!/usr/bin/env python3
"""Validate the repository skill without depending on a Codex installation."""

from __future__ import annotations

import re
import sys
from pathlib import Path


NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")

REQUIRED_BY_SKILL = {
    "xdf-plan-campus-capacity": [
        "scripts/plan_capacity.py",
        "references/data-contract.md",
        "references/business-rules.md",
        "references/management-output.md",
        "assets/capacity-input-template.xlsx",
        "assets/capacity-output-template.xlsx",
    ],
    "xdf-normalize-teacher-schedule": [
        "scripts/normalize_schedule.py",
        "scripts/manage_schedule_rules.py",
        "references/data-contract.md",
        "references/business-rules.md",
        "references/output-spec.md",
    ],
}


def parse_frontmatter(text: str) -> dict[str, str]:
    if not text.startswith("---\n"):
        raise ValueError("SKILL.md 必须以 YAML frontmatter 开始")
    try:
        block = text.split("---\n", 2)[1]
    except IndexError as exc:
        raise ValueError("SKILL.md frontmatter 未闭合") from exc
    values: dict[str, str] = {}
    for line in block.splitlines():
        if not line.strip():
            continue
        if ":" not in line:
            raise ValueError(f"无效 frontmatter 行: {line}")
        key, value = line.split(":", 1)
        values[key.strip()] = value.strip().strip('"')
    return values


def validate(skill_dir: Path) -> list[str]:
    errors: list[str] = []
    skill_file = skill_dir / "SKILL.md"
    if not skill_file.is_file():
        return ["缺少 SKILL.md"]
    text = skill_file.read_text(encoding="utf-8")
    try:
        metadata = parse_frontmatter(text)
    except ValueError as exc:
        return [str(exc)]
    if set(metadata) != {"name", "description"}:
        errors.append("frontmatter 只能包含 name 和 description")
    name = metadata.get("name", "")
    if not NAME_RE.fullmatch(name):
        errors.append("name 必须为小写 kebab-case")
    if name != skill_dir.name:
        errors.append("name 必须与目录名一致")
    description = metadata.get("description", "")
    if not description or len(description) > 1024:
        errors.append("description 必须为 1-1024 字符")
    if "TODO" in text:
        errors.append("Skill 中不得保留 TODO")
    if len(text.splitlines()) >= 500:
        errors.append("SKILL.md 必须少于 500 行")
    required = [skill_dir / "agents" / "openai.yaml"]
    required.extend(skill_dir / relative for relative in REQUIRED_BY_SKILL.get(name, []))
    for path in required:
        if not path.is_file():
            errors.append(f"缺少资源: {path.relative_to(skill_dir)}")
    openai_yaml = skill_dir / "agents" / "openai.yaml"
    if openai_yaml.is_file():
        yaml_text = openai_yaml.read_text(encoding="utf-8")
        for field in ("display_name:", "short_description:", "default_prompt:"):
            if field not in yaml_text:
                errors.append(f"agents/openai.yaml 缺少 {field[:-1]}")
        if f"${name}" not in yaml_text:
            errors.append("default_prompt 必须显式触发当前 Skill")
    return errors


def main() -> int:
    if len(sys.argv) != 2:
        print("用法: validate_skill.py <skill-dir>", file=sys.stderr)
        return 2
    errors = validate(Path(sys.argv[1]))
    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1
    print("Skill validation passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
