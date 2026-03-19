#!/usr/bin/env python3
"""Create project backups and append memory entries."""

from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path


def normalize(path: Path | None) -> str:
    if path is None:
        return "none"
    return path.resolve().as_posix()


def ensure_backup_dir(project_root: Path) -> Path:
    backup_dir = project_root / ".backup"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def create_backup(project_root: Path, target: Path) -> dict:
    backup_dir = ensure_backup_dir(project_root)
    target = target.resolve()
    if not target.exists():
        return {
            "created": False,
            "target": normalize(target),
            "backup_path": "none",
        }

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_name = f"{timestamp}-{target.name}"
    backup_path = backup_dir / backup_name
    shutil.copy2(target, backup_path)
    return {
        "created": True,
        "target": normalize(target),
        "backup_path": normalize(backup_path),
    }


def append_memory(
    project_root: Path,
    target: Path,
    backup: str,
    approved: str,
    changes: str,
    result: str,
    warnings: str,
) -> dict:
    backup_dir = ensure_backup_dir(project_root)
    memory_path = backup_dir / "memories.md"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    entry = "\n".join(
        [
            f"## {timestamp}",
            f"- target: {normalize(target.resolve())}",
            f"- backup: {backup or 'none'}",
            f"- approved: {approved or 'none'}",
            f"- changes: {changes or 'none'}",
            f"- result: {result or 'none'}",
            f"- warnings: {warnings or 'none'}",
            "",
        ]
    )
    with memory_path.open("a", encoding="utf-8") as handle:
        handle.write(entry)
    return {
        "memory_path": normalize(memory_path),
        "timestamp": timestamp,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create project backups and append memory entries.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    backup_parser = subparsers.add_parser("backup", help="Create a timestamped backup for an existing target file.")
    backup_parser.add_argument("--project-root", required=True, help="Project root that owns the .backup directory.")
    backup_parser.add_argument("--target", required=True, help="Target file path to back up if it exists.")

    remember_parser = subparsers.add_parser("remember", help="Append a memory entry to .backup/memories.md.")
    remember_parser.add_argument("--project-root", required=True, help="Project root that owns the .backup directory.")
    remember_parser.add_argument("--target", required=True, help="Edited or generated target file path.")
    remember_parser.add_argument("--backup", default="none", help="Backup file path, or 'none'.")
    remember_parser.add_argument("--approved", default="none", help="Approved change items.")
    remember_parser.add_argument("--changes", default="none", help="What was actually changed.")
    remember_parser.add_argument("--result", default="none", help="Result summary.")
    remember_parser.add_argument("--warnings", default="none", help="Warnings or verification gaps.")

    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(args.project_root).resolve()
    target = Path(args.target)

    if args.command == "backup":
        result = create_backup(project_root, target)
    else:
        result = append_memory(
            project_root=project_root,
            target=target,
            backup=args.backup,
            approved=args.approved,
            changes=args.changes,
            result=args.result,
            warnings=args.warnings,
        )

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
