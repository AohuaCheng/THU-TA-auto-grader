#!/usr/bin/env python3
"""Filter score summary rows to match official class roster (.xls)."""

from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

import openpyxl
import xlrd

from export_score_summary import COMBINED_OUTPUT_NAME, ESSAY2_STANDALONE_NAME
from paths import download_root

DEFAULT_ROSTER = Path(__file__).resolve().parent.parent.parent / "学生信息.xls"


def normalize_student_id(raw) -> str:
    if raw is None:
        return ""
    sid = str(raw).strip()
    if sid.endswith(".0") and sid[:-2].isdigit():
        sid = sid[:-2]
    return sid


def load_roster_ids(roster_path: Path) -> set[str]:
    book = xlrd.open_workbook(str(roster_path))
    sheet = book.sheet_by_index(0)
    ids: set[str] = set()
    for row in range(1, sheet.nrows):
        sid = normalize_student_id(sheet.cell_value(row, 0))
        if sid and re.fullmatch(r"\d{10}", sid):
            ids.add(sid)
    if not ids:
        raise ValueError(f"未从 roster 解析到学号: {roster_path}")
    return ids


def filter_sheet_rows(ws, roster_ids: set[str]) -> list[str]:
    """Delete data rows (row>=3) whose 学号 not in roster. Returns removed ids."""
    removed: list[str] = []
    for row_idx in range(ws.max_row, 2, -1):
        sid = normalize_student_id(ws.cell(row_idx, 1).value)
        if not sid:
            ws.delete_rows(row_idx, 1)
            continue
        if sid not in roster_ids:
            removed.append(sid)
            ws.delete_rows(row_idx, 1)
    return removed


def filter_score_workbook(
    score_xlsx: Path,
    roster_path: Path,
    *,
    sheets: tuple[str, ...] = ("Essay 1", "Essay 2"),
) -> dict[str, list[str]]:
    roster_ids = load_roster_ids(roster_path)
    wb = openpyxl.load_workbook(score_xlsx)
    report: dict[str, list[str]] = {}
    for name in sheets:
        if name not in wb.sheetnames:
            continue
        report[name] = filter_sheet_rows(wb[name], roster_ids)
    wb.save(score_xlsx)
    return report


def filter_combined_and_essay2_copy(
    *,
    download_root_path: Path,
    roster_path: Path,
    combined_name: str = COMBINED_OUTPUT_NAME,
    essay2_name: str = ESSAY2_STANDALONE_NAME,
) -> tuple[Path, dict[str, list[str]]]:
    combined = download_root_path / combined_name
    essay2_only = download_root_path / essay2_name
    if not combined.exists():
        raise FileNotFoundError(combined)
    report = filter_score_workbook(combined, roster_path)
    if essay2_only.exists():
        shutil.copy2(combined, essay2_only)
        wb = openpyxl.load_workbook(essay2_only)
        for sn in list(wb.sheetnames):
            if sn != "Essay 2":
                del wb[sn]
        wb.save(essay2_only)
    return combined, report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="按最终学生名单过滤分数汇总 xlsx")
    parser.add_argument(
        "--roster",
        type=Path,
        default=DEFAULT_ROSTER,
        help="学生信息.xls 路径",
    )
    parser.add_argument(
        "--scores",
        type=Path,
        help=f"汇总 xlsx（默认 downloads/{COMBINED_OUTPUT_NAME}）",
    )
    args = parser.parse_args(argv)

    root = download_root()
    scores = args.scores or (root / COMBINED_OUTPUT_NAME)
    roster = args.roster
    if not roster.exists():
        raise SystemExit(f"找不到名单: {roster}")
    combined, report = filter_combined_and_essay2_copy(
        download_root_path=root,
        roster_path=roster,
    )
    roster_n = len(load_roster_ids(roster))
    print(f"名单学号: {roster_n} 人")
    print(f"已更新: {combined}")
    print(f"已同步: {root / ESSAY2_STANDALONE_NAME}")
    for sheet, removed in report.items():
        if removed:
            print(f"  {sheet} 删除 {len(removed)} 人: {', '.join(sorted(removed))}")
        else:
            print(f"  {sheet} 无多余行")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
