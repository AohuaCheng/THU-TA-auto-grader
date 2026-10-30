#!/usr/bin/env python3
"""Build MINC total grade workbook (Essay scores + course total template)."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import openpyxl
from export_score_summary import COMBINED_OUTPUT_NAME
from paths import download_root, templates_dir
from roster_filter import DEFAULT_ROSTER, load_roster_ids, normalize_student_id

TOTAL_TEMPLATE_NAME = "MINC 2025 总成绩 样例.xlsx"
DEFAULT_OUTPUT_NAME = "Minc2026总成绩.xlsx"
SCORES_SOURCE_NAME = COMBINED_OUTPUT_NAME
COL_OVERALL = 16

ESSAY_SCORES_HEADERS = (
    "学号",
    "姓名",
    "Essay 1 original",
    "Essay 1 resubmission",
    "Essay 1 final score",
    "Essay 2 original",
    "Essay 2 resubmission",
    "Essay 2 final score",
)


def load_roster_rows(roster_path: Path) -> list[tuple[str, str]]:
    import xlrd

    book = xlrd.open_workbook(str(roster_path))
    sheet = book.sheet_by_index(0)
    rows: list[tuple[str, str]] = []
    for r in range(1, sheet.nrows):
        sid = normalize_student_id(sheet.cell_value(r, 0))
        name = str(sheet.cell_value(r, 1)).strip()
        if sid:
            rows.append((sid, name))
    return rows


def read_overall_by_sheet(scores_xlsx: Path) -> dict[str, dict[str, float | None]]:
    """sheet_name -> {student_id: overall}"""
    wb = openpyxl.load_workbook(scores_xlsx, data_only=True)
    out: dict[str, dict[str, float | None]] = {}
    for sheet_name in ("Essay 1", "Essay 2"):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        mapping: dict[str, float | None] = {}
        for r in range(3, ws.max_row + 1):
            sid = normalize_student_id(ws.cell(r, 1).value)
            if not sid:
                continue
            val = ws.cell(r, COL_OVERALL).value
            if isinstance(val, (int, float)):
                mapping[sid] = float(val)
            else:
                mapping[sid] = None
        out[sheet_name] = mapping
    return out


def _final(original: float | None, resub: float | None) -> float | None:
    if resub is not None:
        return resub
    return original


def build_essay_scores_sheet(ws, roster: list[tuple[str, str]], scores: dict[str, dict[str, float | None]]) -> None:
    e1 = scores.get("Essay 1", {})
    e2 = scores.get("Essay 2", {})
    for col, header in enumerate(ESSAY_SCORES_HEADERS, start=1):
        ws.cell(row=1, column=col, value=header)

    for idx, (sid, name) in enumerate(roster, start=2):
        o1 = e1.get(sid)
        o2 = e2.get(sid)
        r1 = None  # resubmission placeholders
        r2 = None
        f1 = _final(o1, r1)
        f2 = _final(o2, r2)
        ws.cell(row=idx, column=1, value=sid)
        ws.cell(row=idx, column=2, value=name)
        ws.cell(row=idx, column=3, value=o1)
        ws.cell(row=idx, column=4, value=r1)
        ws.cell(row=idx, column=5, value=f1)
        ws.cell(row=idx, column=6, value=o2)
        ws.cell(row=idx, column=7, value=r2)
        ws.cell(row=idx, column=8, value=f2)


def prepare_total_grade_sheet(ws) -> None:
    """Adapt MINC 2025 template to Essay 1–2 only (drop Essay 3–6 columns)."""
    # Delete Essay 3–6 columns (F–I) after headers set
    ws.delete_cols(6, 4)  # old F-I Essay 3-6
    ws.cell(row=1, column=4, value="Essay (60%)")
    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=5)
    headers = [
        "姓名",
        "学号",
        "Attendance (10%)",
        "Essay 1",
        "Essay 2",
        "Total",
        "Lecture Notes (5% Each)",
        "Presentation (10% Each)",
        "In Class Surveys (10%)",
        "Notes",
        "Final Score",
    ]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=2, column=col, value=h)


def build_total_grade_sheet(
    ws,
    roster: list[tuple[str, str]],
    essay_finals: dict[str, tuple[float | None, float | None]],
) -> None:
    prepare_total_grade_sheet(ws)
    for idx, (sid, name) in enumerate(roster, start=3):
        e1, e2 = essay_finals.get(sid, (None, None))
        ws.cell(row=idx, column=1, value=name)
        ws.cell(row=idx, column=2, value=sid)
        ws.cell(row=idx, column=3, value=None)  # attendance TBD
        ws.cell(row=idx, column=4, value=e1)
        ws.cell(row=idx, column=5, value=e2)
        ws.cell(row=idx, column=6, value=f"=AVERAGE(D{idx}:E{idx})*0.6")
        ws.cell(row=idx, column=7, value=None)
        ws.cell(row=idx, column=8, value=None)
        ws.cell(row=idx, column=9, value=None)
        ws.cell(row=idx, column=10, value=None)
        ws.cell(row=idx, column=11, value=f"=SUM(G{idx}:I{idx},F{idx},C{idx})")


def export_total_grades(
    *,
    roster_path: Path,
    scores_xlsx: Path,
    output_path: Path,
    template_path: Path | None = None,
) -> Path:
    roster = load_roster_rows(roster_path)
    roster_ids = {sid for sid, _ in roster}
    if roster_ids != set(load_roster_ids(roster_path)):
        pass

    by_sheet = read_overall_by_sheet(scores_xlsx)
    essay_finals: dict[str, tuple[float | None, float | None]] = {}
    for sid, _ in roster:
        o1 = by_sheet.get("Essay 1", {}).get(sid)
        o2 = by_sheet.get("Essay 2", {}).get(sid)
        essay_finals[sid] = (_final(o1, None), _final(o2, None))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    tpl = template_path or (templates_dir() / TOTAL_TEMPLATE_NAME)
    if tpl.exists():
        shutil.copy2(tpl, output_path)
        wb = openpyxl.load_workbook(output_path)
        old = wb.active.title
        ws_total = wb[old]
        ws_total.title = "总成绩"
    else:
        wb = openpyxl.Workbook()
        ws_total = wb.active
        ws_total.title = "总成绩"

    if "Essay scores" in wb.sheetnames:
        del wb["Essay scores"]
    ws_scores = wb.create_sheet("Essay scores", 0)
    build_essay_scores_sheet(ws_scores, roster, by_sheet)

    # Clear sample rows on total sheet then rebuild
    if ws_total.max_row > 2:
        ws_total.delete_rows(3, ws_total.max_row - 2)
    build_total_grade_sheet(ws_total, roster, essay_finals)

    wb.save(output_path)
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 MINC 总成绩汇总 xlsx")
    parser.add_argument("--roster", type=Path, default=DEFAULT_ROSTER)
    parser.add_argument(
        "--scores",
        type=Path,
        help=f"Essay 分数汇总（默认 downloads/{SCORES_SOURCE_NAME}）",
    )
    parser.add_argument("--output", type=Path, help=f"输出（默认 downloads/{DEFAULT_OUTPUT_NAME}）")
    parser.add_argument("--template", type=Path, help="总成绩模板 xlsx")
    args = parser.parse_args(argv)

    root = download_root()
    scores = args.scores or (root / SCORES_SOURCE_NAME)
    output = args.output or (root / DEFAULT_OUTPUT_NAME)
    if not args.roster.exists():
        raise SystemExit(f"找不到名单: {args.roster}")
    if not scores.exists():
        raise SystemExit(f"找不到: {scores}")

    path = export_total_grades(
        roster_path=args.roster,
        scores_xlsx=scores,
        output_path=output,
        template_path=args.template,
    )
    print(f"已生成: {path}")
    print("  - Sheet「Essay scores」: 第 1/2 次 Essay 原始分（重交列留空，final=original）")
    print("  - Sheet「总成绩」: Essay1/2 + 公式（Attendance/Profile 需手工填写）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
