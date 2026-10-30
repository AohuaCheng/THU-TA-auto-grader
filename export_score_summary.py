#!/usr/bin/env python3
"""Export essay scores to xlsx summary (template-based, multi-sheet)."""

from __future__ import annotations

import argparse
import ast
import re
import shutil
from pathlib import Path

import openpyxl
from docx import Document

from paths import download_root, templates_dir
from topic_detector import describe_topic, essay_number_from_dir
from yuketang_grader import collect_docx_jobs, feedback_path_for

TEMPLATE_NAME = "MINC 2025 Essay 分数汇总 样例.xlsx"
DEFAULT_OUTPUT_NAME = "Essay1 分数汇总.xlsx"
COMBINED_OUTPUT_NAME = "Minc2026Essays分数汇总.xlsx"
ESSAY2_STANDALONE_NAME = "Minc2026Essays_Essay2分数汇总.xlsx"
FIRST_ROUND_END_COL = 15
COL_EXCELLENT = 17
COL_WORD_COUNT = 18
COL_REMARK = 19

SCORE_FIELDS = (
    ("Ideational Flexibility", "diversity"),
    ("Originality", "originality"),
    ("Relevance", "relevance"),
    ("Evaluation", "evaluation"),
    ("Analysis", "analysis"),
    ("Inference", "inference"),
    ("Explanation", "explanation"),
    ("Interpretation", "interpretation"),
    ("Factuality", "factuality"),
    ("Clarity", "clarity"),
)


def template_path() -> Path:
    path = templates_dir() / TEMPLATE_NAME
    if not path.exists():
        raise FileNotFoundError(f"找不到模板: {path}")
    return path


def sheet_name_for_dir(download_dir: Path) -> str:
    n = essay_number_from_dir(download_dir)
    return f"Essay {n}"


def topic_label(docx_path: Path, download_dir: Path | None = None) -> str:
    _, label, _ = describe_topic(docx_path, download_dir)
    return label


def extract_scores(feedback_path: Path) -> dict[str, float]:
    doc = Document(str(feedback_path))
    for paragraph in reversed(doc.paragraphs):
        text = paragraph.text
        if "Overall" in text and "Ideational Flexibility" in text:
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                data = ast.literal_eval(match.group())
                return {str(k): float(v) for k, v in data.items()}
    raise ValueError(f"未找到分数: {feedback_path}")


def setup_first_round_headers(ws) -> None:
    for merged in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(merged))
        except KeyError:
            pass

    if ws.max_column > FIRST_ROUND_END_COL:
        ws.delete_cols(FIRST_ROUND_END_COL + 1, ws.max_column - FIRST_ROUND_END_COL)

    if ws.max_row > 2:
        ws.delete_rows(3, ws.max_row - 2)

    if ws.max_column < 3 or ws.cell(1, 3).value != "Topic":
        ws.insert_cols(3)

    ws.cell(row=1, column=1, value="学号")
    ws.cell(row=1, column=2, value="姓名")
    ws.cell(row=1, column=3, value="Topic")
    ws.cell(row=1, column=7, value="Sum")
    ws.cell(row=1, column=13, value="Sum")
    ws.merge_cells("D1:F1")
    ws["D1"] = "Creativity"
    ws.merge_cells("H1:L1")
    ws["H1"] = "Critical Thinking"

    ws.cell(row=2, column=4, value="Diversity")
    ws.cell(row=2, column=5, value="Originality")
    ws.cell(row=2, column=6, value="Relevance")
    ws.cell(row=2, column=7, value="Out of 45")
    ws.cell(row=2, column=8, value="Evaluation")
    ws.cell(row=2, column=9, value="Analysis ")
    ws.cell(row=2, column=10, value="Inference")
    ws.cell(row=2, column=11, value="Explanation")
    ws.cell(row=2, column=12, value="Interpretation")
    ws.cell(row=2, column=13, value="Out of 45")
    ws.cell(row=2, column=14, value="Factuality")
    ws.cell(row=2, column=15, value="Clarity")
    ws.cell(row=2, column=16, value="Overall")
    ws.cell(row=1, column=COL_EXCELLENT, value="优秀作业")
    ws.cell(row=1, column=COL_WORD_COUNT, value="字数")
    ws.cell(row=1, column=COL_REMARK, value="备注")


def read_extra_columns(ws) -> dict[str, dict[int, object]]:
    """By student_id -> {col_index: value} for cols 17-19."""
    extras: dict[str, dict[int, object]] = {}
    for row_idx in range(3, ws.max_row + 1):
        sid = ws.cell(row_idx, 1).value
        if not sid:
            continue
        sid = str(sid).strip()
        extras[sid] = {
            COL_EXCELLENT: ws.cell(row_idx, COL_EXCELLENT).value,
            COL_WORD_COUNT: ws.cell(row_idx, COL_WORD_COUNT).value,
            COL_REMARK: ws.cell(row_idx, COL_REMARK).value,
        }
    return extras


def write_row(
    ws,
    row_idx: int,
    student_id: str,
    name: str,
    topic: str,
    scores: dict[str, float],
    extras: dict[int, object] | None = None,
) -> None:
    ws.cell(row=row_idx, column=1, value=student_id)
    ws.cell(row=row_idx, column=2, value=name)
    ws.cell(row=row_idx, column=3, value=topic)

    ws.cell(row=row_idx, column=4, value=scores.get("Ideational Flexibility"))
    ws.cell(row=row_idx, column=5, value=scores.get("Originality"))
    ws.cell(row=row_idx, column=6, value=scores.get("Relevance"))
    ws.cell(row=row_idx, column=7, value=f"=SUM(D{row_idx}:F{row_idx})")

    ws.cell(row=row_idx, column=8, value=scores.get("Evaluation"))
    ws.cell(row=row_idx, column=9, value=scores.get("Analysis"))
    ws.cell(row=row_idx, column=10, value=scores.get("Inference"))
    ws.cell(row=row_idx, column=11, value=scores.get("Explanation"))
    ws.cell(row=row_idx, column=12, value=scores.get("Interpretation"))
    ws.cell(row=row_idx, column=13, value=f"=SUM(H{row_idx}:L{row_idx})")

    ws.cell(row=row_idx, column=14, value=scores.get("Factuality"))
    ws.cell(row=row_idx, column=15, value=scores.get("Clarity"))
    ws.cell(row=row_idx, column=16, value=scores.get("Overall"))

    if extras:
        for col, val in extras.items():
            ws.cell(row=row_idx, column=col, value=val)


def _ensure_sheet_from_template(
    wb: openpyxl.Workbook, tpl_wb: openpyxl.Workbook, sheet_name: str
) -> None:
    if sheet_name in wb.sheetnames:
        return
    if sheet_name not in tpl_wb.sheetnames:
        raise ValueError(f"模板缺少 sheet: {sheet_name}")
    src = tpl_wb[sheet_name]
    dst = wb.create_sheet(sheet_name)
    for row in src.iter_rows():
        for cell in row:
            dst.cell(cell.row, cell.column, value=cell.value)
    for merged in src.merged_cells.ranges:
        dst.merge_cells(str(merged))


def load_or_create_workbook(output_path: Path, template: Path | None = None) -> openpyxl.Workbook:
    src = template or template_path()
    if output_path.exists():
        return openpyxl.load_workbook(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, output_path)
    return openpyxl.load_workbook(output_path)


def export_scores_to_sheet(
    wb: openpyxl.Workbook,
    download_dir: Path,
    sheet_name: str,
    *,
    preserve_extras: dict[str, dict[int, object]] | None = None,
) -> int:
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"工作簿中无 sheet: {sheet_name}，现有: {wb.sheetnames}")
    ws = wb[sheet_name]
    setup_first_round_headers(ws)

    jobs = sorted(collect_docx_jobs(download_dir), key=lambda j: j.student_id)
    if not jobs:
        raise ValueError(f"未找到学生作业: {download_dir}")

    row_idx = 3
    for job in jobs:
        feedback = feedback_path_for(job)
        if not feedback.is_file():
            print(f"  跳过（无 feedback）: {job.student_id}_{job.name}")
            continue
        scores = extract_scores(feedback)
        label = topic_label(job.docx_path, download_dir)
        extras = (preserve_extras or {}).get(job.student_id)
        write_row(ws, row_idx, job.student_id, job.name, label, scores, extras=extras)
        row_idx += 1

    return row_idx - 3


def export_score_summary(
    download_dir: Path,
    *,
    template: Path | None = None,
    output_path: Path | None = None,
    sheet_name: str | None = None,
    preserve_extras_from: Path | None = None,
) -> Path:
    sheet = sheet_name or sheet_name_for_dir(download_dir)
    target = output_path or (download_dir / DEFAULT_OUTPUT_NAME)

    extras: dict[str, dict[int, object]] | None = None
    if preserve_extras_from and preserve_extras_from.exists():
        old_wb = openpyxl.load_workbook(preserve_extras_from)
        if sheet in old_wb.sheetnames:
            extras = read_extra_columns(old_wb[sheet])

    if target.exists() and target.resolve() != preserve_extras_from:
        wb = openpyxl.load_workbook(target)
    else:
        wb = load_or_create_workbook(target, template=template)

    export_scores_to_sheet(wb, download_dir, sheet, preserve_extras=extras)
    wb.save(target)
    return target


def sync_combined_summary(
    *,
    download_root_path: Path,
    combined_path: Path | None = None,
    essay2_standalone_path: Path | None = None,
) -> tuple[Path, Path]:
    """Refresh Essay 1 & 2 in combined xlsx; also write Essay2 standalone copy."""
    combined = combined_path or (download_root_path / COMBINED_OUTPUT_NAME)
    essay2_only = essay2_standalone_path or (download_root_path / ESSAY2_STANDALONE_NAME)
    essay1_dir = download_root_path / "Essay1"
    essay2_dir = download_root_path / "Essay2"

    if not essay1_dir.is_dir() or not essay2_dir.is_dir():
        raise FileNotFoundError("需要 downloads/Essay1 与 downloads/Essay2")

    tpl = template_path()
    had_combined = combined.exists()
    if had_combined:
        wb = openpyxl.load_workbook(combined)
    else:
        shutil.copy2(tpl, combined)
        wb = openpyxl.load_workbook(combined)

    tpl_wb = openpyxl.load_workbook(tpl)
    for name in ("Essay 1", "Essay 2"):
        _ensure_sheet_from_template(wb, tpl_wb, name)

    extras_e1 = read_extra_columns(wb["Essay 1"]) if had_combined else {}
    if had_combined:
        print("保留已有 Essay 1 sheet（分数/优秀作业/字数/备注），仅刷新 Essay 2。")
    else:
        export_scores_to_sheet(wb, essay1_dir, "Essay 1", preserve_extras=extras_e1)
    export_scores_to_sheet(wb, essay2_dir, "Essay 2", preserve_extras=None)
    wb.save(combined)

    shutil.copy2(combined, essay2_only)
    wb2 = openpyxl.load_workbook(combined)
    for sn in list(wb2.sheetnames):
        if sn != "Essay 2":
            del wb2[sn]
    wb2.save(essay2_only)

    return combined, essay2_only


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="导出 Essay 分数汇总 xlsx")
    parser.add_argument("--dir", help="作业目录（默认 downloads 下唯一子目录）")
    parser.add_argument("--template", help="分数汇总模板 xlsx")
    parser.add_argument("--output", help="输出文件路径")
    parser.add_argument(
        "--combined",
        action="store_true",
        help="更新 downloads/Minc2026Essays分数汇总.xlsx（Essay1+Essay2）并生成 Essay2 独立副本",
    )
    args = parser.parse_args(argv)

    root = download_root()

    if args.combined:
        combined, essay2_only = sync_combined_summary(download_root_path=root)
        print(f"已更新汇总: {combined}")
        print(f"已保留 Essay2 独立文件: {essay2_only}")
        return 0

    if args.dir:
        download_dir = Path(args.dir)
    else:
        subdirs = sorted(p for p in root.iterdir() if p.is_dir())
        download_dir = subdirs[0] if len(subdirs) == 1 else root

    output = Path(args.output) if args.output else None
    template = Path(args.template) if args.template else None
    preserve = None
    combined_default = root / COMBINED_OUTPUT_NAME
    if combined_default.exists() and sheet_name_for_dir(download_dir) == "Essay 1":
        preserve = combined_default
    path = export_score_summary(
        download_dir,
        template=template,
        output_path=output,
        preserve_extras_from=preserve,
    )
    print(f"已生成: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
