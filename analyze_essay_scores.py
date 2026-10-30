#!/usr/bin/env python3
"""Topic and longitudinal score analysis for MINC essay summaries."""

from __future__ import annotations

import argparse
import math
from pathlib import Path

import openpyxl

from paths import download_root

COMBINED_NAME = "Minc2026Essays分数汇总.xlsx"
REPORT_NAME = "Essay_score_analysis.md"
COL_OVERALL = 16
COL_TOPIC = 3


def _overall(ws, row: int) -> float | None:
    val = ws.cell(row, COL_OVERALL).value
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None


def read_sheet_scores(path: Path, sheet: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows: list[dict] = []
    for r in range(3, ws.max_row + 1):
        sid = ws.cell(r, 1).value
        if not sid:
            continue
        overall = _overall(ws, r)
        if overall is None:
            continue
        rows.append(
            {
                "student_id": str(sid).strip(),
                "name": str(ws.cell(r, 2).value or "").strip(),
                "topic": str(ws.cell(r, COL_TOPIC).value or "").strip(),
                "overall": overall,
            }
        )
    return rows


def mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else float("nan")


def stdev(xs: list[float]) -> float:
    if len(xs) < 2:
        return float("nan")
    m = mean(xs)
    var = sum((x - m) ** 2 for x in xs) / (len(xs) - 1)
    return math.sqrt(var)


def welch_ttest(a: list[float], b: list[float]) -> tuple[float, float]:
    """Return (t statistic, approximate two-sided p-value)."""
    if len(a) < 2 or len(b) < 2:
        return float("nan"), float("nan")
    ma, mb = mean(a), mean(b)
    va, vb = stdev(a) ** 2, stdev(b) ** 2
    na, nb = len(a), len(b)
    se = math.sqrt(va / na + vb / nb)
    if se == 0:
        return 0.0, 1.0
    t = (ma - mb) / se
    # Welch–Satterthwaite df
    num = (va / na + vb / nb) ** 2
    den = (va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1)
    df = num / den if den else 1
    # Normal approx for large df; rough for small n
    p = 2 * (1 - _normal_cdf(abs(t)))
    return t, p


def _normal_cdf(x: float) -> float:
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def analyze_by_topic(rows: list[dict], essay_label: str) -> str:
    by_topic: dict[str, list[float]] = {}
    for row in rows:
        by_topic.setdefault(row["topic"], []).append(row["overall"])

    topics = sorted(by_topic.keys())
    lines = [f"### {essay_label}：题目间 Overall 比较", ""]
    if len(topics) < 2:
        lines.append("（仅一个题目或 Topic 缺失，跳过）")
        return "\n".join(lines)

    t1, t2 = topics[0], topics[1]
    a, b = by_topic[t1], by_topic[t2]
    ma, mb = mean(a), mean(b)
    t_stat, p_val = welch_ttest(a, b)
    diff = ma - mb

    lines.extend(
        [
            f"| 题目 (Topic) | n | 平均分 | 标准差 |",
            f"|---|---:|---:|---:|",
        ]
    )
    for topic in topics:
        xs = by_topic[topic]
        lines.append(
            f"| {topic} | {len(xs)} | {mean(xs):.2f} | {stdev(xs):.2f} |"
        )
    lines.extend(
        [
            "",
            f"- **差值**（{t1} − {t2}）: **{diff:+.2f}** 分",
            f"- Welch t ≈ {t_stat:.3f}，近似双侧 p ≈ {p_val:.4f} "
            f"（p<0.05 时可视为题目间均分差异较明显；样本量有限，宜结合效应量解读）",
            "",
        ]
    )
    return "\n".join(lines)


def analyze_longitudinal(
    e1: list[dict], e2: list[dict], *, big_drop: float = 10.0
) -> str:
    m1 = {r["student_id"]: r for r in e1}
    m2 = {r["student_id"]: r for r in e2}
    common = sorted(set(m1) & set(m2))
    deltas: list[tuple[str, str, float, float, float]] = []
    for sid in common:
        o1, o2 = m1[sid]["overall"], m2[sid]["overall"]
        deltas.append((sid, m1[sid]["name"], o1, o2, o2 - o1))

    improved = [d for d in deltas if d[4] > 0]
    declined = [d for d in deltas if d[4] < 0]
    same = [d for d in deltas if d[4] == 0]
    big_declines = [d for d in deltas if d[4] <= -big_drop]

    paired_diffs = [d[4] for d in deltas]
    md = mean(paired_diffs)
    lines = [
        "## Essay 1 → Essay 2 同一学生 Overall 变化",
        "",
        f"- 可配对人数: **{len(common)}**",
        f"- 提升（Essay2 > Essay1）: **{len(improved)}** 人 ({100*len(improved)/len(common):.1f}%)",
        f"- 下降: **{len(declined)}** 人 ({100*len(declined)/len(common):.1f}%)",
        f"- 不变: **{len(same)}** 人",
        f"- 平均变化（Essay2 − Essay1）: **{md:+.2f}** 分",
        f"- 明显下降（≤ −{big_drop:.0f} 分）: **{len(big_declines)}** 人",
        "",
    ]
    if len(paired_diffs) >= 2:
        md = mean(paired_diffs)
        sd = stdev(paired_diffs)
        se = sd / math.sqrt(len(paired_diffs))
        t_stat = md / se if se else 0.0
        p_val = 2 * (1 - _normal_cdf(abs(t_stat)))
        lines.append(
            f"- 配对 t（Essay2−Essay1 差值均分是否为 0）: t ≈ {t_stat:.3f}，"
            f"近似 p ≈ {p_val:.4f}"
        )
        lines.append("")

    if big_declines:
        lines.append("### 明显下降（≤ −10 分）")
        lines.append("")
        lines.append("| 学号 | 姓名 | Essay1 | Essay2 | 变化 |")
        lines.append("|---|---|---:|---:|---:|")
        for sid, name, o1, o2, delta in sorted(big_declines, key=lambda x: x[4]):
            lines.append(f"| {sid} | {name} | {o1:.0f} | {o2:.0f} | {delta:+.0f} |")
        lines.append("")

    top_gain = sorted(deltas, key=lambda x: -x[4])[:5]
    top_loss = sorted(deltas, key=lambda x: x[4])[:5]
    lines.append("### 变化最大（前 5 名提升 / 下降）")
    lines.append("")
    lines.append("**提升**")
    for sid, name, o1, o2, delta in top_gain:
        if delta <= 0:
            break
        lines.append(f"- {sid} {name}: {o1:.0f} → {o2:.0f} ({delta:+.0f})")
    lines.append("")
    lines.append("**下降**")
    for sid, name, o1, o2, delta in top_loss:
        if delta >= 0:
            break
        lines.append(f"- {sid} {name}: {o1:.0f} → {o2:.0f} ({delta:+.0f})")
    lines.append("")
    return "\n".join(lines)


def run_analysis(combined_path: Path, report_path: Path) -> str:
    e1 = read_sheet_scores(combined_path, "Essay 1")
    e2 = read_sheet_scores(combined_path, "Essay 2")
    body = [
        "# MINC 2026 Essay 分数分析",
        "",
        f"数据来源: `{combined_path.name}`",
        "",
        analyze_by_topic(e1, "Essay 1（Plants vs Psyche）"),
        analyze_by_topic(e2, "Essay 2（Brain vs AI2050）"),
        analyze_longitudinal(e1, e2),
    ]
    text = "\n".join(body)
    report_path.write_text(text, encoding="utf-8")
    return text


def main() -> int:
    parser = argparse.ArgumentParser(description="Essay 题目与纵向分数分析")
    parser.add_argument(
        "--scores",
        type=Path,
        default=None,
        help=f"汇总 xlsx（默认 downloads/{COMBINED_NAME}）",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=f"报告 markdown（默认 downloads/{REPORT_NAME}）",
    )
    args = parser.parse_args()
    root = download_root()
    combined = args.scores or (root / COMBINED_NAME)
    report = args.output or (root / REPORT_NAME)
    if not combined.exists():
        raise SystemExit(f"找不到: {combined}")
    text = run_analysis(combined, report)
    print(text)
    print(f"\n已写入: {report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
