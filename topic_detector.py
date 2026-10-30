"""Detect which essay topic a student chose from their docx."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document

REQUIREMENT_TEXT = (
    "Show critical and creative thinking by engaging with multiple perspectives, "
    "including counterarguments. Balance academic evidence with personal insights, "
    "using critical thinking to question assumptions and creativity to develop "
    "fresh interpretations."
)

# Essay 1
ESSAY1_TOPIC_1_TITLE = "Do plants have consciousness?"
ESSAY1_TOPIC_2_TITLE = (
    'Find out more about "psyche", and share your understanding to '
    "its symbolic meaning."
)

# Essay 2
ESSAY2_TOPIC_1_TITLE = (
    "Do you think the structure of the brain has been designed or has been formed "
    "by chance through evolution. Do you think the current design is good/bad?"
)
ESSAY2_TOPIC_2_TITLE = (
    "Predict the development of AI by the year 2050 and analyze its potential "
    "impact on a specific field, such as education, science, society, workplace, "
    "entertainment etc. This analysis should not only focus on the technical "
    "aspects of AI but also explore how AI might reshape the field in terms of "
    "human interactions and ethics"
)


def essay_number_from_dir(download_dir: Path | None) -> int:
    if download_dir is None:
        return 1
    match = re.search(r"essay\s*(\d+)", download_dir.name, re.I)
    return int(match.group(1)) if match else 1


def extract_docx_text(path: Path) -> str:
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return "\n".join(parts)


def detect_essay1_topic(text: str) -> int:
    lowered = text.lower()

    topic1_signals = [
        "do plants have consciousness" in lowered,
        bool(re.search(r"\bplant(s)?\b", lowered) and "conscious" in lowered),
    ]
    topic2_signals = [
        "psyche" in lowered,
        "symbolic meaning" in lowered,
        bool(re.search(r"\bpsyche\b", lowered) and "symbol" in lowered),
    ]

    score1 = sum(topic1_signals) + len(re.findall(r"conscious|plant", lowered)) * 0.1
    score2 = sum(topic2_signals) + len(re.findall(r"psyche|symbolic", lowered)) * 0.1

    if score1 == score2 == 0:
        return 1
    return 1 if score1 >= score2 else 2


def detect_essay2_topic(text: str) -> int:
    lowered = text.lower()

    topic1_signals = [
        "structure of the brain" in lowered,
        "brain" in lowered and ("evolution" in lowered or "designed" in lowered),
        "formed by chance" in lowered,
        "current design is good" in lowered or "current design is bad" in lowered,
        bool(re.search(r"\bbrain\b", lowered) and "design" in lowered),
    ]
    topic2_signals = [
        "2050" in lowered,
        "artificial intelligence" in lowered,
        bool(re.search(r"\bai\b", lowered)),
        "predict the development" in lowered,
        "human interactions and ethics" in lowered,
        "impact on" in lowered
        and any(
            w in lowered
            for w in (
                "education",
                "science",
                "society",
                "workplace",
                "entertainment",
            )
        ),
    ]

    score1 = sum(topic1_signals) + len(re.findall(r"brain|evolution|designed", lowered)) * 0.05
    score2 = sum(topic2_signals) + len(
        re.findall(r"\bai\b|2050|artificial intelligence", lowered)
    ) * 0.05

    if score1 == score2 == 0:
        return 1
    return 1 if score1 >= score2 else 2


def detect_topic(text: str, *, essay_number: int = 1) -> int:
    if essay_number >= 2:
        return detect_essay2_topic(text)
    return detect_essay1_topic(text)


def topic_title(topic_id: int, *, essay_number: int = 1) -> str:
    if essay_number >= 2:
        return ESSAY2_TOPIC_1_TITLE if topic_id == 1 else ESSAY2_TOPIC_2_TITLE
    return ESSAY1_TOPIC_1_TITLE if topic_id == 1 else ESSAY1_TOPIC_2_TITLE


def topic_short_label(topic_id: int, *, essay_number: int = 1) -> str:
    if essay_number >= 2:
        return "Brain" if topic_id == 1 else "AI2050"
    return "Plants" if topic_id == 1 else "Psyche"


def grading_rule_for(topic_id: int, *, essay_number: int = 1) -> str:
    """Yuketang rule: Essay 2 option 1 → Scientific, option 2 → Reflective."""
    if essay_number >= 2:
        return "Scientific" if topic_id == 1 else "Reflective"
    return "Reflective"


def build_topic_field(text: str, *, essay_number: int = 1) -> tuple[int, str]:
    topic_id = detect_topic(text, essay_number=essay_number)
    title = topic_title(topic_id, essay_number=essay_number)
    return topic_id, f"{title} {REQUIREMENT_TEXT}"


def build_topic_field_from_docx(
    path: Path, download_dir: Path | None = None
) -> tuple[int, str]:
    essay_number = essay_number_from_dir(download_dir)
    return build_topic_field(extract_docx_text(path), essay_number=essay_number)


def describe_topic(path: Path, download_dir: Path | None = None) -> tuple[int, str, str]:
    """Return (topic_id, short_label, grading_rule)."""
    essay_number = essay_number_from_dir(download_dir)
    text = extract_docx_text(path)
    topic_id = detect_topic(text, essay_number=essay_number)
    return (
        topic_id,
        topic_short_label(topic_id, essay_number=essay_number),
        grading_rule_for(topic_id, essay_number=essay_number),
    )
