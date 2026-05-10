from __future__ import annotations

import re

CNP_RE = re.compile(r"\b[1-9]\d{12}\b")
DATE_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b")
PHONE_RE = re.compile(r"\b(?:\+?40|0)\s?7\d{2}\s?\d{3}\s?\d{3}\b")
EMAIL_RE = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")

# Heuristic for Romanian all-caps names after standard formulae.
NAME_AFTER_PATIENT_RE = re.compile(
    r"(?i)(vă informăm că|va informam ca|pacient(?:ul|a)?\s*:?)\s+([A-ZĂÂÎȘȚ][A-ZĂÂÎȘȚ\-]+(?:\s+[A-ZĂÂÎȘȚ][A-ZĂÂÎȘȚ\-]+){1,3})"
)


def mask_cnp(cnp: str) -> str:
    return cnp[:1] + "***********" + cnp[-1:]


def anonymize_text(text: str, mask_dates: bool = False) -> str:
    """Basic anonymization for pasted examples. Review manually before sharing."""
    text = CNP_RE.sub(lambda m: mask_cnp(m.group(0)), text)
    text = PHONE_RE.sub("[TELEFON_ANONIMIZAT]", text)
    text = EMAIL_RE.sub("[EMAIL_ANONIMIZAT]", text)
    text = NAME_AFTER_PATIENT_RE.sub(lambda m: f"{m.group(1)} [PACIENT_ANONIMIZAT]", text)
    if mask_dates:
        text = DATE_RE.sub("[DATA_ANONIMIZATA]", text)
    return text
