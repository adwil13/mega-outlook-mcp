"""Backend-agnostic helpers for building EmailSummary / EmailFullMetadata."""

from __future__ import annotations


def preview_text(plain_body: str, length: int = 200) -> str:
    text = " ".join((plain_body or "").split())
    if len(text) <= length:
        return text
    return text[: length - 1].rstrip() + "…"


import re

# Where the quoted history of a reply/forward starts, in the formats Outlook,
# Gmail, Apple Mail and Thunderbird produce (English / Polish / German /
# French / Spanish header words).
_QUOTE_START = re.compile(
    r"^[ \t]*(?:"
    r"(?:from|od|von|de)[ \t]*:[ \t]*\S.*"                        # Outlook header block
    r"|-{2,}[ \t]*(?:original message|wiadomość oryginalna|ursprüngliche nachricht"
    r"|forwarded message|wiadomość przekazana dalej|message d'origine).*"
    r"|_{10,}[ \t]*"                                              # Outlook separator line
    r"|.*\b(?:wrote|napisał\(a\)|napisała|napisał|pisze|schrieb|a écrit|escribió)[ \t]*:[ \t]*"
    r"|>.*"                                                       # '>'-quoted lines
    r")$",
    re.IGNORECASE | re.MULTILINE,
)


def own_text(plain_body: str) -> str:
    """The author's own part of a message: everything before quoted history."""
    body = plain_body or ""
    match = _QUOTE_START.search(body)
    return body[: match.start()] if match else body


def detect_importance(value: object) -> str:
    from ..constants import IMPORTANCE_MAP

    try:
        idx = int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return "normal"
    return IMPORTANCE_MAP.get(idx, "normal")


def domain_of(smtp: str) -> str:
    if not smtp or "@" not in smtp:
        return ""
    return smtp.rsplit("@", 1)[-1].lower()
