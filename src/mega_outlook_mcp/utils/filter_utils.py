"""Outlook Restrict/DASL filter string builders (Windows)."""

from __future__ import annotations

from datetime import datetime

from .time_utils import outlook_restrict_format, outlook_restrict_format_locale


def escape_restrict(value: str) -> str:
    """Escape a value for inclusion in an Outlook Restrict string."""
    return value.replace("'", "''")


def escape_dasl(value: str) -> str:
    """Escape a value for inclusion in a DASL Items.Restrict query."""
    return value.replace("'", "''").replace('"', '""')


def build_time_range_restrict(
    field: str, start: datetime, end: datetime
) -> str:
    # Locale format: the US 'mm/dd/yyyy' form is read as dd/mm on pl-PL
    # whenever the day is <= 12, so mail ranges came back wrong.
    start_str = outlook_restrict_format_locale(start)
    end_str = outlook_restrict_format_locale(end)
    return f"[{field}] >= '{start_str}' AND [{field}] < '{end_str}'"


def build_calendar_overlap_restrict(start: datetime, end: datetime) -> str:
    """Events overlapping [start, end): '[Start] < end AND [End] > start',
    dates in the user's locale format (see outlook_restrict_format_locale)."""
    start_str = outlook_restrict_format_locale(start)
    end_str = outlook_restrict_format_locale(end)
    return f"[Start] < '{end_str}' AND [End] > '{start_str}'"
