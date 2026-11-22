"""outlook_list_calendar_events, outlook_get_calendar_event."""

from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from typing import Any

from ..backends.base import Backend, CalendarEvent
from ..models.inputs import GetCalendarEventInput, ListCalendarEventsInput


def _event_to_json(e: CalendarEvent) -> dict[str, Any]:
    raw = asdict(e)
    for key in ("start_utc", "end_utc"):
        value = raw.get(key)
        if isinstance(value, datetime):
            raw[key] = value.isoformat()
    return raw


LIST_BODY_MAX_CHARS = 300


def _trim_body(raw: dict[str, Any], max_chars: int) -> dict[str, Any]:
    body = raw.get("body_plain") or ""
    if max_chars <= 0:
        raw["body_plain"] = ""
        raw["body_truncated"] = bool(body)
    elif len(body) > max_chars:
        raw["body_plain"] = body[:max_chars].rstrip() + "…"
        raw["body_truncated"] = True
    else:
        raw["body_truncated"] = False
    return raw


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_list_calendar_events",
        description=(
            "List calendar events overlapping [start_utc, end_utc), sorted by start, with "
            "recurring series expanded into individual occurrences. Dates are ISO 8601; "
            "without an offset they are taken as UTC. calendar_id defaults to the default "
            "mailbox's calendar; it also accepts a folder EntryID, a folder path in the same "
            "syntax as mail folders (e.g. '//archive/Kalendarz'), or just a data file "
            "('//archive/') for that file's default calendar. See outlook_list_stores for each "
            "data file's calendar_path. body_plain is cut to body_max_chars (default 300, "
            "0 = omit; body_truncated tells if it was cut); use outlook_get_calendar_event "
            "for the full body."
        ),
    )
    async def outlook_list_calendar_events(
        start_utc: str,
        end_utc: str,
        calendar_id: str | None = None,
        max_results: int = 200,
        body_max_chars: int = LIST_BODY_MAX_CHARS,
    ) -> dict[str, Any]:
        params = ListCalendarEventsInput(
            start_utc=datetime.fromisoformat(start_utc),
            end_utc=datetime.fromisoformat(end_utc),
            calendar_id=calendar_id,
            max_results=max_results,
        )
        events = await backend.list_calendar_events(
            params.start_utc, params.end_utc, params.calendar_id, params.max_results
        )
        return {"events": [_trim_body(_event_to_json(e), body_max_chars) for e in events]}

    @mcp.tool(
        name="outlook_get_calendar_event",
        description="Return a single calendar event including attendees and body.",
    )
    async def outlook_get_calendar_event(event_id: str) -> dict[str, Any]:
        params = GetCalendarEventInput(event_id=event_id)
        event = await backend.get_calendar_event(params.event_id)
        return _event_to_json(event)
