"""outlook_list_folders."""

from __future__ import annotations

from dataclasses import asdict
from typing import Any

from ..backends.base import Backend
from ..models.inputs import ListFoldersInput


def register(mcp: Any, backend: Backend) -> None:
    @mcp.tool(
        name="outlook_list_folders",
        description=(
            "List all mail folders visible to the logged-in Outlook profile, in every data "
            "file. Default-mailbox paths are plain ('Skrzynka odbiorcza'); folders of other "
            "data files (PST archives, extra mailboxes) start with '//<store>/'. Any full_path "
            "can be passed as a folder to other tools."
        ),
    )
    async def outlook_list_folders(include_subfolders: bool = True) -> dict[str, Any]:
        params = ListFoldersInput(include_subfolders=include_subfolders)
        folders = await backend.list_folders(params.include_subfolders)
        return {"folders": [asdict(f) for f in folders]}

    if getattr(backend, "list_stores", None) is not None:

        @mcp.tool(
            name="outlook_list_stores",
            description=(
                "List the data files / mailboxes in the Outlook profile (default mailbox and "
                "PST archives) with the folder prefix to use for each, e.g. "
                "'//archive@example.com/' (a unique fragment like '//archive/' or the PST "
                "file name also works), and where each one's Inbox and Sent folders are."
            ),
        )
        async def outlook_list_stores() -> dict[str, Any]:
            return {"stores": await backend.list_stores()}  # type: ignore[attr-defined]
