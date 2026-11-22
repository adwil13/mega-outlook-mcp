"""Windows COM backend backed by pywin32.

All COM calls are marshalled through `OutlookComBridge` onto the STA
thread. Helpers `_sync_*` run on that thread and take `namespace` as their
first argument.
"""

from __future__ import annotations

import logging
import re
import time
from datetime import datetime, timezone
from typing import Any

from ..baseline import load_baseline
from ..com_bridge import OutlookComBridge
from ..constants import (
    ATTACH_TYPES_REAL,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_CONTACTS,
    OL_FOLDER_DELETED,
    OL_FOLDER_INBOX,
    OL_FOLDER_JUNK,
)
from ..errors import (
    AttachmentError,
    ConversationNotFoundError,
    EmailNotFoundError,
    FolderNotFoundError,
)
from ..utils.email_extract import detect_importance, domain_of, preview_text
from ..utils.filter_utils import build_time_range_restrict, escape_dasl
from ..utils.folder_utils import canonicalize_folder_name, split_folder_path
from ..utils.mapi_props import (
    PR_IN_REPLY_TO_ID,
    PR_INTERNET_MESSAGE_ID,
    PR_INTERNET_REFERENCES,
    PR_SENT_REPRESENTING_SMTP,
    PR_TRANSPORT_HEADERS,
    safe_get,
)
from ..utils.smtp_resolver import resolve_recipient, resolve_sender
from ..utils.subject_utils import normalize_subject
from .base import (
    AccountInfo,
    AttachmentInfo,
    Backend,
    CalendarEvent,
    Contact,
    DiagnosticsResult,
    EmailBody,
    EmailFullMetadata,
    EmailSummary,
    FolderInfo,
    FreeBusyResponse,
    FreeBusySlot,
    MailboxInfo,
    MailboxQuota,
    NoteInfo,
    OperationResult,
    OutOfOfficeStatus,
    RuleInfo,
    SignatureInfo,
    TaskInfo,
    ThreadMetadata,
)


class WindowsComBackend(Backend):
    def __init__(self) -> None:
        self._bridge = OutlookComBridge()

    # ------------------------------------------------------------------
    # Mailbox / folders
    # ------------------------------------------------------------------
    async def get_mailbox_info(self) -> MailboxInfo:
        return await self._bridge.execute(_sync_get_mailbox_info)

    async def list_folders(self, include_subfolders: bool) -> list[FolderInfo]:
        return await self._bridge.execute(_sync_list_folders, include_subfolders)

    # ------------------------------------------------------------------
    # Emails
    # ------------------------------------------------------------------
    async def get_emails_in_time_range(
        self,
        start_utc: datetime,
        end_utc: datetime,
        folders: list[str],
        max_results: int,
    ) -> list[EmailSummary]:
        return await self._bridge.execute(
            _sync_get_emails_in_time_range, start_utc, end_utc, folders, max_results
        )

    # Optional fast paths used by composite tools (feature-detected there, so
    # other backends need not implement them).
    async def get_emails_in_time_range_light(
        self,
        start_utc: datetime,
        end_utc: datetime,
        folders: list[str],
        max_results: int,
    ) -> list[EmailSummary]:
        """Like get_emails_in_time_range, but to_smtp/cc_smtp are empty and
        has_real_attachments is False until passed through enrich_summaries."""
        return await self._bridge.execute(
            _sync_get_emails_in_time_range, start_utc, end_utc, folders, max_results,
            details=False,
        )

    async def enrich_summaries(self, summaries: list[EmailSummary]) -> list[EmailSummary]:
        if not summaries:
            return []
        return await self._bridge.execute(_sync_enrich_summaries, list(summaries))

    async def get_plain_bodies(self, entry_ids: list[str]) -> dict[str, str]:
        if not entry_ids:
            return {}
        return await self._bridge.execute(_sync_get_plain_bodies, list(entry_ids))

    async def get_conversation_senders(
        self, conversation_id: str, anchor_entry_id: str | None = None
    ) -> list[EmailSummary] | None:
        return await self._bridge.execute(
            _sync_conversation_senders, conversation_id, anchor_entry_id
        )

    async def get_conversation_thread(
        self,
        conversation_id: str,
        conversation_topic: str | None,
        max_messages: int,
    ) -> list[EmailSummary]:
        return await self._bridge.execute(
            _sync_get_conversation_thread,
            conversation_id,
            conversation_topic,
            max_messages,
        )

    async def get_email_full_metadata(self, entry_id: str) -> EmailFullMetadata:
        return await self._bridge.execute(_sync_get_email_full_metadata, entry_id)

    async def save_attachment(
        self, entry_id: str, attachment_index: int, save_path: str
    ) -> AttachmentInfo:
        return await self._bridge.execute(
            _sync_save_attachment, entry_id, attachment_index, save_path
        )

    async def search_emails(
        self, query: str, folders: list[str], field: str, max_results: int
    ) -> list[EmailSummary]:
        return await self._bridge.execute(
            _sync_search_emails, query, folders, field, max_results
        )

    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------
    async def list_calendar_events(
        self,
        start_utc: datetime,
        end_utc: datetime,
        calendar_id: str | None,
        max_results: int,
    ) -> list[CalendarEvent]:
        return await self._bridge.execute(
            _sync_list_calendar_events, start_utc, end_utc, calendar_id, max_results
        )

    async def get_calendar_event(self, event_id: str) -> CalendarEvent:
        return await self._bridge.execute(_sync_get_calendar_event, event_id)

    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------
    async def list_contacts(self, limit: int, offset: int) -> list[Contact]:
        return await self._bridge.execute(_sync_list_contacts, limit, offset)

    async def search_contacts(self, query: str, limit: int) -> list[Contact]:
        return await self._bridge.execute(_sync_search_contacts, query, limit)

    async def get_contact(self, contact_id: str) -> Contact:
        return await self._bridge.execute(_sync_get_contact, contact_id)

    async def diagnostics(self, sample_message_id: str | None = None) -> DiagnosticsResult:
        return await self._bridge.execute(_sync_diagnostics, sample_message_id)

    # ------------------------------------------------------------------
    # Mail write
    # ------------------------------------------------------------------
    async def send_email(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._bridge.execute(_sync_send_email, to, subject, body, body_type, cc, bcc, attachments, send=True)

    async def create_draft(self, to, subject, body, body_type, cc, bcc, attachments) -> OperationResult:
        return await self._bridge.execute(_sync_send_email, to, subject, body, body_type, cc, bcc, attachments, send=False)

    async def reply_email(self, entry_id, body, body_type, reply_all) -> OperationResult:
        return await self._bridge.execute(_sync_reply_email, entry_id, body, body_type, reply_all)

    async def forward_email(self, entry_id, to, body, body_type) -> OperationResult:
        return await self._bridge.execute(_sync_forward_email, entry_id, to, body, body_type)

    # ------------------------------------------------------------------
    # Mail organize
    # ------------------------------------------------------------------
    async def mark_email_read(self, entry_id, is_read) -> OperationResult:
        return await self._bridge.execute(_sync_mark_email_read, entry_id, is_read)

    async def set_email_flag(self, entry_id, flag_status, due_date_utc) -> OperationResult:
        return await self._bridge.execute(_sync_set_email_flag, entry_id, flag_status, due_date_utc)

    async def set_email_categories(self, entry_id, categories) -> OperationResult:
        return await self._bridge.execute(_sync_set_email_categories, entry_id, categories)

    # ------------------------------------------------------------------
    # Mail destructive
    # ------------------------------------------------------------------
    async def move_email(self, entry_id, destination_folder) -> OperationResult:
        return await self._bridge.execute(_sync_move_email, entry_id, destination_folder)

    async def archive_email(self, entry_id) -> OperationResult:
        return await self._bridge.execute(_sync_move_email, entry_id, "Archive")

    async def delete_email(self, entry_id, permanent) -> OperationResult:
        return await self._bridge.execute(_sync_delete_email, entry_id, permanent)

    async def junk_email(self, entry_id) -> OperationResult:
        return await self._bridge.execute(_sync_junk_email, entry_id)

    # ------------------------------------------------------------------
    # Folder management
    # ------------------------------------------------------------------
    async def create_folder(self, parent_path, name) -> OperationResult:
        return await self._bridge.execute(_sync_create_folder, parent_path, name)

    async def rename_folder(self, folder_path, new_name) -> OperationResult:
        return await self._bridge.execute(_sync_rename_folder, folder_path, new_name)

    async def move_folder(self, folder_path, new_parent_path) -> OperationResult:
        return await self._bridge.execute(_sync_move_folder, folder_path, new_parent_path)

    async def delete_folder(self, folder_path) -> OperationResult:
        return await self._bridge.execute(_sync_delete_folder, folder_path)

    async def empty_folder(self, folder_path) -> OperationResult:
        return await self._bridge.execute(_sync_empty_folder, folder_path)

    # ------------------------------------------------------------------
    # Calendar write
    # ------------------------------------------------------------------
    async def create_calendar_event(self, subject, start_utc, end_utc, attendees, location, body, is_all_day) -> OperationResult:
        return await self._bridge.execute(_sync_create_calendar_event, subject, start_utc, end_utc, attendees, location, body, is_all_day)

    async def update_calendar_event(self, event_id, subject, start_utc, end_utc, location, body) -> OperationResult:
        return await self._bridge.execute(_sync_update_calendar_event, event_id, subject, start_utc, end_utc, location, body)

    async def delete_calendar_event(self, event_id) -> OperationResult:
        return await self._bridge.execute(_sync_delete_calendar_event, event_id)

    async def respond_to_event(self, event_id, response, send_response) -> OperationResult:
        return await self._bridge.execute(_sync_respond_to_event, event_id, response, send_response)

    # ------------------------------------------------------------------
    # Tasks
    # ------------------------------------------------------------------
    async def list_tasks(self, limit, offset) -> list[TaskInfo]:
        return await self._bridge.execute(_sync_list_tasks, limit, offset)

    async def search_tasks(self, query, limit) -> list[TaskInfo]:
        return await self._bridge.execute(_sync_search_tasks, query, limit)

    async def get_task(self, task_id) -> TaskInfo:
        return await self._bridge.execute(_sync_get_task, task_id)

    # ------------------------------------------------------------------
    # Notes
    # ------------------------------------------------------------------
    async def list_notes(self, limit, offset) -> list[NoteInfo]:
        return await self._bridge.execute(_sync_list_notes, limit, offset)

    async def search_notes(self, query, limit) -> list[NoteInfo]:
        return await self._bridge.execute(_sync_search_notes, query, limit)

    async def get_note(self, note_id) -> NoteInfo:
        return await self._bridge.execute(_sync_get_note, note_id)

    # ------------------------------------------------------------------
    # Accounts
    # ------------------------------------------------------------------
    async def list_accounts(self) -> list[AccountInfo]:
        return await self._bridge.execute(_sync_list_accounts)

    async def get_unread_count(self, folder) -> int:
        return await self._bridge.execute(_sync_get_unread_count, folder)

    # ------------------------------------------------------------------
    # Exchange-specific
    # ------------------------------------------------------------------
    async def get_out_of_office(self) -> OutOfOfficeStatus:
        return await self._bridge.execute(_sync_get_out_of_office)

    async def set_out_of_office(self, status: OutOfOfficeStatus) -> OperationResult:
        return await self._bridge.execute(_sync_set_out_of_office, status)

    async def get_signature(self, account_id: str | None) -> SignatureInfo:
        return await self._bridge.execute(_sync_get_signature, account_id)

    async def set_signature(self, account_id: str | None, body_html: str, body_plain: str) -> OperationResult:
        return await self._bridge.execute(_sync_set_signature, account_id, body_html, body_plain)

    async def list_rules(self) -> list[RuleInfo]:
        return await self._bridge.execute(_sync_list_rules)

    async def toggle_rule(self, rule_id: str, enabled: bool) -> OperationResult:
        return await self._bridge.execute(_sync_toggle_rule, rule_id, enabled)

    async def calendar_freebusy(self, smtps, start_utc, end_utc, slot_minutes) -> list[FreeBusyResponse]:
        return await self._bridge.execute(_sync_calendar_freebusy, smtps, start_utc, end_utc, slot_minutes)

    async def meeting_room_finder(self, start_utc, end_utc, capacity, location_hint) -> list[Contact]:
        return await self._bridge.execute(_sync_meeting_room_finder, start_utc, end_utc, capacity, location_hint)

    async def gal_search(self, query: str, limit: int) -> list[Contact]:
        return await self._bridge.execute(_sync_gal_search, query, limit)

    async def list_delegated_mailboxes(self) -> list[MailboxInfo]:
        return await self._bridge.execute(_sync_list_delegated_mailboxes)

    async def list_public_folders(self) -> list[FolderInfo]:
        return await self._bridge.execute(_sync_list_public_folders)

    async def get_mailbox_quota(self) -> MailboxQuota:
        return await self._bridge.execute(_sync_get_mailbox_quota)

    async def close(self) -> None:
        self._bridge.close()


# ---------------------------------------------------------------------------
# Synchronous helpers (STA thread).
# ---------------------------------------------------------------------------


# Well-known folders. Outlook localizes folder names ("Inbox" is
# "Skrzynka odbiorcza", "Posteingang", "Boîte de réception"...), and IMAP
# accounts keep some of them on the server under provider-specific names
# (e.g. "[Gmail]/..."). Nothing here depends on a language: the English
# words below are only this tool's API vocabulary, and every other name is
# read from Outlook itself (GetDefaultFolder(...).Name) at runtime.
_DEFAULT_FOLDER_IDS: dict[str, int] = {
    "inbox": 6,     # olFolderInbox
    "sent": 5,      # olFolderSentMail
    "drafts": 16,   # olFolderDrafts
    "deleted": 3,   # olFolderDeletedItems
    "junk": 23,     # olFolderJunk
    "outbox": 4,    # olFolderOutbox
}
_API_FOLDER_ALIASES: dict[str, str] = {
    "inbox": "inbox",
    "sent": "sent", "sent items": "sent", "sent mail": "sent",
    "drafts": "drafts",
    "deleted": "deleted", "deleted items": "deleted", "trash": "deleted",
    "junk": "junk", "junk email": "junk", "spam": "junk",
    "outbox": "outbox",
}
_PAREN_SUFFIX_RE = re.compile(r"\s*\([^)]*\)\s*$")


def _norm_name(name: str) -> str:
    """Lower-case, drop a trailing '(This computer only)'-style suffix."""
    return _PAREN_SUFFIX_RE.sub("", (name or "").strip()).strip().lower()


def _walk_path(root: Any, path: str) -> Any | None:
    folder = root
    for part in split_folder_path(path) or [path]:
        try:
            folder = folder.Folders[part]
        except Exception:
            return None
    return folder


def _folder_item_count(folder: Any) -> int:
    try:
        return int(folder.Items.Count)
    except Exception:
        return 0


def _folder_display_path(folder: Any, fallback: str) -> str:
    """Path as outlook_list_folders shows it, e.g. '[Gmail]/Wysłane'."""
    try:
        raw = folder.FolderPath  # "\\\\store\\[Gmail]\\Wysłane"
        parts = [p for p in raw.split("\\") if p]
        if len(parts) >= 2:
            return "/".join(parts[1:])
    except Exception:
        pass
    return fallback


def _localized_default_names(namespace: Any) -> dict[str, str]:
    """{normalized localized name: kind}, read from this Outlook install."""
    names: dict[str, str] = {}
    for kind, folder_id in _DEFAULT_FOLDER_IDS.items():
        try:
            names.setdefault(_norm_name(namespace.GetDefaultFolder(folder_id).Name), kind)
        except Exception:
            continue
    return names


def _folder_kind(namespace: Any, path: str) -> str | None:
    """Which well-known folder `path` means, in any UI language."""
    if not path:
        return None
    last = split_folder_path(path.strip())
    candidates = {_norm_name(path), _norm_name(last[-1]) if last else ""}
    for name in candidates:
        if name in _API_FOLDER_ALIASES:
            return _API_FOLDER_ALIASES[name]
    localized = _localized_default_names(namespace)
    for name in candidates:
        if name in localized:
            return localized[name]
    return None


def _is_sent_folder(namespace: Any, folder: Any, requested: str, label: str) -> bool:
    if _folder_kind(namespace, requested) == "sent" or _folder_kind(namespace, label) == "sent":
        return True
    try:
        entry_id = folder.EntryID
    except Exception:
        return False
    try:
        return namespace.GetDefaultFolder(_DEFAULT_FOLDER_IDS["sent"]).EntryID == entry_id
    except Exception:
        return False


def _resolve_folder(namespace: Any, path: str) -> Any:
    try:
        root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    except Exception:
        root = None

    # 1. The path exactly as given (real folder names always win).
    if root is not None and path and path.strip():
        folder = _walk_path(root, path.strip())
        if folder is not None:
            return folder

    # 2. A well-known folder, named either with this tool's English API word
    #    ("Inbox", "Sent Items") or with Outlook's own name for it in
    #    whatever language Outlook runs in: use Outlook's default folder.
    kind = _folder_kind(namespace, path)
    if kind is not None:
        default_folder = None
        try:
            default_folder = namespace.GetDefaultFolder(_DEFAULT_FOLDER_IDS[kind])
        except Exception:
            default_folder = None
        _log.info(
            "resolve %r -> kind=%s default=%r items=%s",
            path, kind,
            _folder_display_path(default_folder, "?") if default_folder is not None else None,
            _folder_item_count(default_folder) if default_folder is not None else None,
        )
        if default_folder is not None:
            return default_folder

    raise FolderNotFoundError(
        f"Folder {path!r} not found. Call outlook_list_folders to see available folders."
    )


def _pywintime_to_utc(value: Any) -> datetime | None:
    if value is None:
        return None
    try:
        year = value.year
    except Exception:
        return None
    try:
        return datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
            tzinfo=timezone.utc,
        )
    except Exception:
        return None


def _pywintime_local_to_utc(value: Any) -> datetime | None:
    """Convert an Outlook object-model time (local wall clock) to UTC.

    Properties such as MailItem.ReceivedTime / SentOn are returned by
    Outlook in local time, and pywin32 attaches a misleading tzinfo to
    them. Drop it, interpret the wall-clock value in the system time zone
    (DST-aware) and convert to UTC.
    """
    if value is None:
        return None
    try:
        naive = datetime(
            value.year, value.month, value.day, value.hour, value.minute, value.second
        )
    except Exception:
        return None
    if naive.year >= 4500:  # Outlook's "no date" sentinel (1/1/4501)
        return None
    try:
        return naive.astimezone(timezone.utc)
    except Exception:
        return None


def _default_account_smtp(namespace: Any) -> str:
    try:
        store_id = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Store.StoreID
    except Exception:
        store_id = None
    fallback = ""
    try:
        accounts = list(namespace.Accounts)
    except Exception:
        accounts = []
    for account in accounts:
        try:
            smtp = account.SmtpAddress or ""
        except Exception:
            continue
        if "@" not in smtp:
            continue
        if store_id is not None:
            try:
                if account.DeliveryStore.StoreID == store_id:
                    return smtp
            except Exception:
                pass
        fallback = fallback or smtp
    if fallback:
        return fallback
    try:
        addr = namespace.CurrentUser.Address or ""
        return addr if "@" in addr and "/" not in addr else ""
    except Exception:
        return ""


def _sync_get_mailbox_info(namespace: Any) -> MailboxInfo:
    user = namespace.CurrentUser
    smtp = ""
    try:
        smtp = safe_get(user.AddressEntry, PR_SENT_REPRESENTING_SMTP) or ""
    except Exception:
        smtp = ""
    if not smtp:
        try:
            exchange_user = user.AddressEntry.GetExchangeUser()
            smtp = exchange_user.PrimarySmtpAddress if exchange_user else ""
        except Exception:
            smtp = ""
    if not smtp:
        # Non-Exchange profiles (IMAP/POP): CurrentUser has no Exchange user;
        # use the SMTP address of the account that delivers to the default
        # Inbox's store, else the first account with one.
        smtp = _default_account_smtp(namespace)
    display = ""
    try:
        display = user.Name or ""
    except Exception:
        display = ""
    return MailboxInfo(
        display_name=display,
        smtp_address=smtp,
        domain=domain_of(smtp),
    )


def _sync_list_folders(namespace: Any, include_subfolders: bool) -> list[FolderInfo]:
    results: list[FolderInfo] = []

    def walk(folder: Any, parent_path: str | None) -> None:
        name = folder.Name
        full_path = name if parent_path is None else f"{parent_path}/{name}"
        try:
            item_count = int(folder.Items.Count)
        except Exception:
            item_count = 0
        try:
            unread_count = int(folder.UnReadItemCount)
        except Exception:
            unread_count = 0
        results.append(
            FolderInfo(
                name=name,
                full_path=full_path,
                item_count=item_count,
                unread_count=unread_count,
                parent_path=parent_path,
            )
        )
        if not include_subfolders:
            return
        for sub in folder.Folders:
            walk(sub, full_path)

    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    for folder in root.Folders:
        walk(folder, None)
    return results


def _extract_summary(item: Any, folder_path: str) -> EmailSummary:
    sender_smtp = resolve_sender(item)
    to_list: list[str] = []
    cc_list: list[str] = []
    try:
        for recip in item.Recipients:
            smtp = resolve_recipient(recip)
            if not smtp:
                continue
            kind = getattr(recip, "Type", 1)
            if kind == 2:
                cc_list.append(smtp)
            else:
                to_list.append(smtp)
    except Exception:
        pass
    try:
        received = _pywintime_local_to_utc(item.ReceivedTime)
    except Exception:
        received = None
    try:
        sent = _pywintime_local_to_utc(item.SentOn)
    except Exception:
        sent = None
    try:
        conv_id = item.ConversationID or ""
    except Exception:
        conv_id = ""
    if conv_id:
        try:
            _remember_conversation(conv_id, item.EntryID)
        except Exception:
            pass
    try:
        has_real = any(
            int(getattr(a, "Type", 0)) in ATTACH_TYPES_REAL for a in item.Attachments
        )
    except Exception:
        has_real = False
    try:
        body_plain = item.Body or ""
    except Exception:
        body_plain = ""
    return EmailSummary(
        entry_id=item.EntryID,
        conversation_id=conv_id,
        subject=getattr(item, "Subject", "") or "",
        sender_name=getattr(item, "SenderName", "") or "",
        sender_smtp=sender_smtp,
        to_smtp=to_list,
        cc_smtp=cc_list,
        received_utc=received,
        sent_utc=sent,
        is_read=not bool(getattr(item, "UnRead", False)),
        importance=detect_importance(getattr(item, "Importance", 1)),
        has_real_attachments=has_real,
        folder_path=folder_path,
        preview=preview_text(body_plain),
    )


def _sync_get_emails_in_time_range(
    namespace: Any,
    start_utc: datetime,
    end_utc: datetime,
    folders: list[str],
    max_results: int,
    details: bool = True,
) -> list[EmailSummary]:
    results: list[EmailSummary] = []
    for folder_name in folders:
        folder = _resolve_folder(namespace, folder_name)
        folder_label = _folder_display_path(folder, folder_name)
        field = "SentOn" if _is_sent_folder(namespace, folder, folder_name, folder_label) else "ReceivedTime"
        restrict = build_time_range_restrict(field, start_utc, end_utc)
        try:
            results.extend(
                _table_summaries(
                    namespace, folder, restrict, field, max_results - len(results), folder_label,
                    details,
                )
            )
        except Exception as exc:
            raise EmailNotFoundError(
                f"Restrict failed on {folder_name!r}: {exc}"
            ) from exc
        if len(results) >= max_results:
            break
    return results


# --- Table-based listing ------------------------------------------------------
#
# Opening a MailItem and reading its properties costs tens of milliseconds per
# message over COM; folder.GetTable() returns the same scalar properties in
# bulk (~0.4 ms/row). Rows are read from the table and an item is opened only
# for what a table cannot provide (recipient SMTP addresses, Exchange sender
# resolution, inline-vs-real attachment check when the message has any).

_PT = "http://schemas.microsoft.com/mapi/proptag/"
_COL_CONVERSATION_ID = _PT + "0x30130102"     # PR_CONVERSATION_ID (binary)
_COL_HAS_ATTACH = _PT + "0x0E1B000B"          # PR_HASATTACH
_COL_DELIVERY_TIME = _PT + "0x0E060040"       # PR_MESSAGE_DELIVERY_TIME (UTC)
_COL_SUBMIT_TIME = _PT + "0x00390040"         # PR_CLIENT_SUBMIT_TIME (UTC)
_COL_PREVIEW = "urn:schemas:httpmail:textdescription"
_SUMMARY_COLUMNS = [
    "EntryID",
    "Subject",
    "SenderName",
    "SenderEmailAddress",
    "UnRead",
    "Importance",
    "MessageClass",
    _COL_CONVERSATION_ID,
    _COL_HAS_ATTACH,
    _COL_DELIVERY_TIME,
    _COL_SUBMIT_TIME,
    _COL_PREVIEW,
]

# conversation_id -> entry_id of one message in it, filled as a side effect
# of every listing so get_conversation_thread can usually skip any search.
_CONVERSATION_ANCHORS: dict[str, str] = {}
_CONVERSATION_ANCHORS_MAX = 50_000


def _remember_conversation(conversation_id: str, entry_id: str) -> None:
    if not conversation_id or not entry_id:
        return
    if len(_CONVERSATION_ANCHORS) >= _CONVERSATION_ANCHORS_MAX:
        _CONVERSATION_ANCHORS.clear()
    _CONVERSATION_ANCHORS[_hex_id(conversation_id)] = entry_id


def _is_mail_class(message_class: Any) -> bool:
    # MailItem (olMail, Class 43) == IPM.Note and its sub-classes.
    return str(message_class or "").upper().startswith("IPM.NOTE")


def _iter_table_rows(table: Any, width: int):
    while not table.EndOfTable:
        try:
            block = table.GetArray(500)
        except Exception:
            block = None
        if block is None:
            row = table.GetNextRow()
            if row is None:
                return
            yield tuple(row.GetValues())
            continue
        if not block:
            return
        rows = list(block)
        if rows and len(rows) == width and len(rows[0]) != width:
            rows = list(zip(*rows))
        for row in rows:
            yield tuple(row)


def _item_details(item: Any, has_attach: Any) -> tuple[list[str], list[str], bool]:
    """(to, cc, has_real_attachments) — the parts a table row cannot give."""
    to_list: list[str] = []
    cc_list: list[str] = []
    has_real = False
    try:
        for recip in item.Recipients:
            smtp = resolve_recipient(recip)
            if not smtp:
                continue
            if getattr(recip, "Type", 1) == 2:
                cc_list.append(smtp)
            else:
                to_list.append(smtp)
    except Exception:
        pass
    if has_attach:
        try:
            has_real = any(
                int(getattr(a, "Type", 0)) in ATTACH_TYPES_REAL for a in item.Attachments
            )
        except Exception:
            has_real = False
    return to_list, cc_list, has_real


# entry_id -> PR_HASATTACH for summaries produced without details, so that
# enrich_summaries() can skip the attachment check exactly as a full read would.
_PENDING_HAS_ATTACH: dict[str, bool] = {}
_PENDING_HAS_ATTACH_MAX = 50_000


def _summary_from_row(
    namespace: Any, row: tuple, folder_path: str, details: bool = True
) -> EmailSummary:
    """Build a summary from a table row.

    details=False skips opening the item (recipients and the inline-vs-real
    attachment check are left empty); enrich_summaries() fills them later
    for just the summaries a caller actually returns.
    """
    (
        entry_id, subject, sender_name, sender_addr, unread, importance,
        _message_class, conversation, has_attach, delivered, submitted, preview,
    ) = row
    conv_id = _hex_id(conversation)
    if conv_id:
        _remember_conversation(conv_id, entry_id)

    sender_smtp = sender_addr if _looks_like_smtp(sender_addr) else ""
    item = None
    if details or not sender_smtp:
        try:
            item = namespace.GetItemFromID(entry_id)
        except Exception:
            item = None
    if not sender_smtp and item is not None:
        sender_smtp = resolve_sender(item)

    to_list: list[str] = []
    cc_list: list[str] = []
    has_real = False
    if details:
        if item is not None:
            to_list, cc_list, has_real = _item_details(item, has_attach)
    else:
        if len(_PENDING_HAS_ATTACH) >= _PENDING_HAS_ATTACH_MAX:
            _PENDING_HAS_ATTACH.clear()
        _PENDING_HAS_ATTACH[entry_id] = bool(has_attach)

    return EmailSummary(
        entry_id=entry_id,
        conversation_id=conv_id,
        subject=subject or "",
        sender_name=sender_name or "",
        sender_smtp=sender_smtp,
        to_smtp=to_list,
        cc_smtp=cc_list,
        received_utc=_pywintime_to_utc(delivered),
        sent_utc=_pywintime_to_utc(submitted),
        is_read=not bool(unread),
        importance=detect_importance(importance if importance is not None else 1),
        has_real_attachments=has_real,
        folder_path=folder_path,
        preview=preview_text(preview if isinstance(preview, str) else ""),
    )


def _looks_like_smtp(value: Any) -> bool:
    return isinstance(value, str) and "@" in value and "/" not in value


def _table_summaries(
    namespace: Any,
    folder: Any,
    table_filter: str,
    sort_field: str | None,
    max_count: int,
    folder_label: str,
    details: bool = True,
) -> list[EmailSummary]:
    if max_count <= 0:
        return []
    table = folder.GetTable(table_filter or "", _OL_USER_ITEMS)
    if sort_field:
        try:
            table.Sort(f"[{sort_field}]", True)
        except Exception:
            pass
    table.Columns.RemoveAll()
    for column in _SUMMARY_COLUMNS:
        table.Columns.Add(column)
    out: list[EmailSummary] = []
    for row in _iter_table_rows(table, len(_SUMMARY_COLUMNS)):
        if len(out) >= max_count:
            break
        if not row or not row[0] or not _is_mail_class(row[6]):
            continue
        out.append(_summary_from_row(namespace, row, folder_label, details))
    return out


def _sync_enrich_summaries(namespace: Any, summaries: list[EmailSummary]) -> list[EmailSummary]:
    """Fill recipients / has_real_attachments for summaries built with details=False."""
    from dataclasses import replace

    out: list[EmailSummary] = []
    for s in summaries:
        try:
            item = namespace.GetItemFromID(s.entry_id)
        except Exception:
            out.append(s)
            continue
        has_attach = _PENDING_HAS_ATTACH.get(s.entry_id, True)
        to_list, cc_list, has_real = _item_details(item, has_attach)
        out.append(replace(s, to_smtp=to_list, cc_smtp=cc_list, has_real_attachments=has_real))
    return out


def _sync_get_plain_bodies(namespace: Any, entry_ids: list[str]) -> dict[str, str]:
    """Full plain-text body per entry_id (table columns truncate at 255 chars)."""
    out: dict[str, str] = {}
    for entry_id in entry_ids:
        try:
            out[entry_id] = namespace.GetItemFromID(entry_id).Body or ""
        except Exception:
            continue
    return out


def _sync_conversation_senders(
    namespace: Any, conversation_id: str, anchor_entry_id: str | None
) -> list[EmailSummary] | None:
    """Sender + times of every message in a conversation, from the
    conversation table only (no item opened unless a sender is not SMTP).

    Returns None when the conversation cannot be located this way, so the
    caller can fall back to get_conversation_thread.
    """
    target = _hex_id(conversation_id)
    anchor = anchor_entry_id or _CONVERSATION_ANCHORS.get(target) or _find_conversation_anchor(
        namespace, target, "", time.monotonic() + _THREAD_SCAN_BUDGET_S
    )
    if not anchor:
        return None
    try:
        item = namespace.GetItemFromID(anchor)
        if _hex_id(getattr(item, "ConversationID", "")) != target:
            return None
        conversation = item.GetConversation()
    except Exception:
        return None
    if conversation is None:
        return None
    columns = ["EntryID", "SenderEmailAddress", "MessageClass", _COL_DELIVERY_TIME, _COL_SUBMIT_TIME]
    table = conversation.GetTable()
    table.Columns.RemoveAll()
    for column in columns:
        table.Columns.Add(column)
    out: list[EmailSummary] = []
    seen: set[str] = set()
    for entry_id, sender_addr, message_class, delivered, submitted in _iter_table_rows(table, len(columns)):
        if not entry_id or entry_id in seen or not _is_mail_class(message_class):
            continue
        seen.add(entry_id)
        sender_smtp = sender_addr if _looks_like_smtp(sender_addr) else ""
        if not sender_smtp:
            try:
                sender_smtp = resolve_sender(namespace.GetItemFromID(entry_id))
            except Exception:
                sender_smtp = ""
        out.append(
            EmailSummary(
                entry_id=entry_id,
                conversation_id=target,
                subject="",
                sender_name="",
                sender_smtp=sender_smtp,
                to_smtp=[],
                cc_smtp=[],
                received_utc=_pywintime_to_utc(delivered),
                sent_utc=_pywintime_to_utc(submitted),
                is_read=True,
                importance="normal",
                has_real_attachments=False,
                folder_path="",
                preview="",
            )
        )
    return out


# --- Conversation lookup ---------------------------------------------------
#
# The previous implementation opened every message in every top-level folder
# and compared item.ConversationID one COM call at a time. On a large IMAP
# store (tens of thousands of items, plus Outlook's sync-log folders) that
# takes many minutes, and because the COM bridge runs jobs one at a time on a
# single thread, every later tool call queued behind it and timed out too.
#
# This version reads lightweight MAPI tables (folder.GetTable) instead of
# opening items, skips Outlook's own sync-log folders, walks subfolders (so
# e.g. "[Gmail]/Wysłane" is included), narrows by topic when one is given and
# stops after a fixed time budget so it can never block the bridge for long.

_log = logging.getLogger(__name__)

_OL_MAIL_ITEM = 0  # OlItemType.olMailItem (Folder.DefaultItemType)
_OL_USER_ITEMS = 0  # OlTableContents.olUserItems
_OL_SKIP_DEFAULT_FOLDERS = (
    19,  # olFolderConflicts
    20,  # olFolderSyncIssues
    21,  # olFolderLocalFailures
    22,  # olFolderServerFailures
)
_PR_CONVERSATION_ID = "http://schemas.microsoft.com/mapi/proptag/0x30130102"
_PR_CONVERSATION_TOPIC = "http://schemas.microsoft.com/mapi/proptag/0x0070001F"
_THREAD_SCAN_BUDGET_S = 45.0
_TABLE_BATCH = 500
_TOPIC_PREFIX_RE = re.compile(r"^\s*(re|fw|fwd|odp|pd|aw|wg)\s*:\s*", re.IGNORECASE)


def _hex_id(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (bytes, bytearray, memoryview)):
        return bytes(value).hex().upper()
    return str(value).replace("-", "").replace(" ", "").strip().upper()


def _strip_topic_prefixes(topic: str) -> str:
    previous = None
    while previous != topic:
        previous = topic
        topic = _TOPIC_PREFIX_RE.sub("", topic)
    return topic.strip()


def _skipped_folder_ids(namespace: Any) -> set[str]:
    ids: set[str] = set()
    for folder_type in _OL_SKIP_DEFAULT_FOLDERS:
        try:
            ids.add(namespace.GetDefaultFolder(folder_type).EntryID)
        except Exception:
            continue
    return ids


def _iter_mail_folders(parent: Any, skip_ids: set[str], parent_path: str | None = None):
    try:
        children = list(parent.Folders)
    except Exception:
        return
    for folder in children:
        try:
            name = folder.Name
        except Exception:
            continue
        full_path = name if parent_path is None else f"{parent_path}/{name}"
        try:
            if folder.EntryID in skip_ids:
                continue
        except Exception:
            pass
        try:
            item_type = int(folder.DefaultItemType)
        except Exception:
            item_type = _OL_MAIL_ITEM
        if item_type == _OL_MAIL_ITEM:
            yield folder, full_path
        yield from _iter_mail_folders(folder, skip_ids, full_path)


def _table_rows(folder: Any, table_filter: str, columns: list[str], deadline: float):
    """Yield rows (tuples in `columns` order) of a folder's MAPI table."""
    table = folder.GetTable(table_filter, _OL_USER_ITEMS)
    table.Columns.RemoveAll()
    for column in columns:
        table.Columns.Add(column)
    width = len(columns)
    while not table.EndOfTable:
        if time.monotonic() > deadline:
            return
        try:
            block = table.GetArray(_TABLE_BATCH)
        except Exception:
            block = None
        if block is None:
            # Fallback: row-by-row reads (slower, but still no item opens).
            row = table.GetNextRow()
            if row is None:
                return
            values = []
            for idx in range(1, width + 1):
                try:
                    values.append(row.Item(idx))
                except Exception:
                    values.append(None)
            yield tuple(values)
            continue
        if not block:
            return
        rows = list(block)
        # pywin32 should hand back one tuple per row; transpose if it gave
        # one tuple per column instead.
        if rows and len(rows) == width and len(rows[0]) != width:
            rows = list(zip(*rows))
        for row in rows:
            yield tuple(row)


_ANCHOR_SEARCH_BUDGET_S = 20.0


def _find_conversation_anchor(
    namespace: Any, target: str, topic: str, deadline: float
) -> str | None:
    """EntryID of one message in conversation `target`, from Inbox/Sent tables."""
    budget_end = min(deadline, time.monotonic() + _ANCHOR_SEARCH_BUDGET_S)
    table_filter = ""
    if topic:
        table_filter = (
            f'@SQL="{_PR_CONVERSATION_TOPIC}" = \'{topic.replace(chr(39), chr(39) * 2)}\''
        )
    for folder_id, sort_field in (
        (OL_FOLDER_INBOX, "ReceivedTime"),
        (_DEFAULT_FOLDER_IDS["sent"], "SentOn"),
    ):
        try:
            folder = namespace.GetDefaultFolder(folder_id)
            try:
                table = folder.GetTable(table_filter, _OL_USER_ITEMS)
            except Exception:
                table = folder.GetTable("", _OL_USER_ITEMS)
            try:
                table.Sort(f"[{sort_field}]", True)
            except Exception:
                pass
            table.Columns.RemoveAll()
            table.Columns.Add("EntryID")
            table.Columns.Add(_PR_CONVERSATION_ID)
            for row in _iter_table_rows(table, 2):
                if time.monotonic() > budget_end:
                    return None
                if len(row) > 1 and row[0] and _hex_id(row[1]) == target:
                    _remember_conversation(target, row[0])
                    return row[0]
        except Exception as exc:  # noqa: BLE001
            _log.info("anchor search skipped folder %s: %s", folder_id, exc)
    return None


def _thread_via_conversation(
    namespace: Any, anchor_entry_id: str, target: str, max_messages: int
) -> list[EmailSummary]:
    item = namespace.GetItemFromID(anchor_entry_id)
    if _hex_id(getattr(item, "ConversationID", "")) != target:
        return []
    conversation = item.GetConversation()
    if conversation is None:
        return []
    table = conversation.GetTable()
    table.Columns.RemoveAll()
    table.Columns.Add("EntryID")
    collected: list[EmailSummary] = []
    seen: set[str] = set()
    for row in _iter_table_rows(table, 1):
        if len(collected) >= max_messages:
            break
        entry_id = row[0] if row else None
        if not entry_id or entry_id in seen:
            continue
        seen.add(entry_id)
        try:
            message = namespace.GetItemFromID(entry_id)
        except Exception:
            continue
        if getattr(message, "Class", 43) != 43:
            continue
        try:
            parent = message.Parent
            label = _folder_display_path(parent, getattr(parent, "Name", ""))
        except Exception:
            label = ""
        collected.append(_extract_summary(message, label))
    collected.sort(
        key=lambda s: s.sent_utc or s.received_utc or datetime.min.replace(tzinfo=timezone.utc)
    )
    return collected


def _sync_get_conversation_thread(
    namespace: Any,
    conversation_id: str,
    conversation_topic: str | None,
    max_messages: int,
) -> list[EmailSummary]:
    target = _hex_id(conversation_id)
    deadline = time.monotonic() + _THREAD_SCAN_BUDGET_S
    topic = _strip_topic_prefixes(conversation_topic) if conversation_topic else ""

    # Fast path: find one message of the conversation, then let Outlook's
    # own conversation index (MailItem.GetConversation) return the rest,
    # across every folder of the store.
    anchor = _CONVERSATION_ANCHORS.get(target) or _find_conversation_anchor(
        namespace, target, topic, deadline
    )
    if anchor:
        try:
            thread = _thread_via_conversation(namespace, anchor, target, max_messages)
        except Exception as exc:  # noqa: BLE001
            _log.info("GetConversation path failed for %s: %s", conversation_id, exc)
            thread = None
        if thread:
            return thread

    table_filter = ""
    if topic:
        table_filter = (
            f'@SQL="{_PR_CONVERSATION_TOPIC}" = \'{topic.replace(chr(39), chr(39) * 2)}\''
        )

    columns = ["EntryID", _PR_CONVERSATION_ID]
    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    skip_ids = _skipped_folder_ids(namespace)

    collected: list[EmailSummary] = []
    seen: set[str] = set()
    rows_without_id = 0
    timed_out = False

    for folder, folder_path in _iter_mail_folders(root, skip_ids):
        if len(collected) >= max_messages:
            break
        if time.monotonic() > deadline:
            timed_out = True
            break
        try:
            try:
                rows = _table_rows(folder, table_filter, columns, deadline)
                row_iter = iter(rows)
                first = next(row_iter, None)
            except Exception:
                if not table_filter:
                    raise
                # Filter rejected by this store: scan the folder unfiltered.
                rows = _table_rows(folder, "", columns, deadline)
                row_iter = iter(rows)
                first = next(row_iter, None)
            pending = [] if first is None else [first]
            for row in _chain(pending, row_iter):
                if len(collected) >= max_messages:
                    break
                entry_id = row[0] if row else None
                if not entry_id or entry_id in seen:
                    continue
                row_cid = _hex_id(row[1] if len(row) > 1 else None)
                if row_cid:
                    if row_cid != target:
                        continue
                else:
                    rows_without_id += 1
                    # Store does not expose the ID in tables. Only open the
                    # item when a topic filter keeps the candidate set small.
                    if not topic:
                        continue
                try:
                    item = namespace.GetItemFromID(entry_id)
                except Exception:
                    continue
                if getattr(item, "Class", 43) != 43:
                    continue
                if not row_cid:
                    try:
                        if _hex_id(item.ConversationID) != target:
                            continue
                    except Exception:
                        continue
                seen.add(entry_id)
                collected.append(_extract_summary(item, folder_path))
        except Exception as exc:  # noqa: BLE001
            _log.warning("conversation scan skipped folder %r: %s", folder_path, exc)
            continue
    if time.monotonic() > deadline:
        timed_out = True

    if not collected:
        hints = []
        if timed_out:
            hints.append(
                f"the scan stopped after {_THREAD_SCAN_BUDGET_S:.0f}s"
            )
        if rows_without_id and not topic:
            hints.append("this store does not expose conversation IDs in tables")
        hint = ""
        if hints:
            hint = " (" + "; ".join(hints) + "). Pass conversation_topic to narrow the search."
        raise ConversationNotFoundError(
            f"No messages found for conversation_id {conversation_id!r}{hint}"
        )
    if timed_out:
        _log.warning(
            "conversation %s: scan budget reached, returning %d message(s) found so far",
            conversation_id,
            len(collected),
        )
    collected.sort(key=lambda s: s.sent_utc or s.received_utc or datetime.min.replace(tzinfo=timezone.utc))
    return collected


def _chain(head: list[Any], tail: Any):
    yield from head
    yield from tail


def _sync_get_email_full_metadata(namespace: Any, entry_id: str) -> EmailFullMetadata:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    folder_path = ""
    try:
        folder_path = item.Parent.Name if item.Parent else ""
    except Exception:
        folder_path = ""
    summary = _extract_summary(item, folder_path)
    try:
        plain = item.Body or ""
    except Exception:
        plain = ""
    try:
        html = item.HTMLBody or None
    except Exception:
        html = None
    attachments: list[AttachmentInfo] = []
    try:
        for idx, att in enumerate(item.Attachments, start=1):
            kind = int(getattr(att, "Type", 0))
            attachments.append(
                AttachmentInfo(
                    index=idx,
                    filename=getattr(att, "FileName", "") or "",
                    size_bytes=int(getattr(att, "Size", 0) or 0),
                    is_inline=kind not in ATTACH_TYPES_REAL,
                )
            )
    except Exception:
        pass

    msg_id = safe_get(item, PR_INTERNET_MESSAGE_ID)
    in_reply_to = safe_get(item, PR_IN_REPLY_TO_ID)
    references = safe_get(item, PR_INTERNET_REFERENCES)
    headers_raw = safe_get(item, PR_TRANSPORT_HEADERS)
    references_list: list[str] | None = None
    if references:
        references_list = [r for r in references.split() if r]

    return EmailFullMetadata(
        summary=summary,
        body=EmailBody(plain_text=plain, html=html),
        attachments=attachments,
        internet_message_id=msg_id or None,
        in_reply_to=in_reply_to or None,
        references=references_list,
        mapi_headers_raw=headers_raw or None,
        delegation_sender_smtp=summary.sender_smtp or None,
        delegation_representing_smtp=(
            safe_get(item, PR_SENT_REPRESENTING_SMTP) or None
        ),
    )


def _sync_save_attachment(
    namespace: Any, entry_id: str, attachment_index: int, save_path: str
) -> AttachmentInfo:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        att = item.Attachments.Item(attachment_index)
    except Exception as exc:
        raise AttachmentError(
            f"Attachment index {attachment_index} not found on email {entry_id!r}."
        ) from exc
    kind = int(getattr(att, "Type", 0))
    if kind not in ATTACH_TYPES_REAL:
        raise AttachmentError(
            f"Attachment {attachment_index} on {entry_id!r} is inline (type={kind}); refusing to save."
        )
    try:
        att.SaveAsFile(save_path)
    except Exception as exc:
        raise AttachmentError(f"SaveAsFile failed for {save_path!r}: {exc}") from exc
    return AttachmentInfo(
        index=attachment_index,
        filename=getattr(att, "FileName", "") or "",
        size_bytes=int(getattr(att, "Size", 0) or 0),
        is_inline=False,
    )


def _sync_search_emails(
    namespace: Any,
    query: str,
    folders: list[str],
    field: str,
    max_results: int,
) -> list[EmailSummary]:
    q = escape_dasl(query)
    subject = f"\"urn:schemas:httpmail:subject\" LIKE '%{q}%'"
    sender = (
        f"(\"urn:schemas:httpmail:fromemail\" LIKE '%{q}%' "
        f"OR \"urn:schemas:httpmail:fromname\" LIKE '%{q}%')"
    )
    body_like = f"\"urn:schemas:httpmail:textdescription\" LIKE '%{q}%'"
    # A substring LIKE over message bodies is evaluated item by item and
    # takes minutes on large stores; when the store is indexed by Windows
    # Search, ci_phrasematch answers the same question from the index.
    body_indexed = f"\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{q}'"

    def build(use_index: bool) -> str:
        body = body_indexed if use_index else body_like
        if field == "subject":
            return f"@SQL={subject}"
        if field == "sender":
            return f"@SQL={sender}"
        if field == "body":
            return f"@SQL={body}"
        return f"@SQL=({subject} OR {sender} OR {body})"

    results: list[EmailSummary] = []
    for folder_name in folders:
        folder = _resolve_folder(namespace, folder_name)
        folder_label = _folder_display_path(folder, folder_name)
        use_index = False
        if field not in ("subject", "sender"):
            try:
                use_index = bool(folder.Store.IsInstantSearchEnabled)
            except Exception:
                use_index = False
        remaining = max_results - len(results)
        try:
            found = _table_summaries(
                namespace, folder, build(use_index), "ReceivedTime", remaining, folder_label
            )
        except Exception:
            if not use_index:
                continue
            try:
                found = _table_summaries(
                    namespace, folder, build(False), "ReceivedTime", remaining, folder_label
                )
            except Exception:
                continue
        results.extend(found)
        if len(results) >= max_results:
            break
    return results


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------


def _extract_calendar_event(item: Any) -> CalendarEvent:
    attendees: list[str] = []
    try:
        for recip in item.Recipients:
            smtp = resolve_recipient(recip)
            if smtp:
                attendees.append(smtp)
    except Exception:
        pass
    organizer = ""
    try:
        organizer = resolve_sender(item)
    except Exception:
        organizer = ""
    try:
        start = _pywintime_to_utc(item.StartUTC)
    except Exception:
        try:
            start = _pywintime_local_to_utc(item.Start)
        except Exception:
            start = None
    try:
        end = _pywintime_to_utc(item.EndUTC)
    except Exception:
        try:
            end = _pywintime_local_to_utc(item.End)
        except Exception:
            end = None
    return CalendarEvent(
        event_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        organizer_smtp=organizer,
        start_utc=start,
        end_utc=end,
        location=getattr(item, "Location", "") or "",
        is_all_day=bool(getattr(item, "AllDayEvent", False)),
        attendees_smtp=attendees,
        body_plain=getattr(item, "Body", "") or "",
        is_recurring=bool(getattr(item, "IsRecurring", False)),
    )


def _sync_list_calendar_events(
    namespace: Any,
    start_utc: datetime,
    end_utc: datetime,
    calendar_id: str | None,
    max_results: int,
) -> list[CalendarEvent]:
    calendar = (
        namespace.GetFolderFromID(calendar_id)
        if calendar_id
        else namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    )
    items = calendar.Items
    try:
        items.IncludeRecurrences = True
        items.Sort("[Start]", False)
    except Exception:
        pass
    restrict = build_time_range_restrict("Start", start_utc, end_utc)
    try:
        filtered = items.Restrict(restrict)
    except Exception as exc:
        raise EmailNotFoundError(f"Calendar Restrict failed: {exc}") from exc
    results: list[CalendarEvent] = []
    for item in filtered:
        if len(results) >= max_results:
            break
        results.append(_extract_calendar_event(item))
    return results


def _sync_get_calendar_event(namespace: Any, event_id: str) -> CalendarEvent:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No calendar event with id {event_id!r}.") from exc
    return _extract_calendar_event(item)


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------


def _extract_contact(item: Any) -> Contact:
    phones: dict[str, str] = {}
    for label, attr in (
        ("business", "BusinessTelephoneNumber"),
        ("home", "HomeTelephoneNumber"),
        ("mobile", "MobileTelephoneNumber"),
    ):
        value = getattr(item, attr, "") or ""
        if value:
            phones[label] = value
    other: list[str] = []
    for attr in ("Email2Address", "Email3Address"):
        value = getattr(item, attr, "") or ""
        if value:
            other.append(value)
    return Contact(
        contact_id=item.EntryID,
        display_name=getattr(item, "FullName", "") or getattr(item, "FileAs", "") or "",
        primary_smtp=getattr(item, "Email1Address", "") or "",
        other_smtps=other,
        company=getattr(item, "CompanyName", "") or "",
        job_title=getattr(item, "JobTitle", "") or "",
        phone_numbers=phones,
    )


def _sync_list_contacts(namespace: Any, limit: int, offset: int) -> list[Contact]:
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    results: list[Contact] = []
    idx = 0
    for item in folder.Items:
        if getattr(item, "Class", 40) != 40:  # olContact
            continue
        if idx < offset:
            idx += 1
            continue
        if len(results) >= limit:
            break
        results.append(_extract_contact(item))
        idx += 1
    return results


def _sync_search_contacts(namespace: Any, query: str, limit: int) -> list[Contact]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    results: list[Contact] = []
    for item in folder.Items:
        if getattr(item, "Class", 40) != 40:
            continue
        haystack = " ".join(
            (
                getattr(item, "FullName", "") or "",
                getattr(item, "Email1Address", "") or "",
                getattr(item, "CompanyName", "") or "",
                getattr(item, "JobTitle", "") or "",
            )
        ).lower()
        if q in haystack:
            results.append(_extract_contact(item))
            if len(results) >= limit:
                break
    return results


def _sync_get_contact(namespace: Any, contact_id: str) -> Contact:
    try:
        item = namespace.GetItemFromID(contact_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No contact with id {contact_id!r}.") from exc
    return _extract_contact(item)


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

# Map from "field that broke" → list of tools impacted, used to populate
# DiagnosticsResult.affected_tools when a probe fails.
_FIELD_TO_TOOLS: dict[str, list[str]] = {
    "Inbox": ["outlook_get_emails_in_time_range", "outlook_search_emails"],
    "Calendar": ["outlook_list_calendar_events", "outlook_get_calendar_event"],
    "Contacts": ["outlook_list_contacts", "outlook_search_contacts", "outlook_get_contact"],
    "PR_INTERNET_MESSAGE_ID": ["outlook_get_email_full_metadata"],
    "PR_IN_REPLY_TO_ID": ["outlook_get_email_full_metadata"],
    "PR_INTERNET_REFERENCES": ["outlook_get_email_full_metadata"],
    "PR_TRANSPORT_HEADERS": ["outlook_get_email_full_metadata"],
    "ConversationID": [
        "outlook_get_conversation_thread",
        "outlook_get_thread_metadata",
        "outlook_get_emails_in_time_range",
    ],
    "PropertyAccessor": ["outlook_get_email_full_metadata"],
}


def _sync_diagnostics(namespace: Any, sample_message_id: str | None) -> DiagnosticsResult:
    baseline = load_baseline()
    notes: list[str] = []
    probed: dict[str, str] = {}
    affected: set[str] = set()

    # Outlook version
    try:
        outlook_app = namespace.Application
        version = str(outlook_app.Version)
    except Exception as exc:
        version = f"unknown (error: {exc})"

    # Default folder probes
    defaults = baseline.get("windows_com", {}).get("default_folders", {})
    for name, idx in defaults.items():
        try:
            folder = namespace.GetDefaultFolder(idx)
            _ = folder.Name
            probed[f"OlDefaultFolders.{name}"] = "ok"
        except Exception as exc:
            probed[f"OlDefaultFolders.{name}"] = f"error:{exc}"
            for t in _FIELD_TO_TOOLS.get(name, []):
                affected.add(t)

    # Pick a sample message if not supplied: first item in Inbox.
    sample_item = None
    if sample_message_id:
        try:
            sample_item = namespace.GetItemFromID(sample_message_id)
        except Exception as exc:
            notes.append(f"Could not load supplied sample id: {exc}")
    if sample_item is None:
        try:
            inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            if items.Count > 0:
                sample_item = items.GetFirst()
        except Exception as exc:
            notes.append(f"Could not auto-pick sample message: {exc}")

    if sample_item is None:
        notes.append("No sample message available; per-message field probes skipped.")
    else:
        # Item-level property probes
        for prop in baseline.get("windows_com", {}).get("item_properties", []):
            try:
                _ = getattr(sample_item, prop)
                probed[f"MailItem.{prop}"] = "ok"
            except Exception as exc:
                probed[f"MailItem.{prop}"] = f"error:{exc}"
                for t in _FIELD_TO_TOOLS.get(prop, []):
                    affected.add(t)
        # MAPI tag probes
        for tag_name, tag_value in (
            ("PR_INTERNET_MESSAGE_ID", PR_INTERNET_MESSAGE_ID),
            ("PR_IN_REPLY_TO_ID", PR_IN_REPLY_TO_ID),
            ("PR_INTERNET_REFERENCES", PR_INTERNET_REFERENCES),
            ("PR_TRANSPORT_HEADERS", PR_TRANSPORT_HEADERS),
            ("PR_SENT_REPRESENTING_SMTP", PR_SENT_REPRESENTING_SMTP),
        ):
            value = safe_get(sample_item, tag_value)
            probed[f"MAPI.{tag_name}"] = "ok" if value is not None else "missing"
            if value is None:
                for t in _FIELD_TO_TOOLS.get(tag_name, []):
                    affected.add(t)

    missing = [k for k, v in probed.items() if v != "ok"]
    if not missing:
        status = "HEALTHY"
    elif len(missing) >= len(probed) // 2:
        status = "BROKEN"
    else:
        status = f"DEGRADED:{len(missing)}_fields"

    baseline_version = baseline.get("baseline_outlook_versions", {}).get("windows", "unknown")
    return DiagnosticsResult(
        platform="windows",
        outlook_version=version,
        baseline_version=baseline_version,
        status=status,
        probed_fields=probed,
        affected_tools=sorted(affected),
        notes=notes,
    )


# ---------------------------------------------------------------------------
# Phase 3 (parity) sync helpers
# ---------------------------------------------------------------------------

OL_TASK_FOLDER_INDEX = 13  # OlDefaultFolders.olFolderTasks
OL_NOTES_FOLDER_INDEX = 12  # OlDefaultFolders.olFolderNotes


def _make_mail_item(namespace: Any, to, subject, body, body_type, cc, bcc, attachments) -> Any:
    app = namespace.Application
    item = app.CreateItem(0)  # olMailItem
    item.Subject = subject or ""
    if (body_type or "plain").lower() == "html":
        item.HTMLBody = body or ""
    else:
        item.Body = body or ""
    if to:
        item.To = "; ".join(to)
    if cc:
        item.CC = "; ".join(cc)
    if bcc:
        item.BCC = "; ".join(bcc)
    for path in attachments or []:
        try:
            item.Attachments.Add(path)
        except Exception:
            pass
    return item


def _sync_send_email(namespace: Any, to, subject, body, body_type, cc, bcc, attachments, send: bool) -> OperationResult:
    item = _make_mail_item(namespace, to, subject, body, body_type, cc, bcc, attachments)
    try:
        if send:
            item.Send()
            return OperationResult(ok=True, message="sent", affected_id=getattr(item, "EntryID", None))
        item.Save()
        return OperationResult(ok=True, message="saved as draft", affected_id=getattr(item, "EntryID", None))
    except Exception as exc:
        return OperationResult(ok=False, message=f"send/save failed: {exc}")


def _sync_reply_email(namespace: Any, entry_id: str, body: str, body_type: str, reply_all: bool) -> OperationResult:
    try:
        original = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    reply = original.ReplyAll() if reply_all else original.Reply()
    if (body_type or "plain").lower() == "html":
        reply.HTMLBody = (body or "") + (reply.HTMLBody or "")
    else:
        reply.Body = (body or "") + "\n\n" + (reply.Body or "")
    try:
        reply.Send()
    except Exception as exc:
        return OperationResult(ok=False, message=f"reply send failed: {exc}")
    return OperationResult(ok=True, message="replied")


def _sync_forward_email(namespace: Any, entry_id: str, to: list[str], body: str, body_type: str) -> OperationResult:
    try:
        original = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if (body_type or "plain").lower() == "html":
        fwd.HTMLBody = (body or "") + (fwd.HTMLBody or "")
    else:
        fwd.Body = (body or "") + "\n\n" + (fwd.Body or "")
    try:
        fwd.Send()
    except Exception as exc:
        return OperationResult(ok=False, message=f"forward send failed: {exc}")
    return OperationResult(ok=True, message="forwarded")


def _sync_mark_email_read(namespace: Any, entry_id: str, is_read: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        item.UnRead = not is_read
        item.Save()
        return OperationResult(ok=True, message=f"marked {'read' if is_read else 'unread'}", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"mark failed: {exc}")


def _sync_set_email_flag(namespace: Any, entry_id: str, flag_status: str, due_date_utc) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        # FlagStatus: 0 = none, 1 = complete, 2 = marked
        if flag_status == "complete":
            item.FlagStatus = 1
        elif flag_status == "marked":
            item.FlagStatus = 2
        else:
            item.FlagStatus = 0
        if due_date_utc is not None:
            try:
                item.TaskDueDate = due_date_utc
            except Exception:
                pass
        item.Save()
        return OperationResult(ok=True, message=f"flag set to {flag_status}", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"flag set failed: {exc}")


def _sync_set_email_categories(namespace: Any, entry_id: str, categories: list[str]) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        item.Categories = ", ".join(categories) if categories else ""
        item.Save()
        return OperationResult(ok=True, message="categories set", affected_id=entry_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"category set failed: {exc}")


def _sync_move_email(namespace: Any, entry_id: str, destination_folder: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        target = _resolve_folder(namespace, destination_folder)
        moved = item.Move(target)
        return OperationResult(ok=True, message=f"moved to {destination_folder}",
                               affected_id=getattr(moved, "EntryID", None))
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"move failed: {exc}")


def _sync_delete_email(namespace: Any, entry_id: str, permanent: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        if permanent:
            item.Delete()
            return OperationResult(ok=True, message="permanently deleted")
        deleted = namespace.GetDefaultFolder(OL_FOLDER_DELETED)
        item.Move(deleted)
        return OperationResult(ok=True, message="moved to Deleted Items")
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete failed: {exc}")


def _sync_junk_email(namespace: Any, entry_id: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(entry_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No email with entry_id {entry_id!r}.") from exc
    try:
        junk = namespace.GetDefaultFolder(OL_FOLDER_JUNK)
        item.Move(junk)
        return OperationResult(ok=True, message="moved to Junk")
    except Exception as exc:
        return OperationResult(ok=False, message=f"junk failed: {exc}")


def _sync_create_folder(namespace: Any, parent_path: str | None, name: str) -> OperationResult:
    try:
        if parent_path:
            parent = _resolve_folder(namespace, parent_path)
        else:
            parent = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
        new_folder = parent.Folders.Add(name)
        return OperationResult(ok=True, message="created", affected_id=getattr(new_folder, "EntryID", None))
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"create folder failed: {exc}")


def _sync_rename_folder(namespace: Any, folder_path: str, new_name: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        f.Name = new_name
        return OperationResult(ok=True, message="renamed")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"rename failed: {exc}")


def _sync_move_folder(namespace: Any, folder_path: str, new_parent_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        new_parent = _resolve_folder(namespace, new_parent_path)
        f.MoveTo(new_parent)
        return OperationResult(ok=True, message="moved folder")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"move folder failed: {exc}")


def _sync_delete_folder(namespace: Any, folder_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        f.Delete()
        return OperationResult(ok=True, message="deleted folder")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete folder failed: {exc}")


def _sync_empty_folder(namespace: Any, folder_path: str) -> OperationResult:
    try:
        f = _resolve_folder(namespace, folder_path)
        items = list(f.Items)  # snapshot before mutation
        for item in items:
            try:
                item.Delete()
            except Exception:
                pass
        return OperationResult(ok=True, message=f"emptied ({len(items)} items)")
    except FolderNotFoundError:
        raise
    except Exception as exc:
        return OperationResult(ok=False, message=f"empty folder failed: {exc}")


def _sync_create_calendar_event(namespace: Any, subject, start_utc, end_utc, attendees, location, body, is_all_day) -> OperationResult:
    try:
        app = namespace.Application
        item = app.CreateItem(1)  # olAppointmentItem
        item.Subject = subject or ""
        item.Start = start_utc
        item.End = end_utc
        item.Location = location or ""
        item.Body = body or ""
        item.AllDayEvent = bool(is_all_day)
        if attendees:
            item.MeetingStatus = 1  # olMeeting
            for a in attendees:
                recip = item.Recipients.Add(a)
                recip.Type = 1  # olRequired
            item.Recipients.ResolveAll()
            item.Send()
            return OperationResult(ok=True, message="meeting sent",
                                   affected_id=getattr(item, "EntryID", None))
        item.Save()
        return OperationResult(ok=True, message="event saved",
                               affected_id=getattr(item, "EntryID", None))
    except Exception as exc:
        return OperationResult(ok=False, message=f"create event failed: {exc}")


def _sync_update_calendar_event(namespace: Any, event_id, subject, start_utc, end_utc, location, body) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        if subject is not None:
            item.Subject = subject
        if start_utc is not None:
            item.Start = start_utc
        if end_utc is not None:
            item.End = end_utc
        if location is not None:
            item.Location = location
        if body is not None:
            item.Body = body
        item.Save()
        return OperationResult(ok=True, message="updated", affected_id=event_id)
    except Exception as exc:
        return OperationResult(ok=False, message=f"update failed: {exc}")


def _sync_delete_calendar_event(namespace: Any, event_id: str) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        item.Delete()
        return OperationResult(ok=True, message="deleted")
    except Exception as exc:
        return OperationResult(ok=False, message=f"delete failed: {exc}")


def _sync_respond_to_event(namespace: Any, event_id: str, response: str, send_response: bool) -> OperationResult:
    try:
        item = namespace.GetItemFromID(event_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No event with id {event_id!r}.") from exc
    try:
        # OlMeetingResponse: olMeetingAccepted=3, olMeetingTentative=2, olMeetingDeclined=4
        code = {"accept": 3, "tentative": 2, "decline": 4}.get(response.lower())
        if code is None:
            return OperationResult(ok=False, message=f"unknown response {response!r}")
        resp = item.Respond(code, True, send_response)
        if send_response and resp is not None:
            try:
                resp.Send()
            except Exception:
                pass
        return OperationResult(ok=True, message=f"responded {response}")
    except Exception as exc:
        return OperationResult(ok=False, message=f"respond failed: {exc}")


def _extract_task(item: Any) -> TaskInfo:
    return TaskInfo(
        task_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        due_date_utc=_pywintime_to_utc(getattr(item, "DueDate", None)),
        is_complete=bool(getattr(item, "Complete", False)),
        importance=detect_importance(getattr(item, "Importance", 1)),
        body_plain=getattr(item, "Body", "") or "",
    )


def _sync_list_tasks(namespace: Any, limit: int, offset: int) -> list[TaskInfo]:
    folder = namespace.GetDefaultFolder(OL_TASK_FOLDER_INDEX)
    out: list[TaskInfo] = []
    idx = 0
    for item in folder.Items:
        if idx < offset:
            idx += 1
            continue
        if len(out) >= limit:
            break
        out.append(_extract_task(item))
        idx += 1
    return out


def _sync_search_tasks(namespace: Any, query: str, limit: int) -> list[TaskInfo]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_TASK_FOLDER_INDEX)
    out: list[TaskInfo] = []
    for item in folder.Items:
        if q in (getattr(item, "Subject", "") or "").lower() or q in (getattr(item, "Body", "") or "").lower():
            out.append(_extract_task(item))
            if len(out) >= limit:
                break
    return out


def _sync_get_task(namespace: Any, task_id: str) -> TaskInfo:
    try:
        item = namespace.GetItemFromID(task_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No task with id {task_id!r}.") from exc
    return _extract_task(item)


def _extract_note(item: Any) -> NoteInfo:
    return NoteInfo(
        note_id=item.EntryID,
        subject=getattr(item, "Subject", "") or "",
        body_plain=getattr(item, "Body", "") or "",
        last_modified_utc=_pywintime_local_to_utc(getattr(item, "LastModificationTime", None)),
    )


def _sync_list_notes(namespace: Any, limit: int, offset: int) -> list[NoteInfo]:
    folder = namespace.GetDefaultFolder(OL_NOTES_FOLDER_INDEX)
    out: list[NoteInfo] = []
    idx = 0
    for item in folder.Items:
        if idx < offset:
            idx += 1
            continue
        if len(out) >= limit:
            break
        out.append(_extract_note(item))
        idx += 1
    return out


def _sync_search_notes(namespace: Any, query: str, limit: int) -> list[NoteInfo]:
    q = query.lower()
    folder = namespace.GetDefaultFolder(OL_NOTES_FOLDER_INDEX)
    out: list[NoteInfo] = []
    for item in folder.Items:
        if q in (getattr(item, "Subject", "") or "").lower() or q in (getattr(item, "Body", "") or "").lower():
            out.append(_extract_note(item))
            if len(out) >= limit:
                break
    return out


def _sync_get_note(namespace: Any, note_id: str) -> NoteInfo:
    try:
        item = namespace.GetItemFromID(note_id)
    except Exception as exc:
        raise EmailNotFoundError(f"No note with id {note_id!r}.") from exc
    return _extract_note(item)


def _sync_list_accounts(namespace: Any) -> list[AccountInfo]:
    out: list[AccountInfo] = []
    try:
        for acct in namespace.Application.Session.Accounts:
            account_type_map = {0: "exchange", 1: "imap", 2: "pop", 3: "exchange", 5: "other"}
            try:
                kind = int(getattr(acct, "AccountType", 5))
            except Exception:
                kind = 5
            out.append(
                AccountInfo(
                    account_id=str(getattr(acct, "DisplayName", "") or ""),
                    display_name=getattr(acct, "DisplayName", "") or "",
                    smtp_address=getattr(acct, "SmtpAddress", "") or "",
                    account_type=account_type_map.get(kind, "other"),
                )
            )
    except Exception:
        pass
    return out


def _sync_get_unread_count(namespace: Any, folder: str | None) -> int:
    if folder:
        f = _resolve_folder(namespace, folder)
        try:
            return int(f.UnReadItemCount)
        except Exception:
            return 0
    # Total across all folders.
    total = 0

    def walk(folder: Any) -> None:
        nonlocal total
        try:
            total += int(folder.UnReadItemCount)
        except Exception:
            pass
        for sub in folder.Folders:
            walk(sub)

    root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent
    for f in root.Folders:
        walk(f)
    return total


# ---------------------------------------------------------------------------
# Phase 4 (Exchange-specific) sync helpers
# ---------------------------------------------------------------------------


def _sync_get_out_of_office(namespace: Any) -> OutOfOfficeStatus:
    try:
        ms = namespace.Application.Session.Stores.DefaultStore
        rules = namespace.Application.Session.OOFTemplates  # not standard; fallback below
    except Exception:
        rules = None
    # Outlook COM exposes Account.AutomaticReplies (Outlook 2016+).
    try:
        account = namespace.Application.Session.Accounts.Item(1)
        ar = account.AutomaticReplies
        return OutOfOfficeStatus(
            enabled=bool(ar.AutomaticRepliesEnabled),
            internal_message=str(ar.InternalReply or ""),
            external_message=str(ar.ExternalReply or ""),
            start_utc=_pywintime_to_utc(getattr(ar, "StartTime", None)),
            end_utc=_pywintime_to_utc(getattr(ar, "EndTime", None)),
            external_audience=("all" if getattr(ar, "ExternalAudience", 0) == 2 else
                               "contacts_only" if getattr(ar, "ExternalAudience", 0) == 1 else "none"),
        )
    except Exception as exc:
        return OutOfOfficeStatus(
            enabled=False,
            internal_message=f"unavailable: {exc}",
            external_message="",
        )


def _sync_set_out_of_office(namespace: Any, status: OutOfOfficeStatus) -> OperationResult:
    try:
        account = namespace.Application.Session.Accounts.Item(1)
        ar = account.AutomaticReplies
        ar.AutomaticRepliesEnabled = bool(status.enabled)
        ar.InternalReply = status.internal_message or ""
        ar.ExternalReply = status.external_message or ""
        if status.start_utc is not None:
            ar.StartTime = status.start_utc
        if status.end_utc is not None:
            ar.EndTime = status.end_utc
        audience_map = {"none": 0, "contacts_only": 1, "all": 2}
        ar.ExternalAudience = audience_map.get(status.external_audience or "all", 2)
        return OperationResult(ok=True, message="OOO updated")
    except Exception as exc:
        return OperationResult(ok=False, message=f"OOO set failed: {exc}")


def _sync_get_signature(namespace: Any, account_id: str | None) -> SignatureInfo:
    # Outlook stores signatures as files under %APPDATA%\Microsoft\Signatures
    # Reading them via COM is not supported; we read the filesystem directly.
    import os
    sig_dir = os.path.expandvars(r"%APPDATA%\Microsoft\Signatures")
    name = ""
    plain = ""
    html = ""
    try:
        if os.path.isdir(sig_dir):
            entries = sorted(
                [e for e in os.listdir(sig_dir) if e.lower().endswith(".htm") or e.lower().endswith(".html")]
            )
            if entries:
                name = os.path.splitext(entries[0])[0]
                with open(os.path.join(sig_dir, entries[0]), "r", encoding="utf-8", errors="replace") as fh:
                    html = fh.read()
                txt_path = os.path.join(sig_dir, name + ".txt")
                if os.path.exists(txt_path):
                    with open(txt_path, "r", encoding="utf-8", errors="replace") as fh:
                        plain = fh.read()
    except Exception:
        pass
    return SignatureInfo(name=name, body_html=html, body_plain=plain)


def _sync_set_signature(namespace: Any, account_id: str | None, body_html: str, body_plain: str) -> OperationResult:
    import os
    sig_dir = os.path.expandvars(r"%APPDATA%\Microsoft\Signatures")
    try:
        os.makedirs(sig_dir, exist_ok=True)
        sig_name = "default"
        with open(os.path.join(sig_dir, sig_name + ".htm"), "w", encoding="utf-8") as fh:
            fh.write(body_html or "")
        with open(os.path.join(sig_dir, sig_name + ".txt"), "w", encoding="utf-8") as fh:
            fh.write(body_plain or "")
        return OperationResult(ok=True, message=f"signature written to {sig_dir}")
    except Exception as exc:
        return OperationResult(ok=False, message=f"signature write failed: {exc}")


def _sync_list_rules(namespace: Any) -> list[RuleInfo]:
    try:
        store = namespace.Application.Session.DefaultStore
        rules = store.GetRules()
    except Exception:
        return []
    out: list[RuleInfo] = []
    for r in rules:
        try:
            out.append(
                RuleInfo(
                    rule_id=str(getattr(r, "ID", "") or getattr(r, "Name", "")),
                    name=str(getattr(r, "Name", "") or ""),
                    enabled=bool(getattr(r, "Enabled", False)),
                    description=str(getattr(r, "Description", "") or ""),
                )
            )
        except Exception:
            continue
    return out


def _sync_toggle_rule(namespace: Any, rule_id: str, enabled: bool) -> OperationResult:
    try:
        store = namespace.Application.Session.DefaultStore
        rules = store.GetRules()
        for r in rules:
            if str(getattr(r, "ID", "") or getattr(r, "Name", "")) == rule_id or str(getattr(r, "Name", "")) == rule_id:
                r.Enabled = bool(enabled)
                rules.Save()
                return OperationResult(ok=True, message=f"rule {'enabled' if enabled else 'disabled'}")
        return OperationResult(ok=False, message=f"rule {rule_id!r} not found")
    except Exception as exc:
        return OperationResult(ok=False, message=f"rule toggle failed: {exc}")


def _sync_calendar_freebusy(namespace: Any, smtps, start_utc, end_utc, slot_minutes) -> list[FreeBusyResponse]:
    out: list[FreeBusyResponse] = []
    duration_minutes = int((end_utc - start_utc).total_seconds() / 60)
    for smtp in smtps:
        try:
            recip = namespace.CreateRecipient(smtp)
            recip.Resolve()
            # GetFreeBusy returns a string like "00021..." where each char is a status code.
            fb = recip.AddressEntry.GetFreeBusy(start_utc, slot_minutes, True)
            slots: list[FreeBusySlot] = []
            status_map = {"0": "free", "1": "tentative", "2": "busy", "3": "oof", "4": "unknown"}
            from datetime import timedelta
            for i, ch in enumerate(fb[: max(1, duration_minutes // slot_minutes)]):
                slot_start = start_utc + timedelta(minutes=i * slot_minutes)
                slots.append(
                    FreeBusySlot(
                        start_utc=slot_start,
                        end_utc=slot_start + timedelta(minutes=slot_minutes),
                        status=status_map.get(ch, "unknown"),
                    )
                )
            out.append(FreeBusyResponse(smtp=smtp, slots=slots))
        except Exception as exc:
            out.append(FreeBusyResponse(smtp=smtp, slots=[]))
    return out


def _sync_meeting_room_finder(namespace: Any, start_utc, end_utc, capacity, location_hint):
    # Exchange's "Find Room" requires the Resource Address List, which COM
    # exposes via AddressLists with kind == 5 (olAddressListResource).
    rooms: list[Contact] = []
    try:
        for al in namespace.AddressLists:
            try:
                kind = int(getattr(al, "AddressListType", 0))
            except Exception:
                kind = 0
            if kind != 5:
                continue
            for entry in al.AddressEntries:
                try:
                    smtp = entry.GetExchangeUser().PrimarySmtpAddress if entry.GetExchangeUser() else (entry.Address or "")
                except Exception:
                    smtp = entry.Address or ""
                rooms.append(
                    Contact(
                        contact_id=getattr(entry, "ID", "") or smtp,
                        display_name=getattr(entry, "Name", "") or "",
                        primary_smtp=smtp,
                    )
                )
    except Exception:
        pass
    if location_hint:
        hint = location_hint.lower()
        rooms = [r for r in rooms if hint in (r.display_name or "").lower()]
    # Filter by free/busy in the requested window.
    available: list[Contact] = []
    smtps = [r.primary_smtp for r in rooms if r.primary_smtp]
    if smtps:
        slot_minutes = max(15, int((end_utc - start_utc).total_seconds() / 60))
        fb = _sync_calendar_freebusy(namespace, smtps, start_utc, end_utc, slot_minutes)
        free_set = {f.smtp for f in fb if all(s.status == "free" for s in f.slots)}
        available = [r for r in rooms if r.primary_smtp in free_set]
    return available[: max(1, capacity)]  # capacity limits results, not enforces room size (room size = no COM API)


def _sync_gal_search(namespace: Any, query: str, limit: int) -> list[Contact]:
    q = query.lower()
    out: list[Contact] = []
    try:
        gal = namespace.GetGlobalAddressList()
    except Exception:
        return []
    for entry in gal.AddressEntries:
        if len(out) >= limit:
            break
        try:
            name = getattr(entry, "Name", "") or ""
            smtp = ""
            try:
                user = entry.GetExchangeUser()
                if user:
                    smtp = user.PrimarySmtpAddress or ""
            except Exception:
                pass
            if not smtp:
                smtp = getattr(entry, "Address", "") or ""
            haystack = (name + " " + smtp).lower()
            if q not in haystack:
                continue
            company = ""
            title = ""
            try:
                user = entry.GetExchangeUser()
                if user:
                    company = getattr(user, "CompanyName", "") or ""
                    title = getattr(user, "JobTitle", "") or ""
            except Exception:
                pass
            out.append(
                Contact(
                    contact_id=getattr(entry, "ID", "") or smtp,
                    display_name=name,
                    primary_smtp=smtp,
                    company=company,
                    job_title=title,
                )
            )
        except Exception:
            continue
    return out


def _sync_list_delegated_mailboxes(namespace: Any) -> list[MailboxInfo]:
    out: list[MailboxInfo] = []
    try:
        for store in namespace.Stores:
            try:
                if not store.IsDataFileStore:
                    name = getattr(store, "DisplayName", "") or ""
                    smtp = ""
                    try:
                        # Best-effort SMTP via the root folder owner.
                        root = store.GetRootFolder()
                        smtp = getattr(root, "FolderPath", "") or ""
                    except Exception:
                        pass
                    out.append(
                        MailboxInfo(
                            display_name=name,
                            smtp_address=smtp,
                            domain=domain_of(smtp),
                        )
                    )
            except Exception:
                continue
    except Exception:
        pass
    return out


def _sync_list_public_folders(namespace: Any) -> list[FolderInfo]:
    out: list[FolderInfo] = []
    try:
        for store in namespace.Stores:
            name = getattr(store, "DisplayName", "") or ""
            if "public folder" in name.lower():
                root = store.GetRootFolder()
                for f in root.Folders:
                    try:
                        out.append(
                            FolderInfo(
                                name=getattr(f, "Name", "") or "",
                                full_path=f"public/{getattr(f, 'Name', '')}",
                                item_count=int(getattr(f.Items, "Count", 0)),
                                unread_count=int(getattr(f, "UnReadItemCount", 0)),
                                parent_path="public",
                            )
                        )
                    except Exception:
                        continue
    except Exception:
        pass
    return out


def _sync_get_mailbox_quota(namespace: Any) -> MailboxQuota:
    try:
        store = namespace.Application.Session.DefaultStore
        # PR_MESSAGE_SIZE_EXTENDED gives total store size.
        PR_MESSAGE_SIZE_EXTENDED = "http://schemas.microsoft.com/mapi/proptag/0x0E08000B"
        used = int(safe_get(store, PR_MESSAGE_SIZE_EXTENDED) or 0)
        # PR_PROHIBIT_SEND_QUOTA / PR_STORAGE_QUOTA are not always exposed.
        PR_PROHIBIT_SEND_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x666E0003"
        PR_PROHIBIT_RECEIVE_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x666A0003"
        PR_STORAGE_QUOTA = "http://schemas.microsoft.com/mapi/proptag/0x66700003"
        prohibit_send = safe_get(store, PR_PROHIBIT_SEND_QUOTA)
        warning = safe_get(store, PR_STORAGE_QUOTA)
        prohibit_recv = safe_get(store, PR_PROHIBIT_RECEIVE_QUOTA)
        total = int(prohibit_recv or 0) * 1024 if prohibit_recv else 0
        item_count = 0
        try:
            root = store.GetRootFolder()
            for f in root.Folders:
                try:
                    item_count += int(f.Items.Count)
                except Exception:
                    continue
        except Exception:
            pass
        return MailboxQuota(
            total_bytes=total,
            used_bytes=used,
            warning_bytes=int(warning) * 1024 if warning else None,
            prohibit_send_bytes=int(prohibit_send) * 1024 if prohibit_send else None,
            item_count=item_count,
        )
    except Exception:
        return MailboxQuota(total_bytes=0, used_bytes=0, warning_bytes=None, prohibit_send_bytes=None, item_count=0)


__all__ = [
    "WindowsComBackend",
    "normalize_subject",  # re-exported so tools/thread_tools can use from backend
    "ThreadMetadata",
]
