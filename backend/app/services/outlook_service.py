"""
outlook_service.py
------------------
Thread-safe Outlook COM bridge.  Every public function that touches the COM
object initialises/uninitialises the apartment itself so it is safe to call
from any worker thread.
"""

import pythoncom
import win32com.client
from datetime import datetime
import logging

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _ns():
    """Return a fresh MAPI namespace for the calling thread."""
    pythoncom.CoInitialize()
    outlook = win32com.client.DispatchEx("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    try:
        ns.Logon(None, None, False, False)
    except Exception:
        pass
    return ns


def _folder_dict(folder, store_name):
    """Serialize a COM folder into a plain dict (no COM references)."""
    try:
        count = folder.Items.Count
    except Exception:
        count = 0
    return {
        "store":      store_name,
        "name":       folder.Name,
        "full_path":  folder.FolderPath,
        "entry_id":   folder.EntryID,
        "item_count": count,
        "children":   []
    }


def _recurse(parent_com, store_name, node_dict):
    """Recursively populate node_dict['children'] from COM subfolder."""
    try:
        for subfolder in parent_com.Folders:
            child = _folder_dict(subfolder, store_name)
            node_dict["children"].append(child)
            _recurse(subfolder, store_name, child)
    except Exception as exc:
        log.warning("Error recursing into folder: %s", exc)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def enumerate_stores_tree():
    """
    Returns a list of store-level dicts, each with a nested 'children' tree.
    No COM objects are returned – everything is serialised to plain dicts.

    Schema::
        [
          {
            "store": "My Mailbox",
            "name":  "Top of Information Store",
            "full_path": "\\\\My Mailbox\\...",
            "entry_id":  "<hex>",
            "item_count": 0,
            "children": [ ... ]   # recursive
          },
          ...
        ]
    """
    ns = _ns()
    stores = []
    try:
        for store in ns.Stores:
            try:
                root = store.GetRootFolder()
                node = _folder_dict(root, store.DisplayName)
                _recurse(root, store.DisplayName, node)
                stores.append(node)
            except Exception as exc:
                log.warning("Could not read store '%s': %s", store.DisplayName, exc)
    finally:
        pythoncom.CoUninitialize()
    return stores


def enumerate_all_folders():
    """
    Flat list of all folders (kept for backward-compat with scan / search code).
    Returns plain dicts – no COM objects.
    """
    def _flatten(node, result):
        result.append(node)
        for child in node.get("children", []):
            _flatten(child, result)

    stores = enumerate_stores_tree()
    flat = []
    for store in stores:
        _flatten(store, flat)
    return flat


def search_mails(query_params: dict):
    """
    Searches across all folders using Outlook's built-in Restrict filter
    and an optional body-keyword pass (Python-side).

    query_params keys:
        subject        – substring match (case-insensitive)
        sender         – substring match
        body_contains  – substring match (Python-side, slower)
        date_from      – datetime object
        date_to        – datetime object
    """
    ns = _ns()
    results = []

    subject       = (query_params.get("subject") or "").strip()
    sender        = (query_params.get("sender") or "").strip()
    body_contains = (query_params.get("body_contains") or "").strip()
    date_from     = query_params.get("date_from")
    date_to       = query_params.get("date_to")

    # Build Restrict filter (DASL / Jet-syntax that works in all locales)
    filter_parts = []
    if date_from:
        filter_parts.append(f"[ReceivedTime] >= '{date_from.strftime('%m/%d/%Y %I:%M %p')}'")
    if date_to:
        filter_parts.append(f"[ReceivedTime] <= '{date_to.strftime('%m/%d/%Y %I:%M %p')}'")
    if subject:
        safe = subject.replace("'", "''")
        filter_parts.append(f"[Subject] LIKE '%{safe}%'")
    if sender:
        safe = sender.replace("'", "''")
        filter_parts.append(f"[SenderEmailAddress] LIKE '%{safe}%'")

    filter_str = " AND ".join(filter_parts)

    # Walk every store / folder
    try:
        for store in ns.Stores:
            try:
                root = store.GetRootFolder()
                _search_folder_recursive(root, store.DisplayName, filter_str,
                                         body_contains, results)
            except Exception as exc:
                log.warning("Could not search store '%s': %s", store.DisplayName, exc)
    finally:
        pythoncom.CoUninitialize()

    return results


def _search_folder_recursive(folder_com, store_name, filter_str, body_contains, results):
    """Recursively search a COM folder tree, appending plain-dict results."""
    try:
        items = folder_com.Items
        if filter_str:
            try:
                items = items.Restrict(filter_str)
            except Exception:
                items = folder_com.Items

        for item in items:
            try:
                if getattr(item, "Class", None) != 43:   # 43 = olMail
                    continue

                if body_contains:
                    body_text = (getattr(item, "Body", "") or "")[:50_000]
                    if body_contains.lower() not in body_text.lower():
                        continue

                results.append({
                    "message_id":  getattr(item, "EntryID", "N/A"),
                    "subject":     getattr(item, "Subject", ""),
                    "sender":      getattr(item, "SenderEmailAddress", ""),
                    "received_at": getattr(item, "ReceivedTime", None),
                    "folder_path": folder_com.FolderPath,
                    "store":       store_name,
                    "body":        (getattr(item, "Body", "") or "")[:20_000],
                    "attachments": [att.FileName for att in item.Attachments],
                })
            except Exception:
                continue
    except Exception as exc:
        log.warning("Error searching folder '%s': %s", folder_com.Name, exc)

    # Recurse into sub-folders
    try:
        for subfolder in folder_com.Folders:
            _search_folder_recursive(subfolder, store_name, filter_str,
                                     body_contains, results)
    except Exception:
        pass


def extract_mails(folder_entry_id: str, limit: int = 200):
    """
    Given a folder EntryID string, extracts up to *limit* mails.
    Initialises COM itself so it is safe to call from a worker thread.
    Returns a list of plain dicts.
    """
    ns = _ns()
    mails = []
    try:
        folder_com = ns.GetFolderFromID(folder_entry_id)
        items = folder_com.Items
        items.Sort("[ReceivedTime]", True)   # newest first

        count = 0
        for item in items:
            if count >= limit:
                break
            try:
                if getattr(item, "Class", None) != 43:
                    continue
                mails.append({
                    "message_id":  getattr(item, "EntryID", "N/A"),
                    "subject":     getattr(item, "Subject", "No Subject"),
                    "sender":      getattr(item, "SenderEmailAddress", "Unknown"),
                    "body":        (getattr(item, "Body", "") or "")[:20_000],
                    "received_at": getattr(item, "ReceivedTime", None),
                    "folder_path": folder_com.FolderPath,
                    "store":       folder_com.Store.DisplayName,
                    "attachments": [att.FileName for att in item.Attachments],
                })
                count += 1
            except Exception:
                continue
    except Exception as exc:
        log.error("extract_mails error: %s", exc)
    finally:
        pythoncom.CoUninitialize()
    return mails


def get_folder_stats(folder_entry_id: str):
    """
    Returns lightweight stats for a single folder (no full body read).
    Used by the Statistics panel.
    """
    ns = _ns()
    stats = {"total": 0, "by_sender": {}, "by_month": {}}
    try:
        folder_com = ns.GetFolderFromID(folder_entry_id)
        for item in folder_com.Items:
            try:
                if getattr(item, "Class", None) != 43:
                    continue
                stats["total"] += 1
                sender = getattr(item, "SenderEmailAddress", "Unknown") or "Unknown"
                stats["by_sender"][sender] = stats["by_sender"].get(sender, 0) + 1
                rt = getattr(item, "ReceivedTime", None)
                if rt:
                    key = rt.strftime("%Y-%m") if hasattr(rt, "strftime") else str(rt)[:7]
                    stats["by_month"][key] = stats["by_month"].get(key, 0) + 1
            except Exception:
                continue
    except Exception as exc:
        log.error("get_folder_stats error: %s", exc)
    finally:
        pythoncom.CoUninitialize()
    return stats