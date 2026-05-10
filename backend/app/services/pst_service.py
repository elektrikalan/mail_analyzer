"""
pst_service.py
--------------
Reads PST / OST files directly using pypff (libpff-python).
No Outlook installation required.
"""
import os
import logging
import datetime

log = logging.getLogger(__name__)

try:
    import pypff
    PYPFF_AVAILABLE = True
except ImportError:
    PYPFF_AVAILABLE = False
    log.warning("pypff not installed. PST file mode unavailable.")


def is_pypff_available() -> bool:
    return PYPFF_AVAILABLE


def _folder_to_node(folder, store_name: str, pst_path: str, index_path: list) -> dict:
    """Recursively convert a pypff folder to a plain dict tree node."""
    try:
        msg_count = folder.get_number_of_sub_messages()
    except Exception:
        msg_count = 0

    node = {
        "store":        store_name,
        "name":         folder.name or "Root",
        "full_path":    f"\\\\{store_name}\\" + "\\".join(str(i) for i in index_path),
        "entry_id":     None,   # not used in PST mode
        "item_count":   msg_count,
        "children":     [],
        # PST-specific navigation info
        "_pst_path":    pst_path,
        "_index_path":  index_path[:],   # copy
    }

    try:
        for i in range(folder.get_number_of_sub_folders()):
            sub = folder.get_sub_folder(i)
            child = _folder_to_node(sub, store_name, pst_path, index_path + [i])
            node["children"].append(child)
    except Exception as exc:
        log.warning("Error reading sub-folders: %s", exc)

    return node


def read_pst_tree(pst_path: str) -> dict:
    """Open a PST file and return a nested folder-tree dict."""
    if not PYPFF_AVAILABLE:
        raise RuntimeError("pypff is not installed. Run: pip install libpff-python")
    pf = pypff.file()
    pf.open(pst_path)
    try:
        root = pf.get_root_folder()
        store_name = os.path.basename(pst_path)
        tree = _folder_to_node(root, store_name, pst_path, [])
    finally:
        pf.close()
    return tree


def _navigate_to_folder(pf_root, index_path: list):
    """Navigate from root to a sub-folder using an index path list."""
    folder = pf_root
    for idx in index_path:
        folder = folder.get_sub_folder(idx)
    return folder


def _parse_message(msg, folder_path: str, store_name: str) -> dict:
    """Convert a pypff message object to a plain dict."""
    try:
        body_bytes = msg.get_plain_text_body()
        body = body_bytes.decode("utf-8", errors="replace") if body_bytes else ""
    except Exception:
        body = ""

    dt = None
    try:
        raw = msg.delivery_time
        if raw:
            # pypff returns a datetime-like object
            dt = datetime.datetime(raw.year, raw.month, raw.day,
                                   raw.hours, raw.minutes, raw.seconds)
    except Exception:
        pass

    return {
        "message_id":  str(getattr(msg, "identifier", id(msg))),
        "subject":     getattr(msg, "subject", "") or "",
        "sender":      getattr(msg, "sender_name", "") or "",
        "body":        body[:20_000],
        "received_at": dt,
        "folder_path": folder_path,
        "store":       store_name,
        "attachments": [],
    }


def extract_mails_from_pst(pst_path: str, index_path: list, limit: int = 300) -> list:
    """Extract mails from a PST folder identified by index_path."""
    if not PYPFF_AVAILABLE:
        raise RuntimeError("pypff is not installed.")
    pf = pypff.file()
    pf.open(pst_path)
    mails = []
    store_name = os.path.basename(pst_path)
    try:
        root = pf.get_root_folder()
        folder = _navigate_to_folder(root, index_path)
        folder_path = f"\\\\{store_name}\\" + "\\".join(str(i) for i in index_path)
        count = min(folder.get_number_of_sub_messages(), limit)
        for i in range(count):
            try:
                msg = folder.get_sub_message(i)
                mails.append(_parse_message(msg, folder_path, store_name))
            except Exception:
                continue
    finally:
        pf.close()
    return mails


def search_pst(pst_path: str, query_params: dict) -> list:
    """Search through all messages in a PST file."""
    if not PYPFF_AVAILABLE:
        raise RuntimeError("pypff is not installed.")

    subject_q  = (query_params.get("subject") or "").strip().lower()
    sender_q   = (query_params.get("sender") or "").strip().lower()
    body_q     = (query_params.get("body_contains") or "").strip().lower()

    pf = pypff.file()
    pf.open(pst_path)
    results = []
    store_name = os.path.basename(pst_path)
    try:
        root = pf.get_root_folder()
        _search_folder_pst(root, store_name, pst_path, [],
                           subject_q, sender_q, body_q, results)
    finally:
        pf.close()
    return results


def _search_folder_pst(folder, store_name, pst_path, index_path,
                        subject_q, sender_q, body_q, results):
    folder_path = f"\\\\{store_name}\\" + "\\".join(str(i) for i in index_path)
    try:
        for i in range(folder.get_number_of_sub_messages()):
            try:
                msg = folder.get_sub_message(i)
                subj   = (getattr(msg, "subject", "") or "").lower()
                sender = (getattr(msg, "sender_name", "") or "").lower()
                if subject_q and subject_q not in subj:
                    continue
                if sender_q and sender_q not in sender:
                    continue
                if body_q:
                    try:
                        b = msg.get_plain_text_body()
                        body_text = b.decode("utf-8", errors="replace") if b else ""
                    except Exception:
                        body_text = ""
                    if body_q not in body_text.lower():
                        continue
                results.append(_parse_message(msg, folder_path, store_name))
            except Exception:
                continue
    except Exception:
        pass

    try:
        for i in range(folder.get_number_of_sub_folders()):
            _search_folder_pst(folder.get_sub_folder(i), store_name, pst_path,
                               index_path + [i], subject_q, sender_q, body_q, results)
    except Exception:
        pass
