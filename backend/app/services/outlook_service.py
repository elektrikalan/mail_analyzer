import win32com.client
import pythoncom
from datetime import datetime

def get_outlook_namespace():
    # Use explicit apartment threaded initialization for COM
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    try:
        # Use DispatchEx for better isolation in threads
        outlook = win32com.client.DispatchEx("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Try to force a clean logon session
        namespace.Logon(None, None, False, True)
        return namespace
    except Exception as e:
        # Fallback to standard Dispatch if DispatchEx fails
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        try:
            namespace.Logon(None, None, False, False)
        except Exception as logon_error:
            print(f"Warning: Outlook logon failed: {logon_error}")
            pass
        return namespace

def _get_folders_recursive(parent_folder, store_name, folder_list):
    """
    Recursively finds all folders within a parent folder.
    """
    for folder in parent_folder.Folders:
        try:
            folder_list.append({
                "store": store_name,
                "name": folder.Name,
                "full_path": folder.FolderPath,
                "entry_id": folder.EntryID,
                "item_count": folder.Items.Count,
                "obj": folder
            })
            if folder.Folders.Count > 0:
                _get_folders_recursive(folder, store_name, folder_list)
        except Exception as e:
            print(f"Error accessing subfolder {folder.Name}: {e}")
            continue

def enumerate_all_folders():
    """
    Enumerates every single folder in all Stores (Accounts, PSTs, Archives).
    """
    namespace = get_outlook_namespace()
    all_folders = []
    
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            # Root folder itself
            all_folders.append({
                "store": store.DisplayName,
                "name": root.Name,
                "full_path": root.FolderPath,
                "entry_id": root.EntryID,
                "item_count": root.Items.Count,
                "obj": root
            })
            # Start recursion
            _get_folders_recursive(root, store.DisplayName, all_folders)
        except Exception as e:
            print(f"Error accessing store {store.DisplayName}: {e}")
            continue
            
    return all_folders

def search_mails(query_params: dict):
    """
    Performs a filtered search across a specific folder or all folders.
    query_params can include: subject, sender, body_contains, date_from, date_to
    """
    namespace = get_outlook_namespace()
    results = []

    subject = (query_params.get("subject") or "").strip()
    sender = (query_params.get("sender") or "").strip()
    body_contains = (query_params.get("body_contains") or "").strip()

    filter_parts = []
    if query_params.get("date_from"):
        dt_str = query_params["date_from"].strftime("%m/%d/%Y %I:%M %p")
        filter_parts.append(f"[ReceivedTime] >= '{dt_str}'")

    if query_params.get("date_to"):
        dt_to_str = query_params["date_to"].strftime("%m/%d/%Y %I:%M %p")
        filter_parts.append(f"[ReceivedTime] <= '{dt_to_str}'")

    if subject:
        safe_subject = subject.replace("'", "''")
        filter_parts.append(f"[Subject] LIKE '%{safe_subject}%'")

    if sender:
        safe_sender = sender.replace("'", "''")
        filter_parts.append(f"[SenderEmailAddress] LIKE '%{safe_sender}%'")

    filter_str = " AND ".join(filter_parts) if filter_parts else ""
    folders = enumerate_all_folders()

    for folder_info in folders:
        folder_obj = folder_info["obj"]
        try:
            items = folder_obj.Items
            if filter_str:
                try:
                    items = items.Restrict(filter_str)
                except Exception as restrict_error:
                    print(f"Restrict filter failed for folder {folder_info['name']}: {restrict_error}")
                    items = folder_obj.Items

            count = 0
            for item in items:
                if count > 200:
                    break

                if getattr(item, "Class", None) != 43:
                    continue

                if body_contains:
                    body_text = getattr(item, "Body", "") or ""
                    if len(body_text) > 10000:  # Limit body size for performance
                        body_text = body_text[:10000]
                    if body_contains.lower() not in body_text.lower():
                        continue

                results.append({
                    "message_id": getattr(item, "EntryID", "N/A"),
                    "subject": getattr(item, "Subject", ""),
                    "sender": getattr(item, "SenderEmailAddress", ""),
                    "received_at": getattr(item, "ReceivedTime", None),
                    "folder_path": folder_info["full_path"],
                    "store": folder_info["store"],
                    "body": getattr(item, "Body", ""),
                    "attachments": [att.FileName for att in item.Attachments]
                })
                count += 1
        except Exception as e:
            print(f"Error searching folder {folder_info['name']}: {e}")
            continue

    return results

def extract_mails(folder_obj):
    """
    Takes an Outlook Folder COM object and extracts mail data.
    """
    pythoncom.CoInitialize()
    mails = []
    
    try:
        items = folder_obj.Items
        # Sort by received time descending
        items.Sort("[ReceivedTime]", True)
        
        count = 0
        for item in items:
            if count > 100: break # Safety limit
            
            try:
                if item.Class != 43: continue
                    
                mail_data = {
                    "message_id": getattr(item, "EntryID", "N/A"),
                    "subject": getattr(item, "Subject", "No Subject"),
                    "sender": getattr(item, "SenderEmailAddress", "Unknown"),
                    "body": getattr(item, "Body", ""),
                    "received_at": item.ReceivedTime,
                    "folder_path": folder_obj.FolderPath,
                    "attachments": [att.FileName for att in item.Attachments]
                }
                mails.append(mail_data)
                count += 1
            except Exception as e:
                continue
    except Exception as e:
        print(f"Error accessing folder items: {e}")
        
    return mails

def get_folder_by_id(entry_id):
    namespace = get_outlook_namespace()
    try:
        return namespace.GetFolderFromID(entry_id)
    except Exception as folder_error:
        print(f"Warning: Could not access folder by ID {entry_id}: {folder_error}")
        return None