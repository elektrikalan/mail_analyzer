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
        except:
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
    
    # Restrict string builder for Outlook
    # Example: "[ReceivedTime] >= '01/01/2023 12:00 AM' AND [Subject] = 'Test'"
    filter_parts = []
    
    if query_params.get("date_from"):
        # Format must be MM/DD/YYYY HH:MM AM/PM
        dt_str = query_params["date_from"].strftime("%m/%d/%Y %I:%M %p")
        filter_parts.append(f"[ReceivedTime] >= '{dt_str}'")
        
    if query_params.get("subject"):
        filter_parts.append(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{query_params['subject']}%'")

    filter_str = " AND ".join(filter_parts) if filter_parts else ""

    folders = enumerate_all_folders()
    
    for folder_info in folders:
        folder_obj = folder_info["obj"]
        try:
            items = folder_obj.Items
            if filter_str:
                items = items.Restrict(filter_str)
            
            # Additional manual filtering for fields that Restrict might not handle well
            count = 0
            for item in items:
                if count > 200: break # Safety limit per folder
                
                if item.Class != 43: continue # MailItem
                
                # Check body_contains manually if provided (Restrict on Body can be slow/unreliable)
                if query_params.get("body_contains"):
                    if query_params["body_contains"].lower() not in item.Body.lower():
                        continue

                results.append({
                    "message_id": item.EntryID,
                    "subject": item.Subject,
                    "sender": item.SenderEmailAddress,
                    "received_at": item.ReceivedTime,
                    "folder_path": folder_info["full_path"],
                    "store": folder_info["store"],
                    "body": item.Body,
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
    except:
        return None