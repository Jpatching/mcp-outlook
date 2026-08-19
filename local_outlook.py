"""
Windows Local Outlook Client Bridge via MAPI / COM Automation.
Requires NO Azure App Registration, NO Tenant Permissions, and NO Admin Consent.
Works directly with the desktop Outlook application installed and logged in on Windows.
"""

import sys
from datetime import datetime, timezone, timedelta
from typing import Optional, List, Dict, Any

class WindowsLocalOutlook:
    def __init__(self):
        if sys.platform != "win32":
            raise RuntimeError("Windows Local Outlook COM automation is only available on Windows.")
        try:
            import win32com.client
            self.win32com = win32com.client
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.app.GetNamespace("MAPI")
        except Exception as e:
            raise RuntimeError(f"Failed to connect to local Outlook application: {e}. Is Outlook installed and pywin32 installed?")

    def check_status(self) -> Dict[str, Any]:
        try:
            current_user = self.namespace.CurrentUser.Name
            accounts = [self.namespace.Accounts.Item(i).DisplayName for i in range(1, self.namespace.Accounts.Count + 1)]
            return {
                "available": True,
                "user": current_user,
                "accounts": accounts,
                "mode": "Windows Local MAPI/COM (Zero Cloud Config)"
            }
        except Exception as e:
            return {"available": False, "error": str(e)}

    def list_messages(self, folder_name: str = "inbox", top: int = 10, unread_only: bool = False) -> List[Dict[str, Any]]:
        # Map folder names to MAPI OlDefaultFolders enum
        folder_map = {
            "inbox": 6,
            "drafts": 16,
            "sentitems": 5,
            "sent": 5,
            "deleted": 3,
            "junk": 23,
            "outbox": 4
        }
        folder_idx = folder_map.get(folder_name.lower(), 6)
        folder = self.namespace.GetDefaultFolder(folder_idx)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort descending

        results = []
        count = 0
        for item in items:
            if count >= top:
                break
            try:
                # Class 43 = olMail
                if getattr(item, "Class", 0) != 43:
                    continue
                if unread_only and not getattr(item, "UnRead", False):
                    continue

                entry_id = getattr(item, "EntryID", "")
                subject = getattr(item, "Subject", "(No Subject)")
                sender = getattr(item, "SenderName", "")
                received = getattr(item, "ReceivedTime", None)
                unread = getattr(item, "UnRead", False)
                body = getattr(item, "Body", "")[:120]

                results.append({
                    "id": entry_id,
                    "subject": subject,
                    "from": sender,
                    "received": str(received) if received else "",
                    "isRead": not unread,
                    "preview": body
                })
                count += 1
            except Exception:
                continue
        return results

    def get_message(self, entry_id: str) -> Dict[str, Any]:
        item = self.namespace.GetItemFromID(entry_id)
        return {
            "id": entry_id,
            "subject": getattr(item, "Subject", "(No Subject)"),
            "from": f"{getattr(item, 'SenderName', '')} <{getattr(item, 'SenderEmailAddress', '')}>",
            "to": [getattr(item, "To", "")],
            "cc": [getattr(item, "CC", "")] if getattr(item, "CC", "") else [],
            "received": str(getattr(item, "ReceivedTime", "")),
            "body": getattr(item, "Body", "")
        }

    def search_messages(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        inbox = self.namespace.GetDefaultFolder(6)
        # Search via Jet or DASL or item filter
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        
        results = []
        q_lower = query.lower()
        count = 0
        for item in items:
            if count >= top:
                break
            try:
                if getattr(item, "Class", 0) != 43:
                    continue
                subject = getattr(item, "Subject", "")
                body = getattr(item, "Body", "")
                sender = getattr(item, "SenderName", "")

                if q_lower in subject.lower() or q_lower in body.lower() or q_lower in sender.lower():
                    results.append({
                        "id": getattr(item, "EntryID", ""),
                        "subject": subject,
                        "from": sender,
                        "received": str(getattr(item, "ReceivedTime", "")),
                        "isRead": not getattr(item, "UnRead", False),
                        "preview": body[:120]
                    })
                    count += 1
            except Exception:
                continue
        return results

    def create_draft(self, subject: str, body: str, to_recipients: List[str], cc_recipients: Optional[List[str]] = None, reply_to_entry_id: Optional[str] = None) -> Dict[str, Any]:
        if reply_to_entry_id:
            orig_mail = self.namespace.GetItemFromID(reply_to_entry_id)
            reply = orig_mail.CreateReply()
            reply.Body = body + "\n\n" + getattr(reply, "Body", "")
            reply.Save()
            return {
                "status": "success",
                "action": "reply_draft_created",
                "draft_id": getattr(reply, "EntryID", ""),
                "subject": getattr(reply, "Subject", ""),
                "message": "Draft reply created in Outlook Drafts folder."
            }
        else:
            mail = self.app.CreateItem(0)  # 0 = olMailItem
            mail.Subject = subject
            mail.Body = body
            mail.To = "; ".join(to_recipients)
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
            mail.Save()
            return {
                "status": "success",
                "action": "draft_created",
                "draft_id": getattr(mail, "EntryID", ""),
                "subject": subject,
                "message": "Draft email created in Outlook Drafts folder."
            }

    def list_calendar_events(self, days_ahead: int = 7) -> List[Dict[str, Any]]:
        calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        now = datetime.now()
        end_time = now + timedelta(days=days_ahead)
        date_filter = f"[Start] >= '{now.strftime('%m/%d/%Y %I:%M %p')}' AND [Start] <= '{end_time.strftime('%m/%d/%Y %I:%M %p')}'"
        
        filtered = items.Restrict(date_filter)
        results = []
        for item in filtered:
            try:
                results.append({
                    "id": getattr(item, "EntryID", ""),
                    "subject": getattr(item, "Subject", "(No Title)"),
                    "start": str(getattr(item, "Start", "")),
                    "end": str(getattr(item, "End", "")),
                    "location": getattr(item, "Location", ""),
                    "organizer": getattr(item, "Organizer", ""),
                    "attendees": [getattr(item, "RequiredAttendees", "")]
                })
            except Exception:
                continue
        return results
