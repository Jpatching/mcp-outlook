import re
import httpx
from datetime import datetime, timezone, timedelta
from typing import Optional, List, Dict, Any
from auth import OutlookAuth

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

def clean_html_to_text(html_content: str) -> str:
    """Converts email HTML to clean plain text/markdown for token efficiency."""
    if not html_content:
        return ""
    # Strip script and style tags
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html_content, flags=re.DOTALL | re.IGNORECASE)
    # Convert <br>, <p>, <div> to newlines
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(p|div|tr|li)[^>]*>", "\n", text, flags=re.IGNORECASE)
    # Strip remaining HTML tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common HTML entities
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", "\"").replace("&#39;", "'")
    # Compress multiple newlines and spaces
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

class GraphClient:
    def __init__(self, auth: OutlookAuth):
        self.auth = auth

    def _get_headers(self) -> Dict[str, str]:
        token = self.auth.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Prefer": 'outlook.body-content-type="text"'
        }

    async def get_me(self) -> Dict[str, Any]:
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.get(f"{GRAPH_BASE_URL}/me", headers=self._get_headers())
            resp.raise_for_status()
            return resp.json()

    async def list_folders(self) -> List[Dict[str, Any]]:
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.get(
                f"{GRAPH_BASE_URL}/me/mailFolders?$top=50&$select=id,displayName,unreadItemCount,totalItemCount",
                headers=self._get_headers()
            )
            resp.raise_for_status()
            return resp.json().get("value", [])

    async def list_messages(
        self,
        folder: str = "inbox",
        top: int = 10,
        unread_only: bool = False
    ) -> List[Dict[str, Any]]:
        params = {
            "$top": min(top, 50),
            "$select": "id,subject,receivedDateTime,from,isRead,hasAttachments,bodyPreview,importance",
            "$orderby": "receivedDateTime desc"
        }
        if unread_only:
            params["$filter"] = "isRead eq false"

        endpoint = f"{GRAPH_BASE_URL}/me/mailFolders/{folder}/messages" if folder else f"{GRAPH_BASE_URL}/me/messages"
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.get(endpoint, headers=self._get_headers(), params=params)
            resp.raise_for_status()
            items = resp.json().get("value", [])
            
            # Format cleanly
            results = []
            for m in items:
                sender = m.get("from", {}).get("emailAddress", {})
                results.append({
                    "id": m.get("id"),
                    "subject": m.get("subject") or "(No Subject)",
                    "from": f"{sender.get('name', '')} <{sender.get('address', '')}>",
                    "received": m.get("receivedDateTime"),
                    "isRead": m.get("isRead"),
                    "importance": m.get("importance"),
                    "hasAttachments": m.get("hasAttachments"),
                    "preview": m.get("bodyPreview")
                })
            return results

    async def search_messages(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        params = {
            "$search": f'"{query}"',
            "$top": min(top, 50),
            "$select": "id,subject,receivedDateTime,from,isRead,hasAttachments,bodyPreview"
        }
        async with httpx.AsyncClient(timeout=25.0) as client:
            resp = await client.get(f"{GRAPH_BASE_URL}/me/messages", headers=self._get_headers(), params=params)
            resp.raise_for_status()
            items = resp.json().get("value", [])
            
            results = []
            for m in items:
                sender = m.get("from", {}).get("emailAddress", {})
                results.append({
                    "id": m.get("id"),
                    "subject": m.get("subject") or "(No Subject)",
                    "from": f"{sender.get('name', '')} <{sender.get('address', '')}>",
                    "received": m.get("receivedDateTime"),
                    "isRead": m.get("isRead"),
                    "preview": m.get("bodyPreview")
                })
            return results

    async def get_message(self, message_id: str) -> Dict[str, Any]:
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.get(
                f"{GRAPH_BASE_URL}/me/messages/{message_id}?$select=id,subject,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients,isRead,importance,hasAttachments,body",
                headers=self._get_headers()
            )
            resp.raise_for_status()
            m = resp.json()
            sender = m.get("from", {}).get("emailAddress", {})
            to_list = [f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>" for r in m.get("toRecipients", [])]
            cc_list = [f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>" for r in m.get("ccRecipients", [])]
            
            raw_body = m.get("body", {}).get("content", "")
            body_type = m.get("body", {}).get("contentType", "text")
            clean_body = clean_html_to_text(raw_body) if body_type.lower() == "html" else raw_body

            return {
                "id": m.get("id"),
                "subject": m.get("subject") or "(No Subject)",
                "from": f"{sender.get('name', '')} <{sender.get('address', '')}>",
                "to": to_list,
                "cc": cc_list,
                "received": m.get("receivedDateTime"),
                "importance": m.get("importance"),
                "hasAttachments": m.get("hasAttachments"),
                "body": clean_body
            }

    async def create_draft(
        self,
        subject: str,
        body: str,
        to_recipients: List[str],
        cc_recipients: Optional[List[str]] = None,
        reply_to_message_id: Optional[str] = None
    ) -> Dict[str, Any]:
        headers = self._get_headers()
        async with httpx.AsyncClient(timeout=20.0) as client:
            if reply_to_message_id:
                # Create draft reply to thread
                reply_resp = await client.post(
                    f"{GRAPH_BASE_URL}/me/messages/{reply_to_message_id}/createReply",
                    headers=headers
                )
                reply_resp.raise_for_status()
                draft = reply_resp.json()
                draft_id = draft["id"]
                
                # Update body/comment of the draft reply
                patch_payload = {
                    "body": {
                        "contentType": "Text",
                        "content": body
                    }
                }
                patch_resp = await client.patch(
                    f"{GRAPH_BASE_URL}/me/messages/{draft_id}",
                    headers=headers,
                    json=patch_payload
                )
                patch_resp.raise_for_status()
                return {
                    "status": "success",
                    "action": "reply_draft_created",
                    "draft_id": draft_id,
                    "subject": draft.get("subject"),
                    "message": "Draft reply successfully created in Outlook Drafts folder. You can review and send it."
                }
            else:
                # Standalone draft
                payload = {
                    "subject": subject,
                    "body": {
                        "contentType": "Text",
                        "content": body
                    },
                    "toRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in to_recipients],
                    "ccRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in (cc_recipients or [])]
                }
                resp = await client.post(
                    f"{GRAPH_BASE_URL}/me/messages",
                    headers=headers,
                    json=payload
                )
                resp.raise_for_status()
                draft = resp.json()
                return {
                    "status": "success",
                    "action": "draft_created",
                    "draft_id": draft.get("id"),
                    "subject": draft.get("subject"),
                    "message": "Draft email created in Outlook Drafts folder."
                }

    async def send_mail(
        self,
        subject: str,
        body: str,
        to_recipients: List[str],
        cc_recipients: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in to_recipients],
                "ccRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in (cc_recipients or [])]
            },
            "saveToSentItems": "true"
        }
        async with httpx.AsyncClient(timeout=25.0) as client:
            resp = await client.post(
                f"{GRAPH_BASE_URL}/me/sendMail",
                headers=self._get_headers(),
                json=payload
            )
            resp.raise_for_status()
            return {"status": "success", "message": f"Email successfully sent to {', '.join(to_recipients)}."}

    async def list_calendar_events(
        self,
        days_ahead: int = 7,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        now = datetime.now(timezone.utc)
        if not start_datetime:
            start_iso = now.isoformat()
        else:
            start_iso = start_datetime

        if not end_datetime:
            end_iso = (now + timedelta(days=days_ahead)).isoformat()
        else:
            end_iso = end_datetime

        params = {
            "startDateTime": start_iso,
            "endDateTime": end_iso,
            "$top": min(top, 50),
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,onlineMeeting"
        }

        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.get(
                f"{GRAPH_BASE_URL}/me/calendarView",
                headers=self._get_headers(),
                params=params
            )
            resp.raise_for_status()
            items = resp.json().get("value", [])
            
            results = []
            for ev in items:
                organizer = ev.get("organizer", {}).get("emailAddress", {})
                attendees = [a.get("emailAddress", {}).get("name") or a.get("emailAddress", {}).get("address") for a in ev.get("attendees", [])]
                results.append({
                    "id": ev.get("id"),
                    "subject": ev.get("subject") or "(No Title)",
                    "start": ev.get("start", {}).get("dateTime"),
                    "end": ev.get("end", {}).get("dateTime"),
                    "timeZone": ev.get("start", {}).get("timeZone"),
                    "isAllDay": ev.get("isAllDay"),
                    "isCancelled": ev.get("isCancelled"),
                    "location": ev.get("location", {}).get("displayName"),
                    "organizer": f"{organizer.get('name', '')} <{organizer.get('address', '')}>",
                    "attendees": attendees,
                    "onlineMeetingUrl": ev.get("onlineMeeting", {}).get("joinUrl") if ev.get("onlineMeeting") else None
                })
            return results

    async def create_calendar_event(
        self,
        subject: str,
        start_datetime: str,
        end_datetime: str,
        attendees: Optional[List[str]] = None,
        location: Optional[str] = None,
        body: Optional[str] = None,
        is_online_meeting: bool = False
    ) -> Dict[str, Any]:
        payload: Dict[str, Any] = {
            "subject": subject,
            "start": {
                "dateTime": start_datetime,
                "timeZone": "UTC"
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": "UTC"
            },
            "isOnlineMeeting": is_online_meeting
        }
        if attendees:
            payload["attendees"] = [
                {"emailAddress": {"address": a.strip()}, "type": "required"} for a in attendees
            ]
        if location:
            payload["location"] = {"displayName": location}
        if body:
            payload["body"] = {"contentType": "Text", "content": body}

        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{GRAPH_BASE_URL}/me/events",
                headers=self._get_headers(),
                json=payload
            )
            resp.raise_for_status()
            ev = resp.json()
            return {
                "status": "success",
                "id": ev.get("id"),
                "subject": ev.get("subject"),
                "start": ev.get("start", {}).get("dateTime"),
                "end": ev.get("end", {}).get("dateTime"),
                "message": "Calendar event successfully created in Outlook."
            }
