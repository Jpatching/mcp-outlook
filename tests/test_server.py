import pytest
import json
from unittest.mock import AsyncMock, patch
import server

@pytest.mark.asyncio
async def test_outlook_list_messages():
    mock_messages = [
        {
            "id": "MSG1",
            "subject": "Q3 Planning",
            "from": "alice@example.com",
            "received": "2026-08-19 09:30",
            "isRead": True,
            "hasAttachments": False,
            "preview": "Let's review the Q3 priorities."
        },
        {
            "id": "MSG2",
            "subject": "Incident #404",
            "from": "alerts@example.com",
            "received": "2026-08-19 08:15",
            "isRead": False,
            "hasAttachments": True,
            "preview": "Critical alert fired on worker-01."
        }
    ]

    with patch.object(server.cloud_client, "list_messages", new_callable=AsyncMock) as mock_list:
        mock_list.return_value = mock_messages
        with patch("server.is_local_mode", return_value=False):
            res = await server.outlook_list_messages(folder="inbox", top=10)
            assert "Found 2 messages in 'inbox'" in res
            assert "Q3 Planning" in res
            assert "[UNREAD] Incident #404 📎" in res

@pytest.mark.asyncio
async def test_outlook_create_draft():
    with patch.object(server.cloud_client, "create_draft", new_callable=AsyncMock) as mock_draft:
        mock_draft.return_value = {"status": "success", "draft_id": "DRAFT99"}
        with patch("server.is_local_mode", return_value=False):
            res_str = await server.outlook_create_draft(
                subject="Test Draft",
                body="Draft content",
                to_recipients=["test@example.com"]
            )
            res = json.loads(res_str)
            assert res["status"] == "success"
            assert res["draft_id"] == "DRAFT99"

@pytest.mark.asyncio
async def test_outlook_list_calendar_events():
    mock_events = [
        {
            "id": "EV1",
            "subject": "Architecture Sync",
            "start": "2026-08-20T10:00:00Z",
            "end": "2026-08-20T10:30:00Z",
            "location": "Virtual",
            "organizer": "lead@example.com",
            "attendees": ["dev1@example.com", "dev2@example.com"],
            "onlineMeetingUrl": "https://teams.microsoft.com/meet/123",
            "isCancelled": False
        }
    ]
    with patch.object(server.cloud_client, "list_calendar_events", new_callable=AsyncMock) as mock_cal:
        mock_cal.return_value = mock_events
        with patch("server.is_local_mode", return_value=False):
            res = await server.outlook_list_calendar_events(days_ahead=7)
            assert "Architecture Sync" in res
            assert "https://teams.microsoft.com/meet/123" in res
