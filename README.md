# Microsoft Outlook MCP Server for Claude & Antigravity (AGY)

[![PyPI](https://img.shields.io/pypi/v/mcp-outlook-desktop.svg)](https://pypi.org/project/mcp-outlook-desktop/)
[![CI](https://github.com/Jpatching/mcp-outlook/actions/workflows/ci.yml/badge.svg)](https://github.com/Jpatching/mcp-outlook/actions)

A high-performance **Model Context Protocol (MCP)** server providing direct, secure access to your **Microsoft 365 Outlook** emails, calendar, and contacts from:
- **Claude Code CLI**
- **Claude Desktop**
- **Antigravity CLI (AGY)**

---


## 🖥️ Live Terminal Demo

```ansi
[1;36m>[0m [1mclaude[0m "Check my unread emails and prepare a draft follow-up to Alex about tomorrow's sync"

[1;34m⚡ Calling Tool:[0m outlook_list_messages(folder="inbox", unread_only=True, top=5)
[32m✔ Found 2 unread messages:[0m
  [1] [1mFrom:[0m Alex Mercer <alex@company.com> | [1mSubject:[0m Agenda for Q3 Planning Sync
  [2] [1mFrom:[0m Cloud Alerts <alerts@infra.io> | [1mSubject:[0m Weekly Node Health Report

[1;34m⚡ Calling Tool:[0m outlook_get_message(message_id="MSG_049281")
[32m✔ Retrieved message safely encapsulated in <untrusted_email_body_content> boundary.[0m

[1;34m⚡ Calling Tool:[0m outlook_create_draft(
    to_recipients=["alex@company.com"],
    subject="Re: Agenda for Q3 Planning Sync",
    body="Hi Alex,\n\nThanks for sending over the agenda. I reviewed the distributed queue items and added notes for tomorrow's sync.\n\nBest,\nJosh"
)
[1;32m✔ Draft created successfully in Outlook Drafts folder (ID: DRAFT_839120)[0m
[2;37m  [Human-in-the-Loop Safe]: The draft is ready in your Outlook app for review before sending.[0m
```

---

## 🔒 Privacy & Security Architecture

- **Human-in-the-Loop by Design**: All outbound message drafting defaults to your Outlook `Drafts` folder for manual review. Direct unreviewed sends (`outlook_send_mail`) require an explicit `confirm_send=True` flag to prevent accidental dispatches.
- **Indirect Prompt Injection Defense**: Incoming email bodies are encapsulated in `<untrusted_email_body_content>` XML tags to ensure AI models parse external content strictly as passive data.
- **Zero Cloud Tokens / No Passwords**: Connects 100% locally to your running Windows Outlook Desktop application via Windows MAPI (`win32com`).
- **Strict `.gitignore`**: All local configs, caches, and environments are ignored.

---

## 🚀 Installation & Quickstart

### Method 1: Instant Run with `uvx` (No git clone needed)

If you have `uv` installed, you can run the server directly:

```cmd
uvx mcp-outlook-desktop
```

#### Claude Desktop Configuration (`%APPDATA%\Claude\claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["mcp-outlook"]
    }
  }
}
```

#### Claude Code CLI:
```cmd
claude mcp add outlook uvx mcp-outlook-desktop
```

---

### Method 2: Local Installation with `pip`

```cmd
pip install mcp-outlook-desktop
```

Then in your Claude Desktop config:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "mcp-outlook"
    }
  }
}
```

---

### Method 3: Clone from Source (`C:\projects\mcp-outlook`)

```cmd
cd C:\projects
gh repo clone Jpatching/mcp-outlook
cd mcp-outlook
setup.bat
```

Test connection (with Outlook open):
```cmd
run.bat --status
run.bat --test
```

---

## 🛠️ Available MCP Tools

| Tool Name | Parameters | Description |
| :--- | :--- | :--- |
| `outlook_check_status` | *none* | Check connection to Outlook |
| `outlook_list_messages` | `folder`, `top`, `unread_only` | List recent emails from Inbox, Sent, Drafts |
| `outlook_search_messages` | `query`, `top` | Search messages across mailbox |
| `outlook_get_message` | `message_id` | Retrieve full email body and sender details safely encapsulated in XML boundaries |
| `outlook_create_draft` | `subject`, `body`, `to_recipients`, `reply_to_message_id` | Create a draft email or reply draft in Outlook Drafts folder (Recommended) |
| `outlook_send_mail` | `subject`, `body`, `to_recipients`, `confirm_send` | Send an email directly (diverts to Drafts unless `confirm_send=True`) |
| `outlook_list_calendar_events`| `days_ahead` | List upcoming meetings and appointments |
| `outlook_create_calendar_event`| `subject`, `start_datetime`, `end_datetime`, `attendees`, `location` | Schedule a meeting in Outlook Calendar |
| `outlook_list_folders` | *none* | List mail folders with unread counts |
