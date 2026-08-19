# Microsoft Outlook MCP Server for Claude & Antigravity (AGY)

A unified **Model Context Protocol (MCP)** server providing direct, secure access to your **Work Microsoft 365 Outlook** emails, calendar, and contacts from:
- **Claude Code CLI**
- **Claude Desktop** (Windows & Linux)
- **Antigravity CLI (AGY)**

---

## 🏆 Recommended Solution for Work M365 Tenants: Windows Local Outlook Desktop Bridge

In strict corporate/enterprise tenants where Azure App Registrations and Cloud Device Code flows are restricted by tenant policies (e.g. `AADSTS65002` / Conditional Access), **the Windows Local Desktop Bridge is the simplest, most powerful solution**:

### Why it is the best:
- ✅ **Zero Azure Registration & Zero Admin Consent**: No need to ask IT or register apps in Entra ID.
- ✅ **100% Local & Secure**: Communicates directly with the running Outlook Desktop application on your Windows machine via Windows MAPI (`win32com`).
- ✅ **Instant Access**: Reads your authenticated work inbox, sent items, drafts, and calendar automatically.

---

## 🚀 Windows Setup (Takes 30 Seconds)

1. Copy the `mcp-outlook` folder to your Windows machine (e.g., `C:\Users\<YourUser>\projects\mcp-outlook`).
2. Open PowerShell or Command Prompt in that directory:
   ```cmd
   cd C:\Users\<YourUser>\projects\mcp-outlook
   setup.bat
   ```
3. Test your connection while Outlook is open:
   ```cmd
   run.bat --status
   run.bat --test
   ```

---

## ⚙️ Connecting to Claude & Antigravity

### 1. Claude Desktop (Windows)
Open `%APPDATA%\Claude\claude_desktop_config.json` and add:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\<YourUser>\\projects\\mcp-outlook\\run.bat"
    }
  }
}
```

### 2. Claude Code CLI (Windows)
```bash
claude mcp add outlook C:\Users\<YourUser>\projects\mcp-outlook\run.bat
```

### 3. Antigravity CLI (Windows)
In `%USERPROFILE%\.gemini\config\mcp_config.json`:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\<YourUser>\\projects\\mcp-outlook\\run.bat",
      "args": []
    }
  }
}
```

---

## 🌐 Optional: Connecting Linux Claude/AGY to Windows Outlook via Network (SSE)

If you work on Linux and want Claude/AGY on Linux to query Outlook on your Windows PC:

1. **On your Windows PC**, start the MCP server over your local network:
   ```cmd
   run.bat --sse --host 0.0.0.0 --port 8000
   ```
2. **On your Linux machine**, point your MCP client (Claude Desktop / AGY) to:
   `http://<windows-pc-ip>:8000/sse`

---

## 🛠️ Available MCP Tools

| Tool Name | Parameters | Description |
| :--- | :--- | :--- |
| `outlook_check_status` | *none* | Check connection to Outlook |
| `outlook_list_messages` | `folder`, `top`, `unread_only` | List recent emails from Inbox, Sent, Drafts |
| `outlook_search_messages` | `query`, `top` | Search messages across mailbox |
| `outlook_get_message` | `message_id` | Retrieve full email body and sender details |
| `outlook_create_draft` | `subject`, `body`, `to_recipients`, `reply_to_message_id` | Create a draft email or reply draft in Outlook |
| `outlook_send_mail` | `subject`, `body`, `to_recipients` | Send an email directly |
| `outlook_list_calendar_events`| `days_ahead` | List upcoming meetings and appointments |
| `outlook_create_calendar_event`| `subject`, `start_datetime`, `end_datetime`, `attendees`, `location` | Schedule a meeting in Outlook Calendar |
| `outlook_list_folders` | *none* | List mail folders with unread counts |

---

## 💬 Example AI Prompts

- *"Check my Outlook inbox and summarize the unread emails from today."*
- *"Search my emails for the latest project update from Alex."*
- *"What meetings do I have on my calendar for the rest of the week?"*
- *"Draft a reply to the email with ID `<id>` letting them know I'll send the report on Friday."*
