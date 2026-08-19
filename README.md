# Microsoft Outlook MCP Server for Claude & Antigravity (AGY)

A private, high-performance **Model Context Protocol (MCP)** server providing direct, secure access to your **Work Microsoft 365 Outlook** emails, calendar, and contacts from:
- **Claude Code CLI**
- **Claude Desktop**
- **Antigravity CLI (AGY)**

---

## 🔒 Privacy & Security Architecture

- **Human-in-the-Loop by Design**: All outbound message drafting defaults to your Outlook `Drafts` folder for manual review. Direct unreviewed sends (`outlook_send_mail`) require an explicit `confirm_send=True` flag to prevent accidental dispatches.
- **Indirect Prompt Injection Defense**: Incoming email bodies are encapsulated in `<untrusted_email_body_content>` XML tags to ensure AI models parse external content strictly as passive data.
- **Zero Cloud Tokens / No Passwords**: Connects 100% locally to your running Windows Outlook Desktop application via Windows MAPI (`win32com`).
- **Strict `.gitignore`**: All local configs, caches, and environments are ignored.

---

## 🚀 Setup on Windows (`C:\projects\mcp-outlook`)

### 1. Clone into `C:\projects`
Open Command Prompt or PowerShell on your Windows PC:
```cmd
cd C:\projects
gh repo clone Jpatching/mcp-outlook
cd mcp-outlook
setup.bat
```

### 2. Test Connection (with Outlook open)
```cmd
run.bat --status
run.bat --test
```

---

## ⚙️ Connecting to Claude & Antigravity on Windows

### 1. Claude Desktop (Windows)
Open `%APPDATA%\Claude\claude_desktop_config.json` and add:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\projects\\mcp-outlook\\run.bat"
    }
  }
}
```

### 2. Claude Code CLI (Windows)
```cmd
claude mcp add outlook C:\projects\mcp-outlook\run.bat
```

### 3. Antigravity CLI (Windows)
In `%USERPROFILE%\.gemini\config\mcp_config.json`:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\projects\\mcp-outlook\\run.bat",
      "args": []
    }
  }
}
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
