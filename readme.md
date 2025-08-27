# Teams Chat Exporter — One-File Setup & Run Guide

Export the **entire history** of specific **Microsoft Teams *group chats*** (picked by `chatId`) to:
- **JSONL** (one message per line, full Graph payload)
- **CSV** (flattened, human-readable)
- **Inline media** (`hostedContents`) saved to disk

> ✅ Targets **only** the chat IDs you paste at runtime.  
> ❌ Does **not** pull all chats for a user.

---

## What’s in this repo

- `export_teams_chats.py` — Python 3.12 script that:
  - Authenticates with Microsoft Graph using **app-only** credentials from a `.env` file
  - Prompts you for one or more **chat IDs**
  - Paginates through **all messages** in each chat
  - Saves **JSONL**, **CSV**, and **inline media** (hosted contents)

---

## Prerequisites

- **Python 3.12+**
- Ability to create/modify an **Entra ID (Azure AD) app registration** and **grant admin consent**
- You must be able to obtain the **chatId(s)** (format like `19:xxxxxxxxxxxxxxxx@thread.v2`)

---

## 1) Entra / Microsoft Graph Setup (one-time)

1) **Register an app**
- Entra ID → **App registrations** → **New registration**
- Note the **Application (client) ID** and **Directory (tenant) ID**

2) **Authentication**
- Either:
  - **Client secret**: Certificates & secrets → **New client secret** → copy the value
  - **Certificate** (preferred): Upload your public cert; keep the private key locally (PEM)

3) **API permissions (Application)**
- App → **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
  - **Required**: `Chat.Read.All`
  - (Optional, if you’ll export *channel* messages later): `ChannelMessage.Read.All`
- Click **Grant admin consent** for the tenant

> Some tenants gate Teams chat content behind “Protected APIs.” If you get a protected-API/forbidden error, have a tenant admin approve access for Teams chat content for this app.

4) **Least-privilege alternative (optional)**
- If you want to restrict to specific chats only, use **Resource-Specific Consent (RSC)**:
  - Add `ChatMessage.Read.Chat`
  - Package/install your app **into each target chat** in Teams
  - Then run this tool (same script) — access will be scoped just to those chats

---

## 2) Repo Setup

```bash
# From project root
python3.12 -m venv venv
# macOS/Linux
source venv/bin/activate
# Windows (PowerShell)
.\venv\Scripts\Activate.ps1

pip install --upgrade pip
# Install dependencies
pip install msal requests python-dateutil pandas python-dotenv
