# Teams Chat Exporter Setup & Run Guide

Export the **entire history** of specific **Microsoft Teams *group chats*** (picked by `chatId`) to:
- **JSONL** (one message per line, full Graph payload)
- **CSV** (flattened, human-readable)
- **Inline media** (`hostedContents`) saved to disk

> ✅ Targets **only** the chat IDs you paste at runtime.  

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
- You must be able to obtain the **chatId(s)** (format like `19:xxxxxxxxxxxxxxxx@thread.v2`)

---

## Repo Setup

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

Create a .env file with the following contents filled out according to the tenant:

TENANT_ID=00000000-1111-2222-3333-444444444444
CLIENT_ID=aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee


CLIENT_SECRET=your-super-secret
