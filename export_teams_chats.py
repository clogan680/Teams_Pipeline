import argparse
import csv
import json
import os
import re
import time
from datetime import datetime, timezone
from typing import Dict, Iterable, List, Optional

import requests
from dateutil import parser as dtparse
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_SCOPE = ["https://graph.microsoft.com/.default"]


def backoff_sleep(retry_after: Optional[str], attempt: int):
    if retry_after and retry_after.isdigit():
        time.sleep(int(retry_after))
    else:
        time.sleep(min(60, 2**attempt))


def clean_html_to_text(html: Optional[str]) -> str:
    if not html:
        return ""
    return re.sub(r"<[^>]+>", " ", html).replace("&nbsp;", " ").strip()


def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def write_jsonl(path: str, items: Iterable[Dict]):
    with open(path, "w", encoding="utf-8") as f:
        for it in items:
            f.write(json.dumps(it, ensure_ascii=False) + "\n")


def get_token_from_env() -> str:
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    client_cert_path = os.getenv("CLIENT_CERT_PATH")
    cert_thumbprint = os.getenv("CERT_THUMBPRINT")

    if not tenant_id or not client_id:
        raise SystemExit("Missing TENANT_ID or CLIENT_ID in .env")

    if client_secret:
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
    elif client_cert_path and cert_thumbprint:
        with open(client_cert_path, "rb") as f:
            pem = f.read()
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential={"private_key": pem, "thumbprint": cert_thumbprint},
        )
    else:
        raise SystemExit(
            "Provide CLIENT_SECRET or CLIENT_CERT_PATH + CERT_THUMBPRINT in .env"
        )

    result = app.acquire_token_for_client(scopes=TOKEN_SCOPE)
    if "access_token" not in result:
        raise SystemExit(f"Token error: {result}")
    return result["access_token"]


def graph_get_all(
    url: str, token: str, params: Optional[Dict] = None
) -> Iterable[Dict]:
    sess = requests.Session()
    next_url = url
    q = params or {}
    attempt = 0
    while True:
        r = sess.get(next_url, headers={"Authorization": f"Bearer {token}"}, params=q)
        if r.status_code in (429, 503, 504):
            attempt += 1
            backoff_sleep(r.headers.get("Retry-After"), attempt)
            continue
        if not r.ok:
            raise RuntimeError(f"Graph error {r.status_code}: {r.text}")
        data = r.json()
        for v in data.get("value", []):
            yield v
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        next_url, q = next_link, None  # nextLink already encodes params


def download_hosted_contents(
    chat_id: str, message_id: str, token: str, out_dir: str
) -> List[str]:
    saved = []
    list_url = f"{GRAPH}/chats/{chat_id}/messages/{message_id}/hostedContents"
    for hc in graph_get_all(list_url, token):
        hc_id = hc.get("id")
        if not hc_id:
            continue
        bytes_url = f"{GRAPH}/chats/{chat_id}/messages/{message_id}/hostedContents/{hc_id}/$value"
        attempt = 0
        while True:
            resp = requests.get(bytes_url, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code in (429, 503, 504):
                attempt += 1
                backoff_sleep(resp.headers.get("Retry-After"), attempt)
                continue
            if not resp.ok:
                break
            content = resp.content
            ct = resp.headers.get("Content-Type", "")
            ext = (
                ".png"
                if "png" in ct
                else (
                    ".jpg"
                    if ("jpeg" in ct or "jpg" in ct)
                    else (
                        ".gif"
                        if "gif" in ct
                        else (
                            ".svg"
                            if "svg" in ct
                            else ".webp" if "webp" in ct else ".bin"
                        )
                    )
                )
            )
            fname = f"{message_id}_{hc_id}{ext}"
            path = os.path.join(out_dir, fname)
            with open(path, "wb") as f:
                f.write(content)
            saved.append(path)
            break
    return saved


CSV_FIELDS = [
    "chatId",
    "messageId",
    "createdDateTime",
    "fromUserId",
    "fromDisplayName",
    "messageType",
    "importance",
    "subject",
    "hasAttachments",
    "bodyContentType",
    "bodyText",
    "replyToId",
    "lastModifiedDateTime",
    "deletedDateTime",
    "reactions",
    "hostedContentFiles",
]


def flatten_for_csv(m: Dict) -> Dict:
    from_user = (m.get("from") or {}).get("user") or {}
    reactions = "; ".join(
        f"{r.get('reactionType')}:{((r.get('user') or {}).get('user') or {}).get('displayName','')}"
        for r in m.get("reactions", [])
    )
    body = m.get("body") or {}
    return {
        "chatId": m.get("_chatId", ""),
        "messageId": m.get("id", ""),
        "createdDateTime": m.get("createdDateTime", ""),
        "fromUserId": from_user.get("id", ""),
        "fromDisplayName": from_user.get("displayName", ""),
        "messageType": m.get("messageType", ""),
        "importance": m.get("importance", ""),
        "subject": m.get("subject", ""),
        "hasAttachments": bool(m.get("attachments")),
        "bodyContentType": body.get("contentType", ""),
        "bodyText": clean_html_to_text(body.get("content", "")),
        "replyToId": m.get("replyToId", ""),
        "lastModifiedDateTime": m.get("lastModifiedDateTime", ""),
        "deletedDateTime": m.get("deletedDateTime", ""),
        "reactions": reactions,
        "hostedContentFiles": "; ".join(m.get("_hostedFiles", [])),
    }


def export_csv(path: str, rows: Iterable[Dict]):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def export_chat(
    chat_id: str,
    token: str,
    out_dir: str,
    after: Optional[datetime],
    before: Optional[datetime],
):
    print(f"Exporting chat {chat_id} ...")
    os.makedirs(out_dir, exist_ok=True)
    raw_msgs: List[Dict] = []
    safe = re.sub(r"[^A-Za-z0-9]+", "_", chat_id)
    jsonl_path = os.path.join(out_dir, f"chat_{safe}.jsonl")
    csv_path = os.path.join(out_dir, f"chat_{safe}.csv")
    media_dir = os.path.join(out_dir, "media", safe)
    os.makedirs(media_dir, exist_ok=True)

    url = f"{GRAPH}/chats/{chat_id}/messages"
    for m in graph_get_all(url, token):
        m["_chatId"] = chat_id

        if after or before:
            ts = dtparse.parse(m.get("createdDateTime", "")).astimezone(timezone.utc)
            if after and ts < after:
                continue
            if before and ts > before:
                continue

        saved_files = download_hosted_contents(chat_id, m["id"], token, media_dir)
        if saved_files:
            m["_hostedFiles"] = saved_files

        raw_msgs.append(m)

    write_jsonl(jsonl_path, raw_msgs)
    export_csv(csv_path, (flatten_for_csv(m) for m in raw_msgs))

    print(f"  -> JSONL: {jsonl_path}")
    print(f"  -> CSV  : {csv_path}")
    if any(m.get("_hostedFiles") for m in raw_msgs):
        print(f"  -> Media: {media_dir}")


def parse_dt(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    return dtparse.parse(s).astimezone(timezone.utc)


def main():
    load_dotenv()  # loads .env from current directory

    ap = argparse.ArgumentParser(
        description="Export specific Microsoft Teams group chats by chatId (config from .env)"
    )
    ap.add_argument("--out-dir", default="TeamsChatExports")
    ap.add_argument("--after", help="Export messages on/after this time (ISO8601)")
    ap.add_argument("--before", help="Export messages on/before this time (ISO8601)")
    args = ap.parse_args()

    token = get_token_from_env()
    after = parse_dt(args.after)
    before = parse_dt(args.before)

    # Prompt user for one or more chat IDs
    chat_input = input(
        "Enter one or more chat IDs (comma-separated), e.g. 19:...@thread.v2: "
    ).strip()
    chat_ids = [c.strip() for c in chat_input.split(",") if c.strip()]
    if not chat_ids:
        raise SystemExit("No chat IDs provided.")

    os.makedirs(args.out_dir, exist_ok=True)
    for cid in chat_ids:
        export_chat(cid, token, args.out_dir, after, before)


if __name__ == "__main__":
    main()
