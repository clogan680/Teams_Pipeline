import argparse
import csv
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from typing import Dict, Iterable, List, Optional

import requests
from dateutil import parser as dtparse
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_SCOPE = ["https://graph.microsoft.com/.default"]

# --- Utilities ----------------------------------------------------------------


def backoff_sleep(retry_after: Optional[str], attempt: int):
    if retry_after and retry_after.isdigit():
        time.sleep(int(retry_after))
    else:
        # simple exponential backoff with cap
        time.sleep(min(60, 2**attempt))


def clean_html_to_text(html: Optional[str]) -> str:
    if not html:
        return ""
    # A very light HTML -> text cleaner (good enough for CSV)
    return re.sub(r"<[^>]+>", " ", html).replace("&nbsp;", " ").strip()


def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def write_jsonl(path: str, items: Iterable[Dict]):
    with open(path, "w", encoding="utf-8") as f:
        for it in items:
            f.write(json.dumps(it, ensure_ascii=False) + "\n")


# --- Auth ---------------------------------------------------------------------


def get_token(
    tenant_id: str,
    client_id: str,
    client_secret: Optional[str],
    client_cert: Optional[str],
    cert_thumbprint: Optional[str],
) -> str:
    if client_secret:
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
    elif client_cert:
        with open(client_cert, "rb") as f:
            pem = f.read()
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential={"private_key": pem, "thumbprint": cert_thumbprint},
        )
    else:
        raise SystemExit(
            "Provide either --client-secret or --client-cert/--cert-thumbprint"
        )

    result = app.acquire_token_for_client(scopes=TOKEN_SCOPE)
    if "access_token" not in result:
        raise SystemExit(f"Token error: {result}")
    return result["access_token"]


# --- Graph calling with pagination & retries ----------------------------------


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
        values = data.get("value", [])
        for v in values:
            yield v
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        # after first page, subsequent calls must use nextLink as-is
        next_url, q = next_link, None


# --- Hosted contents (inline images, cards) -----------------------------------


def download_hosted_contents(
    chat_id: str, message_id: str, token: str, out_dir: str
) -> List[str]:
    """
    Returns list of saved file paths for hosted contents in a message (if any).
    """
    saved = []
    list_url = f"{GRAPH}/chats/{chat_id}/messages/{message_id}/hostedContents"
    for hc in graph_get_all(list_url, token):
        hc_id = hc.get("id")
        if not hc_id:
            continue
        # bytes endpoint
        bytes_url = f"{GRAPH}/chats/{chat_id}/messages/{message_id}/hostedContents/{hc_id}/$value"
        attempt = 0
        while True:
            resp = requests.get(bytes_url, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code in (429, 503, 504):
                attempt += 1
                backoff_sleep(resp.headers.get("Retry-After"), attempt)
                continue
            if not resp.ok:
                # If we can't fetch, just skip this blob
                break
            content = resp.content
            # derive an extension if possible
            ct = resp.headers.get("Content-Type", "")
            ext = ""
            if "png" in ct:
                ext = ".png"
            elif "jpeg" in ct or "jpg" in ct:
                ext = ".jpg"
            elif "gif" in ct:
                ext = ".gif"
            elif "svg" in ct:
                ext = ".svg"
            elif "webp" in ct:
                ext = ".webp"
            else:
                ext = ".bin"
            fname = f"{message_id}_{hc_id}{ext}"
            fpath = os.path.join(out_dir, fname)
            with open(fpath, "wb") as f:
                f.write(content)
            saved.append(fpath)
            break
    return saved


# --- Exporters ----------------------------------------------------------------

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
    get = lambda *ks: (ks[0] in m and m[ks[0]]) or ""
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


# --- Main export --------------------------------------------------------------


def export_chat(
    chat_id: str,
    token: str,
    out_dir: str,
    after: Optional[datetime],
    before: Optional[datetime],
):
    print(f"Exporting chat {chat_id} ...")
    ensure_dir(out_dir)
    raw_msgs = []
    jsonl_path = os.path.join(
        out_dir, f"chat_{re.sub(r'[^A-Za-z0-9]+','_', chat_id)}.jsonl"
    )
    csv_path = os.path.join(
        out_dir, f"chat_{re.sub(r'[^A-Za-z0-9]+','_', chat_id)}.csv"
    )
    media_dir = os.path.join(out_dir, "media", re.sub(r"[^A-Za-z0-9]+", "_", chat_id))
    ensure_dir(media_dir)

    url = f"{GRAPH}/chats/{chat_id}/messages"
    for m in graph_get_all(url, token):
        # Keep for JSONL
        m["_chatId"] = chat_id

        # Date filters (client-side)
        if after or before:
            ts = dtparse.parse(m.get("createdDateTime", "")).astimezone(timezone.utc)
            if after and ts < after:
                continue
            if before and ts > before:
                continue

        # Download inline hosted contents (if any)
        saved_files = download_hosted_contents(chat_id, m["id"], token, media_dir)
        if saved_files:
            m["_hostedFiles"] = saved_files

        raw_msgs.append(m)

    # Write JSONL
    write_jsonl(jsonl_path, raw_msgs)

    # Write CSV
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
    ap = argparse.ArgumentParser(
        description="Export specific Microsoft Teams group chats by chatId"
    )
    ap.add_argument("--tenant-id", required=True)
    ap.add_argument("--client-id", required=True)
    auth = ap.add_mutually_exclusive_group(required=True)
    auth.add_argument("--client-secret")
    auth.add_argument("--client-cert", help="Path to PEM/PKCS1 private key file")
    ap.add_argument(
        "--cert-thumbprint", help="Thumbprint string (required if using --client-cert)"
    )
    ap.add_argument(
        "--chat-id",
        action="append",
        required=True,
        help="Repeat for each chatId (e.g., 19:...@thread.v2)",
    )
    ap.add_argument("--out-dir", default="TeamsChatExports")
    ap.add_argument(
        "--after", help="UTC or local ISO8601; export messages on/after this time"
    )
    ap.add_argument(
        "--before", help="UTC or local ISO8601; export messages on/before this time"
    )
    args = ap.parse_args()

    after = parse_dt(args.after)
    before = parse_dt(args.before)

    token = get_token(
        args.tenant_id,
        args.client_id,
        args.client_secret,
        args.client_cert,
        args.cert_thumbprint,
    )
    ensure_dir(args.out_dir)

    for cid in args.chat_id:
        export_chat(cid, token, args.out_dir, after, before)


if __name__ == "__main__":
    main()
