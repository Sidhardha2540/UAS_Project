"""
Outlook Mail Reader - Fetches Inbox emails via Microsoft Graph API.
Can list messages or process PDF attachments with BEO (Banquet Event Order) validation and folder save.
Uses Device Code flow for one-time auth.
"""

import argparse
import os
import sys
from pathlib import Path

import httpx
from dotenv import load_dotenv
from msal import PublicClientApplication
from msal.token_cache import SerializableTokenCache

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
SAVE_TO_ONEDRIVE = os.getenv("SAVE_TO_ONEDRIVE", "").strip().lower() in ("1", "true", "yes")
SCOPES = ["User.Read", "Mail.Read", "Files.ReadWrite.All"]
# Use "common" when saving to OneDrive so work/school accounts can sign in
AUTHORITY = "https://login.microsoftonline.com/common" if SAVE_TO_ONEDRIVE else "https://login.microsoftonline.com/consumers"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_FILE = Path(__file__).parent / "token_cache.bin"


def get_token() -> str:
    """Get access token via MSAL. Uses cache if available, else Device Code flow."""
    if not CLIENT_ID:
        print("Error: CLIENT_ID not set. Create a .env file with CLIENT_ID=your-app-client-id")
        sys.exit(1)

    cache = SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())

    app = PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            _save_cache(app)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow: {flow.get('error_description', flow)}")

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(
            result.get("error_description", "Failed to acquire token")
        )

    _save_cache(app)
    return result["access_token"]


def _save_cache(app: PublicClientApplication) -> None:
    cache = app.token_cache
    if cache and hasattr(cache, "serialize"):
        TOKEN_CACHE_FILE.write_text(cache.serialize())


def fetch_inbox_messages(access_token: str, include_id: bool = False) -> list[dict]:
    """Fetch all Inbox messages from Microsoft Graph, handling pagination."""
    headers = {"Authorization": f"Bearer {access_token}"}
    select = "subject,from,receivedDateTime,bodyPreview"
    if include_id:
        select = "id," + select
    url = (
        f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        f"?$top=100&$select={select}"
        "&$orderby=receivedDateTime desc"
    )
    all_messages = []

    with httpx.Client() as client:
        while url:
            resp = client.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
            all_messages.extend(data.get("value", []))
            url = data.get("@odata.nextLink")

    return all_messages


def list_attachments(access_token: str, message_id: str) -> list[dict]:
    """List attachments for a message. Returns list of attachment dicts with id, name, contentType."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments?$select=id,name,contentType"
    with httpx.Client() as client:
        resp = client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("value", [])


def download_attachment(access_token: str, message_id: str, attachment_id: str) -> bytes:
    """Download raw bytes of an attachment (e.g. PDF) using $value."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments/{attachment_id}/$value"
    with httpx.Client(timeout=60.0) as client:
        resp = client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.content


def is_pdf_attachment(att: dict) -> bool:
    """True if attachment looks like a PDF (by contentType or filename)."""
    name = (att.get("name") or "").lower()
    ct = (att.get("contentType") or "").lower()
    return "pdf" in ct or name.endswith(".pdf")


def print_messages(messages: list[dict]) -> None:
    """Print each message to the terminal."""
    if not messages:
        print("No messages in Inbox.")
        return

    for i, msg in enumerate(messages, 1):
        subject = msg.get("subject", "(No subject)")
        from_info = msg.get("from", {}).get("emailAddress", {})
        from_name = from_info.get("name", "Unknown")
        from_addr = from_info.get("address", "")
        received = msg.get("receivedDateTime", "")
        preview = msg.get("bodyPreview", "") or "(No preview)"

        print(f"\n{'='*60}")
        print(f"Message {i}")
        print(f"{'='*60}")
        print(f"Subject: {subject}")
        print(f"From: {from_name} <{from_addr}>")
        print(f"Received: {received}")
        print(f"Preview: {preview}")

    print(f"\n{'='*60}")
    print(f"Total: {len(messages)} message(s)")
    print(f"{'='*60}")


def run_beo_pipeline(access_token: str) -> None:
    """Fetch inbox, for each message with PDF attachments run BEO processor and save valid PDFs."""
    from beo_processor import process_pdf

    messages = fetch_inbox_messages(access_token, include_id=True)
    saved = 0
    skipped = 0
    for msg in messages:
        msg_id = msg.get("id")
        if not msg_id:
            continue
        attachments = list_attachments(access_token, msg_id)
        pdfs = [a for a in attachments if is_pdf_attachment(a)]
        for att in pdfs:
            try:
                content = download_attachment(access_token, msg_id, att["id"])
                name = att.get("name") or "document.pdf"
                path = process_pdf(content, name, access_token=access_token)
                if path:
                    if isinstance(path, str):
                        print(f"Saved to OneDrive: {path}")
                    else:
                        print(f"Saved: {path}")
                    saved += 1
                else:
                    skipped += 1
            except Exception as e:
                print(f"Error processing {att.get('name', '?')}: {e}", file=sys.stderr)
    print(f"BEO pipeline done. Saved: {saved}, skipped/invalid: {skipped}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Outlook mail reader and BEO PDF processor")
    parser.add_argument(
        "--list",
        action="store_true",
        help="Only list inbox messages (no BEO processing)",
    )
    args = parser.parse_args()

    token = get_token()
    if args.list:
        messages = fetch_inbox_messages(token)
        print_messages(messages)
    else:
        run_beo_pipeline(token)


if __name__ == "__main__":
    main()
