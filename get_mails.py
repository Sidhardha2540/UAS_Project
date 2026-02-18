"""
Outlook Mail Reader - Fetches Inbox emails via Microsoft Graph API.
Can list messages or process PDF attachments with BEO (Banquet Event Order) validation and folder save.
Uses Device Code flow for one-time auth.
"""

import argparse
import csv
import json
import os
import sys
from datetime import datetime
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
PROCESSED_FILE = Path(__file__).parent / "beo_processed.json"
# When True, skip attachments already in beo_processed.json (resume). When False, run from start every time.
BEO_RESUME = os.getenv("BEO_RESUME", "").strip().lower() in ("1", "true", "yes")


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

    with httpx.Client(timeout=60.0) as client:
        while url:
            resp = client.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
            all_messages.extend(data.get("value", []))
            url = data.get("@odata.nextLink")

    return all_messages


def list_attachments(access_token: str, message_id: str) -> list[dict]:
    """List all attachments for a message (handles pagination). Returns list of attachment dicts with id, name, contentType."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments?$select=id,name,contentType&$top=100"
    all_attachments = []
    with httpx.Client(timeout=30.0) as client:
        while url:
            resp = client.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
            all_attachments.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
    return all_attachments


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


def _load_processed_set() -> set[tuple[str, str]]:
    """Load set of (message_id, attachment_id) already processed."""
    if not PROCESSED_FILE.exists():
        return set()
    try:
        data = json.loads(PROCESSED_FILE.read_text(encoding="utf-8"))
        items = data.get("processed") or []
        return {(x["message_id"], x["attachment_id"]) for x in items}
    except Exception:
        return set()


def _save_processed(processed_set: set[tuple[str, str]]) -> None:
    """Persist processed set to beo_processed.json."""
    items = [{"message_id": mid, "attachment_id": aid} for mid, aid in processed_set]
    PROCESSED_FILE.write_text(json.dumps({"processed": items}, indent=2), encoding="utf-8")


def _write_review_report(flagged: list[dict], report_dir: Path):
    """Write a CSV report of flagged items for manual review. Returns path to the report file."""
    if not flagged:
        return
    report_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = report_dir / f"beo_review_report_{timestamp}.csv"
    with open(report_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["attachment_name", "beo_number", "beo_date", "folder_name_used", "reason"])
        writer.writeheader()
        writer.writerows(flagged)
    return report_path


def run_beo_pipeline(access_token: str) -> None:
    """Fetch inbox and process PDFs. By default runs from start; set BEO_RESUME=true to skip already-processed."""
    from beo_processor import process_pdf

    processed_set = _load_processed_set() if BEO_RESUME else set()
    if BEO_RESUME:
        print("Resume mode: skipping attachments already in beo_processed.json")
    else:
        print("Run from start: processing all PDF attachments in Inbox")
    messages = fetch_inbox_messages(access_token, include_id=True)
    print(f"Fetched {len(messages)} message(s) from Inbox")
    saved = 0
    skipped = 0
    already_processed = 0
    flagged_for_review: list[dict] = []
    total_pdfs = 0
    for msg in messages:
        msg_id = msg.get("id")
        if not msg_id:
            continue
        attachments = list_attachments(access_token, msg_id)
        pdfs = [a for a in attachments if is_pdf_attachment(a)]
        total_pdfs += len(pdfs)
        for att in pdfs:
            att_id = att.get("id")
            if not att_id:
                continue
            if BEO_RESUME and (msg_id, att_id) in processed_set:
                already_processed += 1
                continue
            try:
                content = download_attachment(access_token, msg_id, att_id)
                name = att.get("name") or "document.pdf"
                path, report_entry = process_pdf(content, name, access_token=access_token)
                if path:
                    if isinstance(path, str):
                        print(f"Saved to OneDrive: {path}")
                    else:
                        print(f"Saved: {path}")
                    if report_entry:
                        flagged_for_review.append(report_entry)
                        print(f"  [Flagged for review: {report_entry.get('reason', '?')}]")
                    saved += 1
                    if BEO_RESUME:
                        processed_set.add((msg_id, att_id))
                        _save_processed(processed_set)
                else:
                    skipped += 1
            except Exception as e:
                print(f"Error processing {att.get('name', '?')}: {e}", file=sys.stderr)
    print(f"BEO pipeline done. PDFs found: {total_pdfs}, saved: {saved}, skipped/invalid: {skipped}" + (f", already processed: {already_processed}" if BEO_RESUME and already_processed else ""))
    if total_pdfs == 0:
        print("No PDF attachments found in Inbox. Add messages with PDF attachments and run again.")
    if flagged_for_review:
        report_dir = Path(__file__).parent
        report_path = _write_review_report(flagged_for_review, report_dir)
        if report_path is not None:
            print(f"Review report ({len(flagged_for_review)} item(s)): {report_path}")


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
