"""
Outlook Mail Reader - Fetches Inbox emails for jmovva25@outlook.com via Microsoft Graph API.
Prints all received mails to the terminal. Uses Device Code flow for one-time auth.
"""

import os
import sys
from pathlib import Path

import httpx
from dotenv import load_dotenv
from msal import PublicClientApplication
from msal.token_cache import SerializableTokenCache

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
SCOPES = ["User.Read", "Mail.Read"]
AUTHORITY = "https://login.microsoftonline.com/consumers"
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


def fetch_inbox_messages(access_token: str) -> list[dict]:
    """Fetch all Inbox messages from Microsoft Graph, handling pagination."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        "?$top=100&$select=subject,from,receivedDateTime,bodyPreview"
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


def main() -> None:
    token = get_token()
    messages = fetch_inbox_messages(token)
    print_messages(messages)


if __name__ == "__main__":
    main()
