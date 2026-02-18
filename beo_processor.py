"""
BEO (Banquet Event Order) processor: PDF text extraction, GPT-4o mini agent validation,
and folder structure creation for valid BEO documents. Saves to local path or to
OneDrive when SAVE_TO_ONEDRIVE is set and access_token is provided.
"""

import os
import re
from pathlib import Path
from typing import Optional, Tuple, Union
from urllib.parse import quote

import httpx
import pymupdf
from dateutil import parser as date_parser
from dotenv import load_dotenv
from pydantic import BaseModel, Field

from agents import Agent, Runner

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
BEO_BASE_PATH = os.getenv("BEO_BASE_PATH", "").strip() or str(Path(__file__).parent / "beo_output")
_SAVE_TO_ONEDRIVE = os.getenv("SAVE_TO_ONEDRIVE", "").strip().lower() in ("1", "true", "yes")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
# Timeouts for Graph API: longer for uploads (large PDFs)
GRAPH_TIMEOUT = 60.0
GRAPH_UPLOAD_TIMEOUT = 120.0

BEO_AGENT_INSTRUCTIONS = """You are a document analyst. The user will provide the full text extracted from a single PDF document.

Your task:
1. **Condition 1**: Determine whether this document contains BOTH:
   (a) A Hospitality form that is SIGNED (signed by a person, not just a blank form).
   (b) A BEO (Banquet Event Order) attached or included in the document.

2. **Condition 2**: If and only if Condition 1 is satisfied, extract from the document:
   - **BEO number**: A five-digit number (e.g. 12345). Return only the digits as a string.
   - **BEO date**: The date of the BEO in any format you find (e.g. MM/DD/YYYY or YYYY-MM-DD).
   - **Organization name**: The organization name from the "Client/Organization" field (e.g. "Guardian Scholars Program", "Department of Communication Disorder"). Use the organization/company/department name only. If the field clearly contains a person's name instead of an organization, leave this null.
   - **Client name**: The contact or client person name (e.g. from Booking Contact, or the main contact on the form). Use this when Organization name is unclear or when the document only shows a person's name.

Return valid=true when both (a) and (b) are present and you can extract BEO number and BEO date. You must have at least one of Organization name or Client name for valid=true. Prefer Organization name for folder naming; Client name is used as fallback when Organization is missing or ambiguous."""


class BEOResult(BaseModel):
    """Structured output from the BEO document analyst agent."""

    valid: bool = Field(description="True if document contains signed Hospitality form and BEO, and extraction succeeded.")
    beo_number: Optional[str] = Field(default=None, description="Five-digit BEO number as string, or null.")
    beo_date: Optional[str] = Field(default=None, description="Date of the BEO as string, or null.")
    organization_name: Optional[str] = Field(default=None, description="Organization name from Client/Organization field, or null.")
    client_name: Optional[str] = Field(default=None, description="Client/contact person name for fallback when organization unclear, or null.")


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """Extract full text from all pages of a PDF using PyMuPDF."""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        parts = []
        for page in doc:
            parts.append(page.get_text())
        return "\n\n".join(parts).strip() or ""
    finally:
        doc.close()


def _create_beo_agent() -> Agent:
    """Create the GPT-4o mini document analyst agent with structured output."""
    return Agent(
        name="BEO Document Analyst",
        instructions=BEO_AGENT_INSTRUCTIONS,
        model="gpt-4o-mini",
        output_type=BEOResult,
    )


def analyze_pdf_with_agent(pdf_text: str) -> BEOResult:
    """Send full PDF text to the agent and return structured BEO result."""
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not set. Add it to your .env file.")
    agent = _create_beo_agent()
    user_message = (
        "Below is the full text extracted from a PDF document.\n\n"
        "---\n\n"
        f"{pdf_text}\n\n"
        "---\n\n"
        "Based on this document only, determine if it satisfies Condition 1 (signed Hospitality form + BEO present) "
        "and if so, extract the BEO number (five digits), BEO date, organization name (from Client/Organization field), and client/contact name (for fallback). Return the structured response."
    )
    result = Runner.run_sync(agent, user_message)
    output = result.final_output
    if isinstance(output, BEOResult):
        return output
    if isinstance(output, dict):
        return BEOResult(**output)
    raise TypeError(f"Unexpected agent output type: {type(output)}")


def _parse_beo_date(beo_date_str: Optional[str]) -> Optional[tuple[int, int, int]]:
    """Parse BEO date string to (year, month, day). Returns None if unparseable."""
    if not beo_date_str or not beo_date_str.strip():
        return None
    try:
        dt = date_parser.parse(beo_date_str.strip())
        return (dt.year, dt.month, dt.day)
    except Exception:
        return None


def _sanitize_filename(name: str) -> str:
    """Replace characters that are invalid in folder names."""
    if not name:
        return "Unknown"
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip() or "Unknown"


def _normalize_beo_number(beo_number: str) -> str:
    """Strip leading non-digits (e.g. L, E) so L43105 and 43105 become the same folder."""
    s = (beo_number or "").strip()
    digits = re.sub(r"^\D*", "", s)
    return digits if digits else s


def build_beo_path_segments(
    beo_number: str, beo_date_str: Optional[str], name_for_folder: str
) -> Optional[tuple[str, str, str, str]]:
    """
    Build path segments (year, month, day, folder_name) for local or OneDrive.
    name_for_folder can be organization name or client name (fallback). Returns None if date unparseable or missing.
    """
    if not beo_number or not (name_for_folder or "").strip():
        return None
    date_parts = _parse_beo_date(beo_date_str)
    if not date_parts:
        return None
    year, month, day = date_parts
    beo_num_normalized = _normalize_beo_number(beo_number)
    if beo_num_normalized.isdigit():
        beo_num_normalized = beo_num_normalized.zfill(5)
    name_safe = _sanitize_filename(name_for_folder)
    folder_name = f"{beo_num_normalized} - {name_safe}"
    return (str(year), str(month), str(day), folder_name)


def build_beo_folder_path(beo_number: str, beo_date_str: Optional[str], name_for_folder: str) -> Optional[Path]:
    """
    Build local path: {base}/{year}/{month}/{day}/{beo_number} - {name_for_folder}.
    Returns None if date cannot be parsed or required fields are missing.
    """
    segs = build_beo_path_segments(beo_number, beo_date_str, name_for_folder)
    if not segs:
        return None
    year, month, day, folder_name = segs
    return Path(BEO_BASE_PATH) / year / month / day / folder_name


def _get_onedrive_id(access_token: str) -> str:
    """Get the signed-in user's OneDrive drive ID."""
    url = f"{GRAPH_BASE}/me/drive"
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        resp = client.get(url, headers={"Authorization": f"Bearer {access_token}"})
        resp.raise_for_status()
        return resp.json()["id"]


def _ensure_drive_folders(access_token: str, drive_id: str, path_segments: Tuple[str, str, str, str]) -> None:
    """Ensure the folder path exists in the drive (create each level if missing)."""
    year, month, day, folder_name = path_segments
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    segments = (year, month, day, folder_name)
    current_path = ""
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        for segment in segments:
            next_path = f"{current_path}/{segment}".strip("/") if current_path else segment
            path_encoded = quote(next_path, safe="/")
            get_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path_encoded}"
            r = client.get(get_url, headers={"Authorization": f"Bearer {access_token}"})
            if r.status_code == 200:
                current_path = next_path
                continue
            parent_ref = "root" if not current_path else quote(current_path, safe="/")
            post_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{parent_ref}:/children"
            r = client.post(
                post_url,
                headers=headers,
                json={"name": segment, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
            )
            if r.status_code == 409:
                current_path = next_path
                continue
            r.raise_for_status()
            current_path = next_path


def _get_existing_beo_folder_for_day(
    access_token: str, drive_id: str, year: str, month: str, day: str, normalized_beo_number: str
) -> Optional[str]:
    """If a folder for this BEO (e.g. 43105 or L43105) already exists under year/month/day, return its name; else None."""
    day_path = f"{year}/{month}/{day}"
    path_encoded = quote(day_path, safe="/")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path_encoded}:/children"
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        resp = client.get(url, headers={"Authorization": f"Bearer {access_token}"})
    if resp.status_code != 200:
        return None
    data = resp.json()
    for item in data.get("value", []):
        if item.get("folder"):
            name = (item.get("name") or "").strip()
            first_part = name.split(" - ", 1)[0].strip() if " - " in name else name
            n = _normalize_beo_number(first_part)
            n = n.zfill(5) if n.isdigit() else n
            if n == normalized_beo_number:
                return name
    return None


def _file_exists_in_onedrive(access_token: str, drive_id: str, path: str) -> Tuple[bool, Optional[str]]:
    """Check if a file exists at the given path in OneDrive. Returns (exists, webUrl or None)."""
    path_encoded = quote(path, safe="/")
    get_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path_encoded}"
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        resp = client.get(get_url, headers={"Authorization": f"Bearer {access_token}"})
    if resp.status_code == 200:
        data = resp.json()
        return (True, data.get("webUrl"))
    return (False, None)


def _rename_onedrive_folder(
    access_token: str, drive_id: str, year: str, month: str, day: str, current_folder_name: str, new_folder_name: str
) -> None:
    """Rename a folder in OneDrive. No-op if current == new or on error."""
    if (new_folder_name or "").strip() == (current_folder_name or "").strip():
        return
    folder_path = f"{year}/{month}/{day}/{current_folder_name}"
    path_encoded = quote(folder_path, safe="/")
    get_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path_encoded}"
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        resp = client.get(get_url, headers={"Authorization": f"Bearer {access_token}"})
    if resp.status_code != 200:
        return
    item_id = resp.json().get("id")
    if not item_id:
        return
    patch_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
    with httpx.Client(timeout=GRAPH_TIMEOUT) as client:
        r = client.patch(
            patch_url,
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            json={"name": new_folder_name},
        )
    if r.status_code not in (200, 204):
        return


def _upload_pdf_to_onedrive(
    access_token: str,
    drive_id: str,
    path_segments: Tuple[str, str, str, str],
    filename: str,
    pdf_bytes: bytes,
    canonical_folder_name: Optional[str] = None,
) -> str:
    """Upload PDF to OneDrive. Reuses existing folder for same BEO. Skips if file already there. Renames folder to canonical if wrong."""
    year, month, day, folder_name = path_segments
    canonical_folder_name = (canonical_folder_name or folder_name).strip()
    normalized_beo = (folder_name.split(" - ", 1)[0].strip() if " - " in folder_name else folder_name)
    normalized_beo = _normalize_beo_number(normalized_beo)
    if normalized_beo.isdigit():
        normalized_beo = normalized_beo.zfill(5)
    existing_folder = _get_existing_beo_folder_for_day(access_token, drive_id, year, month, day, normalized_beo)
    if existing_folder:
        folder_name = existing_folder
    path_segments = (year, month, day, folder_name)
    safe_name = _sanitize_filename(filename) or "document.pdf"
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"
    path = f"{year}/{month}/{day}/{folder_name}/{safe_name}"
    exists, web_url = _file_exists_in_onedrive(access_token, drive_id, path)
    if exists and web_url:
        if folder_name != canonical_folder_name:
            _rename_onedrive_folder(access_token, drive_id, year, month, day, folder_name, canonical_folder_name)
        return web_url
    _ensure_drive_folders(access_token, drive_id, path_segments)
    path_encoded = quote(path, safe="/")
    upload_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path_encoded}:/content"
    with httpx.Client(timeout=GRAPH_UPLOAD_TIMEOUT) as client:
        resp = client.put(
            upload_url,
            headers={"Authorization": f"Bearer {access_token}"},
            content=pdf_bytes,
        )
        resp.raise_for_status()
        data = resp.json()
        result_url = data.get("webUrl", path) or path
    if folder_name != canonical_folder_name:
        _rename_onedrive_folder(access_token, drive_id, year, month, day, folder_name, canonical_folder_name)
    return result_url


def save_pdf_to_folder(pdf_bytes: bytes, folder_path: Path, original_filename: str = "document.pdf") -> Path:
    """Create folder (and parents) if missing, then save PDF. Returns path to saved file."""
    folder_path.mkdir(parents=True, exist_ok=True)
    safe_name = _sanitize_filename(original_filename) or "document.pdf"
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"
    out_path = folder_path / safe_name
    out_path.write_bytes(pdf_bytes)
    return out_path


def process_pdf(
    pdf_bytes: bytes,
    original_filename: str = "document.pdf",
    access_token: Optional[str] = None,
) -> Tuple[Optional[Union[Path, str]], Optional[dict]]:
    """
    Full pipeline: extract text -> agent analysis -> build path -> save if valid.
    Prefers organization name for folder; falls back to client name and flags for review.
    Returns (saved_path, report_entry). report_entry is set when folder was named by client name or Unknown (for review report).
    """
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if not pdf_text:
        return (None, None)
    result = analyze_pdf_with_agent(pdf_text)
    name_for_folder = (result.organization_name or result.client_name or "").strip() or "Unknown"
    if not result.valid or not result.beo_number or not result.beo_date or not name_for_folder:
        return (None, None)
    flag_reason = None
    if not (result.organization_name or "").strip():
        flag_reason = "used_client_name_fallback" if (result.client_name or "").strip() else "no_organization_or_client"
    path_segments = build_beo_path_segments(result.beo_number, result.beo_date, name_for_folder)
    if not path_segments:
        return (None, None)
    folder_name_used = path_segments[3]
    if _SAVE_TO_ONEDRIVE and access_token:
        drive_id = _get_onedrive_id(access_token)
        canonical_name = path_segments[3]
        saved = _upload_pdf_to_onedrive(
            access_token, drive_id, path_segments, original_filename, pdf_bytes, canonical_folder_name=canonical_name
        )
    else:
        folder_path = build_beo_folder_path(result.beo_number, result.beo_date, name_for_folder)
        if not folder_path:
            return (None, None)
        saved = save_pdf_to_folder(pdf_bytes, folder_path, original_filename)
    report_entry = (
        {
            "attachment_name": original_filename,
            "beo_number": result.beo_number,
            "beo_date": result.beo_date,
            "folder_name_used": folder_name_used,
            "reason": flag_reason or "used_client_name_fallback",
        }
        if flag_reason
        else None
    )
    return (saved, report_entry)
