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
   - **Organization name**: The organization name from the "Client/Organization" field (e.g. "Guardian Scholars Program"). Use the organization/company name only, NOT the individual person or contact name.

Return the structured response with valid=true only when both (a) and (b) are present and you can extract all three fields. Otherwise set valid=false and leave the other fields null."""


class BEOResult(BaseModel):
    """Structured output from the BEO document analyst agent."""

    valid: bool = Field(description="True if document contains signed Hospitality form and BEO, and extraction succeeded.")
    beo_number: Optional[str] = Field(default=None, description="Five-digit BEO number as string, or null.")
    beo_date: Optional[str] = Field(default=None, description="Date of the BEO as string, or null.")
    organization_name: Optional[str] = Field(default=None, description="Organization name from Client/Organization field, or null.")


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
        "and if so, extract the BEO number (five digits), BEO date, and organization name (from Client/Organization field). Return the structured response."
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


def build_beo_path_segments(
    beo_number: str, beo_date_str: Optional[str], organization_name: Optional[str]
) -> Optional[tuple[str, str, str, str]]:
    """
    Build path segments (year, month, day, folder_name) for local or OneDrive.
    Returns None if date cannot be parsed or required fields are missing.
    """
    if not beo_number or not organization_name:
        return None
    date_parts = _parse_beo_date(beo_date_str)
    if not date_parts:
        return None
    year, month, day = date_parts
    beo_num_normalized = str(beo_number).strip()
    if beo_num_normalized.isdigit():
        beo_num_normalized = beo_num_normalized.zfill(5)
    org_safe = _sanitize_filename(organization_name)
    folder_name = f"{beo_num_normalized} - {org_safe}"
    return (str(year), str(month), str(day), folder_name)


def build_beo_folder_path(beo_number: str, beo_date_str: Optional[str], organization_name: Optional[str]) -> Optional[Path]:
    """
    Build local path: {base}/{year}/{month}/{day}/{beo_number} - {organization_name}.
    Returns None if date cannot be parsed or required fields are missing.
    """
    segs = build_beo_path_segments(beo_number, beo_date_str, organization_name)
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


def _upload_pdf_to_onedrive(
    access_token: str, drive_id: str, path_segments: Tuple[str, str, str, str], filename: str, pdf_bytes: bytes
) -> str:
    """Upload PDF to OneDrive at path Year/month/day/folder_name/filename. Returns webUrl or path string.
    One day folder (year/month/day); multiple BEO folders allowed per day (one per event). Skips upload if file exists."""
    year, month, day, folder_name = path_segments
    safe_name = _sanitize_filename(filename) or "document.pdf"
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"
    _ensure_drive_folders(access_token, drive_id, path_segments)
    path = f"{year}/{month}/{day}/{folder_name}/{safe_name}"
    exists, web_url = _file_exists_in_onedrive(access_token, drive_id, path)
    if exists and web_url:
        return web_url
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
        return data.get("webUrl", path) or path


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
) -> Optional[Union[Path, str]]:
    """
    Full pipeline: extract text -> agent analysis -> build path -> save if valid.
    When SAVE_TO_ONEDRIVE is set and access_token is provided, saves to your OneDrive.
    Otherwise saves to local BEO_BASE_PATH.
    Returns Path (local) or str (OneDrive web URL) if valid and saved, else None.
    """
    pdf_text = extract_text_from_pdf(pdf_bytes)
    if not pdf_text:
        return None
    result = analyze_pdf_with_agent(pdf_text)
    if not result.valid or not result.beo_number or not result.beo_date or not result.organization_name:
        return None
    path_segments = build_beo_path_segments(result.beo_number, result.beo_date, result.organization_name)
    if not path_segments:
        return None
    if _SAVE_TO_ONEDRIVE and access_token:
        drive_id = _get_onedrive_id(access_token)
        return _upload_pdf_to_onedrive(access_token, drive_id, path_segments, original_filename, pdf_bytes)
    folder_path = build_beo_folder_path(result.beo_number, result.beo_date, result.organization_name)
    if not folder_path:
        return None
    return save_pdf_to_folder(pdf_bytes, folder_path, original_filename)
