import hashlib
import re
from datetime import datetime, date
from typing import Optional, Union, Tuple

def normalise_text(text: Optional[str]) -> str:
    """Normalise whitespace and lower-case text."""
    if not text:
        return ""
    cleaned = re.sub(r"\s+", " ", str(text)).strip().lower()
    return re.sub(r"[^a-z0-9 ]", "", cleaned).strip()

def text_sha256(text: Optional[str]) -> str:
    """Compute SHA-256 hash of normalised text."""
    normalised = normalise_text(text)
    return hashlib.sha256(normalised.encode("utf-8")).hexdigest()

def parse_date(date_val: Union[str, datetime, date, None]) -> Optional[date]:
    """Parse various date formats into date object."""
    if date_val is None:
        return None
    if isinstance(date_val, datetime):
        return date_val.date()
    if isinstance(date_val, date):
        return date_val
    
    val_str = str(date_val).strip()
    if not val_str:
        return None
        
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str.split(" ")[0], fmt).date()
        except ValueError:
            pass
    return None

def format_number(val: Union[int, float, None]) -> str:
    """Format numbers with thousands separators."""
    if val is None:
        return "—"
    try:
        return f"{int(val):,}"
    except (ValueError, TypeError):
        return str(val)

def extract_workflow_level(text: str) -> Tuple[str, str]:
    """Extract processing workflow level (DA, SS, AO/APFC, RPFC-II, RPFC-I) if present."""
    lower = text.lower()
    if re.search(r"\b(?:da|dealing assistant)\b", lower):
        return "da", "Dealing Assistant (DA)"
    if re.search(r"\b(?:ss|section supervisor)\b", lower):
        return "ss", "Section Supervisor (SS)"
    if re.search(r"\b(?:rpfc[- ]?(?:ii|2))\b", lower):
        return "rpfc_ii", "RPFC-II"
    if re.search(r"\b(?:ao|apfc|accounts officer|assistant provident fund commissioner)\b", lower):
        return "ao_apfc", "AO / APFC"
    if re.search(r"\b(?:rpfc(?:[- ]?(?:i|1))?|approver|approving officer|regional provident fund commissioner)\b", lower):
        return "rpfc_i", "RPFC-I / Approver"
    return "unspecified", "Workflow level not stated"
