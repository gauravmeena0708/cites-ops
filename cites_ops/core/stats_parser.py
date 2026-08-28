"""Parser for CITES / Samadhan Setu daily statistics DOCX documents."""

from __future__ import annotations

import hashlib
import re
import unicodedata
import zipfile
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Union, List, Dict, Any
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

MODULE_ALIASES = {
    "appendix_e": ["appendix_e"],
    "vdr_and_vdr_spl": ["vdr_and_vdr_special"],
    "member": ["member"],
    "member_service_amendment": ["member_service_amendment"],
    "joint_declaration": ["joint_declaration"],
    "ecr_employer": ["ecr", "employer"],
    "dsc_esign": ["dsc_e_sign"],
    "pmvbry": ["pmvbry"],
    "cites_other_functionalities": ["cites_other_functionalities"],
    "form_10c": ["form_10c"],
    "sc_surrender": ["scheme_certificate_surrender"],
    "hr_soft_caiu": ["hr_soft", "caiu"],
    "user_management": ["user_management"],
    "website_email": ["website", "email"],
    "form_10d": ["form_10d"],
    "pension_amendment_brs": ["pension_amendment_brs"],
    "lc_jeevan_praman": ["lc_jeevan_praman"],
    "umang_e_proceedings": ["umang", "e_proceeding"],
    "form_19": ["form_19"],
    "form_20": ["form_20"],
    "form_14_lip": ["form_14_lip"],
    "form_13": ["form_13"],
    "form_31": ["form_31"],
    "form_5if": ["form_5if"],
    "annual_accounts": ["annual_accounts"],
    "surrender_and_past_accumulation": ["surrender_of_exem_and_past_accumulations"],
    "iwu": ["iwu"],
    "payments": ["payments"],
}


class StatsDocxParser:
    """Extracts date-wise category totals and metrics from daily Samadhan Setu DOCX stats."""

    @staticmethod
    def normalize_module_key(value: str) -> str:
        """Converts arbitrary module names to standard snake_case keys."""
        if not value:
            return ""
        val = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode().casefold()
        val = val.replace("e-proceedings", "e proceedings").replace("e-proceeding", "e proceeding")
        val = re.sub(r"\bform[- ]?(\d+[a-z]*)\b", r"form \1", val)
        val = re.sub(r"[^a-z0-9]+", "_", val).strip("_")
        return val

    @classmethod
    def cell_text(cls, cell_elem: ET.Element) -> str:
        """Extracts text content from a single table cell element."""
        paragraphs = []
        for p in cell_elem.findall(".//w:p", NS):
            text = "".join(node.text or "" for node in p.findall(".//w:t", NS)).strip()
            if text:
                paragraphs.append(text)
        return " | ".join(paragraphs)

    @classmethod
    def extract_tables_from_docx(cls, path: Union[str, Path]) -> List[List[List[str]]]:
        """Reads document.xml inside .docx archive and returns raw table cells."""
        path = Path(path)
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        tables = []
        for table in root.findall(".//w:tbl", NS):
            rows = []
            for tr in table.findall("w:tr", NS):
                rows.append([cls.cell_text(tc) for tc in tr.findall("w:tc", NS)])
            tables.append(rows)
        return tables

    @classmethod
    def parse_integer(cls, value: str) -> int:
        cleaned = re.sub(r"[^0-9-]", "", str(value or ""))
        return int(cleaned) if cleaned not in ("", "-") else 0

    @classmethod
    def parse_file(cls, path: Union[str, Path]) -> Optional[Dict[str, Any]]:
        """Parses a single stats .docx file and returns structured category metrics."""
        path = Path(path)
        if not path.is_file():
            return None

        tables = cls.extract_tables_from_docx(path)
        data_date: Optional[date] = None
        totals: Optional[Dict[str, int]] = None
        categories: List[Dict[str, Any]] = []

        for table in tables:
            for row in table:
                if len(row) >= 6:
                    for cell in row:
                        cell_clean = cell.strip()
                        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
                            try:
                                parsed = datetime.strptime(cell_clean, fmt).date()
                                if data_date is None:
                                    data_date = parsed
                                    break
                            except ValueError:
                                pass
                    if data_date and totals is None and len(row) >= 6:
                        totals = {
                            "open": cls.parse_integer(row[2]),
                            "resolved": cls.parse_integer(row[3]),
                            "closed": cls.parse_integer(row[4]),
                            "total": cls.parse_integer(row[5]),
                        }

                if len(row) >= 5 and row[0].strip() and row[0].strip().casefold() not in {"by category", "total", "sl no", "s.no"}:
                    if all(re.fullmatch(r"[\d,]+", val.strip()) for val in row[1:5]):
                        mod_label = row[0].strip()
                        categories.append({
                            "module_key": cls.normalize_module_key(mod_label),
                            "module_label": mod_label,
                            "open": cls.parse_integer(row[1]),
                            "resolved": cls.parse_integer(row[2]),
                            "closed": cls.parse_integer(row[3]),
                            "total": cls.parse_integer(row[4]),
                        })

        if not data_date and not categories:
            # Fallback date from filename if document table didn't have explicit date header
            match = re.search(r"(\d{2})[-_](\d{2})[-_](\d{4})", path.name)
            if match:
                try:
                    data_date = datetime.strptime(f"{match.group(1)}-{match.group(2)}-{match.group(3)}", "%d-%m-%Y").date()
                except ValueError:
                    pass

        if not data_date or not categories:
            return None

        # Compute hash
        digest = hashlib.sha256()
        with path.open("rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                digest.update(chunk)

        return {
            "source": {
                "name": path.name,
                "path": str(path.resolve()),
                "sha256": digest.hexdigest(),
                "rows": len(categories),
                "data_date": data_date.isoformat(),
            },
            "totals": totals or {"open": 0, "resolved": 0, "closed": 0, "total": sum(c["total"] for c in categories)},
            "categories": categories,
        }

    @classmethod
    def parse_directory(cls, dir_path: Union[str, Path]) -> List[Dict[str, Any]]:
        """Finds and parses all stats .docx files in a given directory."""
        dir_path = Path(dir_path)
        if not dir_path.is_dir():
            return []

        results = []
        for file_path in dir_path.rglob("*.docx"):
            if file_path.name.startswith("~$"):
                continue
            if "stat" in file_path.name.lower():
                parsed = cls.parse_file(file_path)
                if parsed:
                    results.append(parsed)

        results.sort(key=lambda x: x["source"]["data_date"])
        return results
