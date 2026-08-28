"""MantisBT REST API client for issue ingestion."""

from __future__ import annotations

import json
import os
import ssl
import urllib.parse
import urllib.request
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd


class MantisClient:
    """
    Client for interacting with MantisBT REST API.
    """

    DEFAULT_URL = "http://localhost:8080/mantisbt/api/rest"
    DEFAULT_TOKEN = ""

    def __init__(
        self,
        base_url: Optional[str] = None,
        token: Optional[str] = None,
        verify_ssl: bool = False,
    ):
        self.base_url = (base_url or os.getenv("MANTIS_API_URL") or self.DEFAULT_URL).rstrip("/")
        self.token = token or os.getenv("MANTIS_API_TOKEN") or self.DEFAULT_TOKEN
        self.verify_ssl = verify_ssl

        # Setup SSL context
        self.ssl_ctx = ssl.create_default_context()
        if not self.verify_ssl:
            self.ssl_ctx.check_hostname = False
            self.ssl_ctx.verify_mode = ssl.CERT_NONE

    def _request(self, endpoint: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Makes an authenticated GET request to the MantisBT REST API."""
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        if params:
            query = urllib.parse.urlencode(params)
            url = f"{url}?{query}"

        headers = {
            "Authorization": self.token,
            "Accept": "application/json",
            "User-Agent": "cites-ops/1.0",
        }

        req = urllib.request.Request(url, headers=headers, method="GET")
        with urllib.request.urlopen(req, context=self.ssl_ctx, timeout=30) as resp:
            data = resp.read().decode("utf-8")
            return json.loads(data)

    def test_connection(self) -> Dict[str, Any]:
        """Tests authentication and returns current user info."""
        return self._request("users/me")

    def get_projects(self) -> List[Dict[str, Any]]:
        """Retrieves list of accessible projects."""
        res = self._request("projects")
        return res.get("projects", [])

    def fetch_issues(
        self,
        project_id: Optional[Union[int, str]] = None,
        page_size: int = 500,
        max_pages: Optional[int] = None,
        verbose: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Fetches issues from MantisBT REST API across all pages.
        """
        all_issues: List[Dict[str, Any]] = []
        page = 1

        while True:
            params: Dict[str, Any] = {"page_size": page_size, "page": page}
            if project_id is not None:
                params["project_id"] = str(project_id)

            try:
                data = self._request("issues", params=params)
            except Exception as e:
                # If page exceeds available data, Mantis may return 400 or empty
                if page > 1 and "400" in str(e):
                    break
                raise e

            issues = data.get("issues", [])
            if not issues:
                break

            all_issues.extend(issues)
            if verbose:
                print(f"  -> Page {page}: fetched {len(issues):,} issues (Total so far: {len(all_issues):,})")

            if len(issues) < page_size:
                break

            page += 1
            if max_pages and page > max_pages:
                break

        return all_issues

    @classmethod
    def to_dataframe(cls, issues: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Transforms MantisBT JSON issues into standard DataFrame matching 'issues YYYY-MM-DD.csv'.
        """
        rows = []
        for issue in issues:
            # Parse dates
            sub_raw = issue.get("created_at") or issue.get("date_submitted") or ""
            upd_raw = issue.get("updated_at") or issue.get("last_updated") or ""

            sub_date = cls._format_iso_date(sub_raw)
            upd_date = cls._format_iso_date(upd_raw)

            raw_id = str(issue.get("id", "") or "").strip()
            formatted_id = raw_id.zfill(7) if raw_id.isdigit() else raw_id

            project_name = (issue.get("project") or {}).get("name", "") if isinstance(issue.get("project"), dict) else str(issue.get("project") or "")
            reporter_name = (issue.get("reporter") or {}).get("name", "") if isinstance(issue.get("reporter"), dict) else str(issue.get("reporter") or "")
            handler_name = (issue.get("handler") or {}).get("name", "") if isinstance(issue.get("handler"), dict) else str(issue.get("handler") or "")
            category_name = (issue.get("category") or {}).get("name", "") if isinstance(issue.get("category"), dict) else str(issue.get("category") or "")
            status_name = (issue.get("status") or {}).get("name", "") if isinstance(issue.get("status"), dict) else str(issue.get("status") or "")
            resolution_name = (issue.get("resolution") or {}).get("name", "") if isinstance(issue.get("resolution"), dict) else str(issue.get("resolution") or "")

            rows.append({
                "Id": formatted_id,
                "Project": project_name,
                "Reporter": reporter_name,
                "Assigned To": handler_name,
                "Category": category_name,
                "Date Submitted": sub_date,
                "Updated": upd_date,
                "Summary": str(issue.get("summary", "") or ""),
                "Status": status_name,
                "Resolution": resolution_name,
                "Fixed in Version": str(issue.get("target_version", "") or ""),
                "Description": str(issue.get("description", "") or ""),
            })

        df = pd.DataFrame(rows, columns=[
            "Id", "Project", "Reporter", "Assigned To", "Category",
            "Date Submitted", "Updated", "Summary", "Status",
            "Resolution", "Fixed in Version", "Description"
        ])
        return df

    @staticmethod
    def _format_iso_date(raw_val: Any) -> str:
        """Formats ISO or UNIX timestamp into YYYY-MM-DD."""
        if not raw_val:
            return ""
        val_str = str(raw_val).strip()
        if val_str.isdigit():
            try:
                return datetime.fromtimestamp(int(val_str)).strftime("%Y-%m-%d")
            except Exception:
                return ""
        for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                clean_str = val_str
                if "+" in clean_str:
                    clean_str = clean_str.split("+")[0]
                elif clean_str.endswith("Z"):
                    clean_str = clean_str[:-1]
                return datetime.fromisoformat(clean_str).strftime("%Y-%m-%d")
            except Exception:
                pass
        return val_str[:10] if len(val_str) >= 10 else val_str
