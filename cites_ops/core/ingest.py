from datetime import date, datetime
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd
from ..utils.helpers import parse_date

class IngestValidator:
    """
    Validates daily intake files and computes snapshot deltas & aging metrics.
    """

    REQUIRED_ISSUE_COLS = {"Id", "Summary", "Description", "Category", "Status"}

    @classmethod
    def validate_issue_csv(cls, file_path: str) -> Tuple[bool, List[str], pd.DataFrame]:
        """
        Validate CSV schema and load into DataFrame.
        Returns (is_valid, errors, dataframe).
        """
        errors = []
        try:
            df = pd.read_csv(file_path, encoding="utf-8-sig", dtype=str)
        except Exception as e:
            try:
                df = pd.read_csv(file_path, encoding="latin1", dtype=str)
            except Exception as e2:
                return False, [f"Failed to read CSV: {e2}"], pd.DataFrame()

        missing = cls.REQUIRED_ISSUE_COLS - set(df.columns)
        if missing:
            errors.append(f"Missing required columns: {sorted(list(missing))}")

        if df.empty:
            errors.append("CSV file contains no data rows.")

        return len(errors) == 0, errors, df

    @classmethod
    def compute_aging(
        cls,
        df: pd.DataFrame,
        as_of_date: Optional[date] = None,
        date_submitted_col: Optional[str] = None,
        date_updated_col: Optional[str] = None,
    ) -> pd.DataFrame:
        """
        Compute age_days and no_movement_days for each ticket.
        Automatically finds the best matching column names.
        """
        ref_date = as_of_date or date.today()
        df_out = df.copy()

        # Resolve submission date column
        sub_col = date_submitted_col
        if not sub_col or sub_col not in df_out.columns:
            for candidate in ("Date Submitted", "date_submitted", "Created", "Submission Date", "date_raised"):
                if candidate in df_out.columns:
                    sub_col = candidate
                    break

        # Resolve updated date column
        upd_col = date_updated_col
        if not upd_col or upd_col not in df_out.columns:
            for candidate in ("Last Update", "Updated", "updated", "date_updated", "Last Modified", "Modified"):
                if candidate in df_out.columns:
                    upd_col = candidate
                    break

        age_days_list = []
        stale_days_list = []

        for _, row in df_out.iterrows():
            sub_date = parse_date(row.get(sub_col)) if sub_col else None
            upd_date = parse_date(row.get(upd_col)) if upd_col else None

            age = (ref_date - sub_date).days if sub_date else None
            stale = (ref_date - upd_date).days if upd_date else age

            age_days_list.append(max(0, age) if age is not None else 0)
            stale_days_list.append(max(0, stale) if stale is not None else 0)

        df_out["age_days"] = age_days_list
        df_out["no_movement_days"] = stale_days_list
        return df_out

    @classmethod
    def compute_daily_delta(
        cls,
        df_current: pd.DataFrame,
        df_previous: Optional[pd.DataFrame] = None,
        id_col: str = "Id",
        status_col: str = "Status",
    ) -> Dict[str, Any]:
        """
        Compare current snapshot with previous snapshot to compute daily delta.
        """
        curr_ids = set(df_current[id_col].dropna().astype(str).str.strip())
        curr_total = len(df_current)

        if df_previous is None or df_previous.empty:
            return {
                "has_previous": False,
                "current_total": curr_total,
                "previous_total": None,
                "new_issues_count": None,
                "closed_issues_count": None,
                "net_change": None,
            }

        prev_ids = set(df_previous[id_col].dropna().astype(str).str.strip())
        prev_total = len(df_previous)

        new_ids = curr_ids - prev_ids
        closed_ids = prev_ids - curr_ids

        return {
            "has_previous": True,
            "current_total": curr_total,
            "previous_total": prev_total,
            "new_issues_count": len(new_ids),
            "closed_issues_count": len(closed_ids),
            "net_change": curr_total - prev_total,
            "new_ids": sorted(list(new_ids)),
            "closed_ids": sorted(list(closed_ids)),
        }
