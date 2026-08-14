import re
from typing import Any, Dict, List, Optional
import pandas as pd
from ..utils.config_loader import ConfigLoader

class WorkforceMapper:
    """
    Generic N-tier Workforce Hierarchy and Accountability Mapper.
    Dynamically maps flat ticket queues into multi-level management structures
    driven by teams.csv and routing.yaml (zero hardcoding).
    """

    def __init__(
        self,
        config_path: Optional[str] = None,
        routing_path: Optional[str] = None,
    ):
        self.config = ConfigLoader.get_default_config(config_path)
        self.routing_config = ConfigLoader.get_routing(routing_path).get("routing_groups", {})
        self.hierarchy_cfg = self.config.get("hierarchy", {})
        
        self.resolved_statuses = set(
            s.lower() for s in self.config.get("resolved_statuses", ["resolved", "fixed", "closed"])
        )
        self.open_statuses = set(
            s.lower() for s in self.config.get("open_statuses", ["open", "assigned", "new", "feedback", "reopened"])
        )

        self.routing_patterns = {
            k: re.compile(v.get("pattern", ".*"), re.IGNORECASE)
            for k, v in self.routing_config.items()
        }

    def assign_routing_group(self, assigned_to: str) -> str:
        """Assign a queue name to a routing group (e.g. internal_tech, vendor_tech, field_office)."""
        assigned_str = str(assigned_to or "").strip()
        for group_key, pattern in self.routing_patterns.items():
            if pattern.search(assigned_str):
                return group_key
        return "other"

    @staticmethod
    def _normalize_key(text: Any) -> str:
        if not text or pd.isna(text):
            return ""
        return re.sub(r"[^a-z0-9]", "", str(text).lower()).strip()

    def process_workload(
        self,
        df_issues: pd.DataFrame,
        df_teams: pd.DataFrame,
    ) -> Dict[str, Any]:
        """
        Merge issues with teams mapping and compute hierarchical workload metrics.
        Supports normalized fuzzy category matching and multi-category splits.
        """
        issue_cat_col = self.hierarchy_cfg.get("issue_category_column", "Category")
        team_cat_col = self.hierarchy_cfg.get("team_column", "Team")
        levels = self.hierarchy_cfg.get("levels", ["JD(IS)", "DD(IS)", "Account handled by"])
        status_col = self.hierarchy_cfg.get("issue_status_column", "Status")
        assigned_col = self.hierarchy_cfg.get("issue_assigned_column", "Assigned To")

        # Copy data
        issues = df_issues.copy()

        # Add routing group
        issues["_routing_group"] = issues[assigned_col].apply(self.assign_routing_group)
        issues["_is_resolved"] = issues[status_col].astype(str).str.lower().isin(self.resolved_statuses)
        issues["_is_open"] = ~issues["_is_resolved"]
        issues["_norm_cat"] = issues[issue_cat_col].apply(self._normalize_key)

        # Prepare teams mapping (expand multi-category rows like 'ECR, Employer')
        team_rows = []
        for _, r in df_teams.iterrows():
            raw_teams = str(r.get(team_cat_col, "")).split(",")
            for t in raw_teams:
                norm_t = self._normalize_key(t)
                if norm_t:
                    row_dict = r.to_dict()
                    row_dict["_norm_cat"] = norm_t
                    team_rows.append(row_dict)

        df_teams_expanded = pd.DataFrame(team_rows).drop_duplicates(subset=["_norm_cat"])

        # Merge with teams mapping
        merged = issues.merge(
            df_teams_expanded,
            on="_norm_cat",
            how="left",
        )

        # Fill missing hierarchy levels
        for lvl in levels:
            if lvl in merged.columns:
                merged[lvl] = merged[lvl].fillna("Ownership Not Mapped").astype(str).str.strip()
            else:
                merged[lvl] = "Ownership Not Mapped"

        # Build recursive hierarchy tree
        tree = self._build_recursive_tree(merged, levels, 0)

        # Unmapped categories
        unmapped = merged[merged["_norm_cat"].isin(set(issues["_norm_cat"]) - set(df_teams_expanded["_norm_cat"]))][issue_cat_col].unique()

        # Compute overall KPIs
        kpis = {
            "total_issues": len(merged),
            "open_issues": int(merged["_is_open"].sum()),
            "resolved_issues": int(merged["_is_resolved"].sum()),
            "routing_breakdown": merged["_routing_group"].value_counts().to_dict(),
        }

        return {
            "kpis": kpis,
            "hierarchy_levels": levels,
            "tree": tree,
            "unmapped_categories": sorted(list(unmapped)),
        }

    def _build_recursive_tree(
        self,
        df: pd.DataFrame,
        levels: List[str],
        current_idx: int,
    ) -> List[Dict[str, Any]]:
        if current_idx >= len(levels):
            return []

        level_name = levels[current_idx]
        nodes = []

        grouped = df.groupby(level_name)
        for val, group in grouped:
            node = {
                "level": level_name,
                "name": str(val),
                "total": len(group),
                "open": int(group["_is_open"].sum()),
                "resolved": int(group["_is_resolved"].sum()),
                "routing": group["_routing_group"].value_counts().to_dict(),
                "children": [],
            }
            if current_idx + 1 < len(levels):
                node["children"] = self._build_recursive_tree(group, levels, current_idx + 1)
            nodes.append(node)

        nodes.sort(key=lambda x: x["total"], reverse=True)
        return nodes
