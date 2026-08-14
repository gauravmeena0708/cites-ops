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

    def process_workload(
        self,
        df_issues: pd.DataFrame,
        df_teams: pd.DataFrame,
    ) -> Dict[str, Any]:
        """
        Merge issues with teams mapping and compute hierarchical workload metrics.
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

        # Merge with teams mapping
        merged = issues.merge(
            df_teams,
            left_on=issue_cat_col,
            right_on=team_cat_col,
            how="left",
        )

        # Fill missing hierarchy levels
        for lvl in levels:
            if lvl in merged.columns:
                merged[lvl] = merged[lvl].fillna("Not Specified").astype(str).str.strip()
            else:
                merged[lvl] = "Not Specified"

        # Build recursive hierarchy tree
        tree = self._build_recursive_tree(merged, levels, 0)

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
            "unmapped_categories": sorted(
                list(merged[merged[team_cat_col].isna()][issue_cat_col].unique())
            ) if team_cat_col in merged.columns else [],
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
