import re
from typing import Any, Dict, List, Optional
import pandas as pd
from ..utils.config_loader import ConfigLoader

class WorkforceMapper:
    """
    Generic N-tier Workforce Hierarchy, Accountability and Topical Defect Mapper.
    Dynamically maps flat ticket queues into multi-level management structures
    (JD -> DD -> AO/EO/DPA/Programmer -> Category -> Specific Problem Topic)
    driven by teams.csv / Issue_teams.csv and rules.yaml.
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

    @staticmethod
    def _clean_person_name(text: Any) -> str:
        if not text or pd.isna(text):
            return "Not Specified"
        val = str(text).strip()
        lines = [line.strip() for line in val.splitlines() if line.strip()]
        return ", ".join(lines) if lines else "Not Specified"

    def process_workload(
        self,
        df_issues: pd.DataFrame,
        df_teams: pd.DataFrame,
    ) -> Dict[str, Any]:
        """
        Merge issues with teams mapping and compute hierarchical workload and topical metrics.
        """
        issue_cat_col = self.hierarchy_cfg.get("issue_category_column", "Category")
        team_cat_col = self.hierarchy_cfg.get("team_column", "Team")
        levels = self.hierarchy_cfg.get("levels", ["JD(IS)", "DD(IS)", "Account handled by", "Category"])
        status_col = self.hierarchy_cfg.get("issue_status_column", "Status")
        assigned_col = self.hierarchy_cfg.get("issue_assigned_column", "Assigned To")
        summary_col = self.hierarchy_cfg.get("issue_summary_column", "Summary")
        id_col = self.hierarchy_cfg.get("issue_id_column", "Id")

        # Copy data
        issues = df_issues.copy()

        # Add routing group & status flags
        issues["_routing_group"] = issues[assigned_col].apply(self.assign_routing_group) if assigned_col in issues.columns else "internal_tech"
        issues["_is_resolved"] = issues[status_col].astype(str).str.lower().isin(self.resolved_statuses) if status_col in issues.columns else False
        issues["_is_open"] = ~issues["_is_resolved"]
        issues["_norm_cat"] = issues[issue_cat_col].apply(self._normalize_key) if issue_cat_col in issues.columns else ""

        # Prepare teams mapping
        team_rows = []
        for _, r in df_teams.iterrows():
            raw_teams = str(r.get(team_cat_col, "")).split(",")
            for t in raw_teams:
                norm_t = self._normalize_key(t)
                if norm_t:
                    row_dict = r.to_dict()
                    row_dict["_norm_cat"] = norm_t
                    for lvl in ["Account handled by", "DD(IS)", "JD(IS)"]:
                        if lvl in row_dict:
                            row_dict[lvl] = self._clean_person_name(row_dict[lvl])
                    team_rows.append(row_dict)

        df_teams_expanded = pd.DataFrame(team_rows).drop_duplicates(subset=["_norm_cat"])

        # Merge with teams mapping
        merged = issues.merge(
            df_teams_expanded,
            on="_norm_cat",
            how="left",
            suffixes=("", "_team"),
        )

        # Fill missing hierarchy levels
        for lvl in levels:
            if lvl in merged.columns:
                merged[lvl] = merged[lvl].fillna("Ownership Not Mapped").astype(str).str.strip()
                merged[lvl] = merged[lvl].replace("", "Ownership Not Mapped")
            else:
                merged[lvl] = "Ownership Not Mapped"

        # Overall KPIs
        total_cnt = len(merged)
        open_cnt = int(merged["_is_open"].sum())
        resolved_cnt = int(merged["_is_resolved"].sum())
        res_rate = f"{round(100 * resolved_cnt / total_cnt, 1)}%" if total_cnt > 0 else "0%"
        routing_counts = merged["_routing_group"].value_counts().to_dict()

        # Unmapped categories
        unmapped = merged[merged["_norm_cat"].isin(set(issues["_norm_cat"]) - set(df_teams_expanded["_norm_cat"]))][issue_cat_col].unique()
        coverage_pct = f"{round(100 * (total_cnt - len(merged[merged['_norm_cat'].isin(set(issues['_norm_cat']) - set(df_teams_expanded['_norm_cat']))])) / total_cnt, 1)}%" if total_cnt > 0 else "100%"

        kpis = {
            "total_issues": total_cnt,
            "open_issues": open_cnt,
            "resolved_issues": resolved_cnt,
            "resolution_rate": res_rate,
            "coverage_pct": coverage_pct,
            "routing_breakdown": routing_counts,
        }

        # 1. Flat Category Summaries (All & Top 10)
        cat_group = merged.groupby(issue_cat_col)
        cat_summary_rows = []
        module_topical_map = {}

        for cat_name, group in cat_group:
            c_tot = len(group)
            c_open = int(group["_is_open"].sum())
            c_res = int(group["_is_resolved"].sum())
            c_rate = f"{round(100 * c_res / c_tot, 1)}%" if c_tot > 0 else "0%"
            c_share = f"{round(100 * c_open / (open_cnt or 1), 1)}%"
            
            h_val = group["Account handled by"].iloc[0] if "Account handled by" in group.columns else "Not Mapped"
            dd_val = group["DD(IS)"].iloc[0] if "DD(IS)" in group.columns else "Not Mapped"
            jd_val = group["JD(IS)"].iloc[0] if "JD(IS)" in group.columns else "Not Mapped"
            
            cdac_cnt = int((group["_routing_group"] == "vendor_tech").sum())
            ro_cnt = int((group["_routing_group"] == "field_office").sum())
            epfo_cnt = int((group["_routing_group"] == "internal_tech").sum())

            # Topical Breakdown for this specific module
            topics_list = []
            if "topic_label" in group.columns and "major_topic_label" in group.columns:
                t_grp = group.groupby(["major_topic_label", "topic_label", "rule_id"])
                for (maj_lbl, min_lbl, r_code), t_sub in t_grp:
                    t_tot = len(t_sub)
                    t_op = int(t_sub["_is_open"].sum())
                    t_re = t_tot - t_op
                    t_sh = f"{round(100 * t_tot / c_tot, 1)}%"
                    
                    # Top 3 sample issues
                    sample_issues = []
                    for _, s_row in t_sub.head(3).iterrows():
                        sample_issues.append({
                            "id": str(s_row.get(id_col, "")),
                            "summary": str(s_row.get(summary_col, ""))[:100],
                            "status": str(s_row.get(status_col, "")),
                        })

                    topics_list.append({
                        "major_topic_label": maj_lbl,
                        "topic_label": min_lbl,
                        "rule_id": r_code,
                        "total": t_tot,
                        "open": t_op,
                        "resolved": t_re,
                        "share_of_module": t_sh,
                        "samples": sample_issues,
                    })

                topics_list.sort(key=lambda x: (x["open"], x["total"]), reverse=True)

            module_topical_map[str(cat_name)] = topics_list

            cat_summary_rows.append({
                "category": str(cat_name),
                "total": c_tot,
                "open": c_open,
                "resolved": c_res,
                "resolution_rate": c_rate,
                "share_of_backlog": c_share,
                "handler": h_val,
                "dd": dd_val,
                "jd": jd_val,
                "cdac_count": cdac_cnt,
                "ro_count": ro_cnt,
                "epfo_count": epfo_cnt,
                "top_topics": topics_list[:5],
            })

        cat_summary_rows.sort(key=lambda x: (x["open"], x["total"]), reverse=True)

        top_10_categories = []
        for rank, row_data in enumerate(cat_summary_rows[:10], start=1):
            item = dict(row_data)
            item["rank"] = rank
            top_10_categories.append(item)

        # 2. Cross-Tabulation: Top 10 Categories x Major Problem Groups
        top_cat_names = [c["category"] for c in top_10_categories]
        cross_tab = {}
        major_cols = []
        if "major_topic_label" in merged.columns:
            major_cols = sorted(merged["major_topic_label"].unique().tolist())
            ct_df = pd.crosstab(merged[merged[issue_cat_col].isin(top_cat_names)][issue_cat_col], merged["major_topic_label"])
            cross_tab = ct_df.to_dict(orient="index")

        # 3. System-Wide Top Systemic Root Causes
        top_systemic_defects = []
        if "topic_label" in merged.columns and "major_topic_label" in merged.columns:
            sys_grp = merged.groupby(["major_topic_label", "topic_label", "rule_id"])
            for (maj_lbl, min_lbl, r_code), s_sub in sys_grp:
                s_tot = len(s_sub)
                s_op = int(s_sub["_is_open"].sum())
                s_re = s_tot - s_op
                s_desc = s_sub["category_description"].iloc[0] if "category_description" in s_sub.columns else ""
                top_systemic_defects.append({
                    "major_topic_label": maj_lbl,
                    "topic_label": min_lbl,
                    "rule_id": r_code,
                    "description": s_desc,
                    "total": s_tot,
                    "open": s_op,
                    "resolved": s_re,
                    "share_of_total": f"{round(100 * s_tot / total_cnt, 1)}%",
                })
            top_systemic_defects.sort(key=lambda x: (x["open"], x["total"]), reverse=True)

        # 4. 5-Tier Hierarchical Tree and Rows (JD -> DD -> Handler -> Category -> Specific Topic)
        tree = self._build_recursive_tree_with_topics(merged, levels, 0, module_topical_map)
        hierarchy_rows = self._flatten_tree_to_rows(tree)

        return {
            "kpis": kpis,
            "hierarchy_levels": levels,
            "top_10_categories": top_10_categories,
            "category_summary": cat_summary_rows,
            "module_topical_map": module_topical_map,
            "cross_tab_matrix": cross_tab,
            "major_columns": major_cols,
            "top_systemic_defects": top_systemic_defects[:10],
            "hierarchy_rows": hierarchy_rows,
            "tree": tree,
            "unmapped_categories": sorted(list(unmapped)),
            "merged_df": merged,
        }

    def _build_recursive_tree_with_topics(
        self,
        df: pd.DataFrame,
        levels: List[str],
        current_idx: int,
        module_topical_map: Dict[str, List[Dict[str, Any]]],
        parent_id: str = "root",
    ) -> List[Dict[str, Any]]:
        if current_idx >= len(levels):
            return []

        level_name = levels[current_idx]
        nodes = []

        grouped = df.groupby(level_name)
        for idx, (val, group) in enumerate(grouped):
            node_id = f"{parent_id}_{current_idx}_{idx}"
            tot = len(group)
            op = int(group["_is_open"].sum())
            res = int(group["_is_resolved"].sum())
            rate = f"{round(100 * res / tot, 1)}%" if tot > 0 else "0%"
            
            cdac_cnt = int((group["_routing_group"] == "vendor_tech").sum())
            ro_cnt = int((group["_routing_group"] == "field_office").sum())
            epfo_cnt = int((group["_routing_group"] == "internal_tech").sum())

            level_label = level_name
            if level_name == "Account handled by":
                level_label = "AO / EO / DPA / Programmer"

            node = {
                "id": node_id,
                "parent_id": parent_id,
                "depth": current_idx,
                "level_key": level_name,
                "level_label": level_label,
                "name": str(val),
                "total": tot,
                "open": op,
                "resolved": res,
                "resolution_rate": rate,
                "cdac_count": cdac_cnt,
                "ro_count": ro_cnt,
                "epfo_count": epfo_cnt,
                "routing": group["_routing_group"].value_counts().to_dict(),
                "children": [],
            }

            if current_idx + 1 < len(levels):
                node["children"] = self._build_recursive_tree_with_topics(
                    group, levels, current_idx + 1, module_topical_map, node_id
                )
            elif current_idx == len(levels) - 1:
                # Attach leaf topical nodes for this category
                cat_key = str(val)
                top_topics = module_topical_map.get(cat_key, [])
                topic_children = []
                for t_idx, top_item in enumerate(top_topics[:6]):
                    t_id = f"{node_id}_t_{t_idx}"
                    topic_children.append({
                        "id": t_id,
                        "parent_id": node_id,
                        "depth": current_idx + 1,
                        "level_key": "Topic",
                        "level_label": f"Topic: {top_item['rule_id']}",
                        "name": f"[{top_item['rule_id']}] {top_item['topic_label']} ({top_item['share_of_module']})",
                        "total": top_item["total"],
                        "open": top_item["open"],
                        "resolved": top_item["resolved"],
                        "resolution_rate": f"{round(100 * top_item['resolved'] / (top_item['total'] or 1), 1)}%",
                        "cdac_count": 0,
                        "ro_count": 0,
                        "epfo_count": top_item["total"],
                        "routing": {},
                        "children": [],
                    })
                node["children"] = topic_children

            nodes.append(node)

        nodes.sort(key=lambda x: (x["open"], x["total"]), reverse=True)
        return nodes

    def _flatten_tree_to_rows(
        self,
        tree: List[Dict[str, Any]],
        rows: Optional[List[Dict[str, Any]]] = None,
    ) -> List[Dict[str, Any]]:
        if rows is None:
            rows = []

        for node in tree:
            has_children = len(node.get("children", [])) > 0
            rows.append({
                "id": node["id"],
                "parent_id": node["parent_id"],
                "depth": node["depth"],
                "level_key": node["level_key"],
                "level_label": node["level_label"],
                "name": node["name"],
                "total": node["total"],
                "open": node["open"],
                "resolved": node["resolved"],
                "resolution_rate": node["resolution_rate"],
                "cdac_count": node["cdac_count"],
                "ro_count": node["ro_count"],
                "epfo_count": node["epfo_count"],
                "has_children": has_children,
            })
            if has_children:
                self._flatten_tree_to_rows(node["children"], rows)

        return rows
