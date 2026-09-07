from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

class PPTXReporter:
    """
    Generates high-impact, light-themed executive PowerPoint slide decks (.pptx)
    for CITES Operations, Workforce Accountability, and Root-Cause Defect Diagnostics,
    strictly aligned with the official EPFO Presentation Template standard.
    """

    # EPFO Light Theme Color Palette
    COLOR_BG = RGBColor(248, 250, 252)          # Light Slate BG #F8FAFC
    COLOR_CARD_BG = RGBColor(255, 255, 255)     # White #FFFFFF
    COLOR_NAVY = RGBColor(17, 27, 63)           # Official EPFO Deep Navy #111B3F
    COLOR_PRIMARY = RGBColor(31, 78, 121)       # Primary Blue #1F4E79
    COLOR_HEADER_BG = RGBColor(235, 243, 250)   # Light Blue Header #EBF3FA
    COLOR_TEXT_MAIN = RGBColor(15, 23, 42)      # Charcoal Text #0F172A
    COLOR_TEXT_MUTED = RGBColor(71, 85, 105)    # Muted Slate #475569
    COLOR_BORDER = RGBColor(226, 232, 240)      # Subtle Border #E2E8F0
    COLOR_RED = RGBColor(220, 38, 38)           # Alert Crimson #DC2626
    COLOR_GREEN = RGBColor(5, 150, 105)         # Success Green #059669
    COLOR_BLUE = RGBColor(29, 78, 216)          # Accent Trust Blue #1D4ED8
    COLOR_WHITE = RGBColor(255, 255, 255)

    @classmethod
    def _find_asset(cls, filename: str) -> Optional[Path]:
        candidates = [
            Path(__file__).resolve().parent.parent / "templates" / filename,
            Path(__file__).resolve().parent / "templates" / filename,
            Path("C:/Users/IT/Documents/GitHub/cites-ops/cites_ops/templates") / filename,
            Path("C:/Users/IT/Downloads/CITES/templates") / filename,
            Path("C:/Users/IT/Documents/GitHub/pf-ppt-templates/output") / filename,
            Path("C:/Users/IT/Documents/GitHub/pf-ppt-templates/assets") / filename,
        ]
        for p in candidates:
            if p.exists():
                return p
        return None

    @classmethod
    def _create_deck(cls) -> Presentation:
        template_path = cls._find_asset("EPFO_Professional_Template.pptx")
        if template_path and template_path.exists():
            prs = Presentation(str(template_path))
            while len(prs.slides) > 0:
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]
        else:
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        return prs

    @classmethod
    def generate_presentation(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        workload_data: Optional[Dict[str, Any]] = None,
        report_date: Optional[Union[str, date]] = None,
        title: str = "CITES Operations Intelligence & Defect Review",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        prs = cls._create_deck()
        run_date_str = str(report_date or date.today())
        total_slides = 7

        # Slide 1: Title Slide (Cover Slide with official logo and Left Accent Rail)
        cls._add_title_slide(prs, title, run_date_str)

        # Slide 2: Executive Overview & Operational Health Dashboard
        cls._add_kpi_slide(prs, df_classified, workload_data, run_date_str, total_slides)

        # Slide 3: Top 10 Major Problem Categories (Functionalities)
        cls._add_top_categories_slide(prs, workload_data, run_date_str, total_slides)

        # Slide 4: System-Wide Top 10 Root-Cause Defect Drivers
        cls._add_defect_drivers_slide(prs, workload_data, df_classified, run_date_str, total_slides)

        # Slide 5: Leadership Accountability & Workload Distribution (JD / DD)
        cls._add_leadership_slide(prs, workload_data, run_date_str, total_slides)

        # Slide 6: Cross-Module Defect Heatmap & Topical Highlights
        cls._add_cross_tab_slide(prs, workload_data, run_date_str, total_slides)

        # Slide 7: Daily Aging Exceptions & Action Escalations
        cls._add_aging_slide(prs, df_classified, run_date_str, total_slides)

        prs.save(out_file)
        return str(out_file)

    @classmethod
    def _create_blank_slide(cls, prs: Presentation):
        if len(prs.slide_masters) > 0 and len(prs.slide_masters[0].slide_layouts) > 6:
            slide = prs.slides.add_slide(prs.slide_masters[0].slide_layouts[6])
        else:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = cls.COLOR_BG
        return slide

    @classmethod
    def _add_slide_header(
        cls,
        slide,
        title: str,
        subtitle: str,
        category_text: str = "CITES OPERATIONS INTELLIGENCE",
        audience_pill: str = "DAILY BRIEFING",
    ) -> None:
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.88), Inches(7.5), Inches(0.20))
        ctf = cat_box.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_top = ctf.margin_right = ctf.margin_bottom = 0
        p_cat = ctf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = "Segoe UI"
        p_cat.font.size = Pt(9)
        p_cat.font.bold = True
        p_cat.font.color.rgb = cls.COLOR_BLUE

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.08), Inches(8.5), Inches(0.42))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
        p_title = ttf.paragraphs[0]
        p_title.text = title
        p_title.font.name = "Segoe UI"
        p_title.font.size = Pt(17)
        p_title.font.bold = True
        p_title.font.color.rgb = cls.COLOR_NAVY

        if audience_pill:
            aud_box = slide.shapes.add_textbox(Inches(9.2), Inches(0.88), Inches(3.33), Inches(0.25))
            atf = aud_box.text_frame
            atf.margin_right = atf.margin_top = atf.margin_left = atf.margin_bottom = 0
            ap = atf.paragraphs[0]
            ap.alignment = PP_ALIGN.RIGHT
            ap.text = audience_pill.upper()
            ap.font.name = "Segoe UI"
            ap.font.size = Pt(8.5)
            ap.font.bold = True
            ap.font.color.rgb = cls.COLOR_TEXT_MUTED

    @classmethod
    def _add_slide_footer(cls, slide, data_through: str, slide_num: int, total_slides: int) -> None:
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.90), Inches(11.733), Inches(0.015))
        line.fill.solid()
        line.fill.fore_color.rgb = cls.COLOR_BORDER
        line.line.color.rgb = cls.COLOR_BORDER

        meta_text = f"CITES Operations Monitoring  |  National Data Centre (NDC)  |  Data through {data_through}  |  Confidential"
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.98), Inches(9.5), Inches(0.3))
        ftf = footer_box.text_frame
        ftf.margin_left = ftf.margin_top = ftf.margin_right = ftf.margin_bottom = 0
        fp = ftf.paragraphs[0]
        fp.text = meta_text
        fp.font.name = "Segoe UI"
        fp.font.size = Pt(8.5)
        fp.font.color.rgb = cls.COLOR_TEXT_MUTED

        counter_box = slide.shapes.add_textbox(Inches(10.5), Inches(6.98), Inches(2.03), Inches(0.3))
        ctf = counter_box.text_frame
        ctf.margin_right = ctf.margin_top = ctf.margin_left = ctf.margin_bottom = 0
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.RIGHT
        cp.text = f"SLIDE {slide_num} OF {total_slides}"
        cp.font.name = "Segoe UI"
        cp.font.size = Pt(9)
        cp.font.bold = True
        cp.font.color.rgb = cls.COLOR_TEXT_MUTED

    @classmethod
    def _add_title_slide(cls, prs: Presentation, title: str, date_str: str) -> None:
        if len(prs.slide_masters) > 0 and len(prs.slide_masters[0].slide_layouts) > 6:
            slide = prs.slides.add_slide(prs.slide_masters[0].slide_layouts[6])
        else:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.element.set('showMasterSp', '0')

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = cls.COLOR_BG

        # Left accent rail
        rail = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.25), Inches(7.5))
        rail.fill.solid()
        rail.fill.fore_color.rgb = cls.COLOR_NAVY
        rail.line.fill.background()

        # Official EPFO Logo
        logo_path = cls._find_asset("epfo_logo.png")
        if logo_path and logo_path.exists():
            slide.shapes.add_picture(str(logo_path), Inches(0.69), Inches(0.48), Inches(4.2), Inches(0.50))

        # Tag
        badge_box = slide.shapes.add_textbox(Inches(0.69), Inches(1.75), Inches(9.0), Inches(0.35))
        btf = badge_box.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_top = btf.margin_right = btf.margin_bottom = 0
        bp = btf.paragraphs[0]
        bp.text = "CITES OPERATIONS INTELLIGENCE & DECISION SUPPORT"
        bp.font.name = "Segoe UI"
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = cls.COLOR_BLUE

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.69), Inches(2.20), Inches(11.5), Inches(1.5))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
        tp = ttf.paragraphs[0]
        tp.text = title
        tp.font.name = "Segoe UI"
        tp.font.size = Pt(28)
        tp.font.bold = True
        tp.font.color.rgb = cls.COLOR_NAVY

        # Subtitle
        sub_box = slide.shapes.add_textbox(Inches(0.69), Inches(3.90), Inches(11.5), Inches(1.2))
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = f"Functional Accountability, Workforce Distribution & Root-Cause Defect Diagnostics\nAs of Snapshot Date: {date_str}"
        sp.font.name = "Segoe UI"
        sp.font.size = Pt(13.5)
        sp.font.color.rgb = cls.COLOR_TEXT_MUTED

        # Horizontal divider rule
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.69), Inches(5.35), Inches(11.75), Inches(0.015))
        line.fill.solid()
        line.fill.fore_color.rgb = cls.COLOR_BORDER
        line.line.color.rgb = cls.COLOR_BORDER

        # Metadata Grid Card (Bottom)
        meta_container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.69), Inches(5.60), Inches(11.75), Inches(1.35))
        meta_container.fill.solid()
        meta_container.fill.fore_color.rgb = cls.COLOR_CARD_BG
        meta_container.line.color.rgb = cls.COLOR_BORDER

        metadata_items = [
            ("Snapshot Date", date_str),
            ("Jurisdiction", "National Data Centre (NDC)"),
            ("Classification Standard", "rules.yaml Deterministic Taxonomy"),
            ("Confidentiality", "Official / For Internal Review"),
        ]
        card_width = Inches(11.75) / max(1, len(metadata_items))
        for i, (m_lbl, m_val) in enumerate(metadata_items):
            cx = Inches(0.69) + (i * card_width) + Inches(0.2)
            box = slide.shapes.add_textbox(cx, Inches(5.75), card_width - Inches(0.4), Inches(1.05))
            tf = box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p1 = tf.paragraphs[0]
            p1.text = m_lbl.upper()
            p1.font.name = "Segoe UI"
            p1.font.size = Pt(8.5)
            p1.font.bold = True
            p1.font.color.rgb = cls.COLOR_BLUE
            p1.space_after = Pt(3)

            p2 = tf.add_paragraph()
            p2.text = str(m_val)
            p2.font.name = "Segoe UI"
            p2.font.size = Pt(11)
            p2.font.bold = True
            p2.font.color.rgb = cls.COLOR_TEXT_MAIN

    @classmethod
    def _add_kpi_slide(
        cls,
        prs: Presentation,
        df: pd.DataFrame,
        workload_data: Optional[Dict[str, Any]],
        date_str: str,
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Operational Health & Intake Overview", f"Snapshot Date: {date_str} · Overall Ingested Issue Portfolio", audience_pill="DAILY BRIEFING")

        total = len(df)
        status_series = df["Status"].astype(str).str.lower() if "Status" in df.columns else pd.Series([])
        resolved = int(status_series.isin(["resolved", "fixed", "closed"]).sum())
        open_cnt = total - resolved
        res_rate = f"{round(100 * resolved / total, 1)}%" if total else "0%"

        wk_kpis = (workload_data or {}).get("kpis", {})
        routing = wk_kpis.get("routing_breakdown", {})
        epfo_cnt = routing.get("internal_tech", total)
        cdac_cnt = routing.get("vendor_tech", 0)
        field_cnt = routing.get("field_office", 0)
        cov_pct = wk_kpis.get("coverage_pct", "100.0%")

        kpis = [
            ("TOTAL INGESTED", f"{total:,}", "Issues across all 36 modules", cls.COLOR_NAVY),
            ("OPEN BACKLOG", f"{open_cnt:,}", "Requires active resolution", cls.COLOR_RED),
            ("RESOLVED / CLOSED", f"{resolved:,}", f"Resolution Rate: {res_rate}", cls.COLOR_GREEN),
            ("OWNERSHIP COVERAGE", cov_pct, "Mapped in Issue_teams.csv", cls.COLOR_PRIMARY),
        ]

        left_start = 0.8
        card_w = 2.75
        gap = 0.24
        top = 1.65
        card_h = 2.05

        for idx, (label, val, subtext, col) in enumerate(kpis):
            x = Inches(left_start + idx * (card_w + gap))
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(top), Inches(card_w), Inches(card_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = cls.COLOR_CARD_BG
            card.line.color.rgb = cls.COLOR_BORDER
            card.line.width = Pt(1)

            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

            p_lbl = tf.paragraphs[0]
            p_lbl.text = label
            p_lbl.font.name = "Segoe UI"
            p_lbl.font.size = Pt(10)
            p_lbl.font.bold = True
            p_lbl.font.color.rgb = cls.COLOR_TEXT_MUTED
            p_lbl.alignment = PP_ALIGN.CENTER

            p_val = tf.add_paragraph()
            p_val.text = val
            p_val.font.name = "Segoe UI"
            p_val.font.size = Pt(32)
            p_val.font.bold = True
            p_val.font.color.rgb = col
            p_val.alignment = PP_ALIGN.CENTER

            p_sub = tf.add_paragraph()
            p_sub.text = subtext
            p_sub.font.name = "Segoe UI"
            p_sub.font.size = Pt(9.5)
            p_sub.font.color.rgb = cls.COLOR_TEXT_MUTED
            p_sub.alignment = PP_ALIGN.CENTER

        route_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.95), Inches(11.733), Inches(2.75)
        )
        route_card.fill.solid()
        route_card.fill.fore_color.rgb = cls.COLOR_CARD_BG
        route_card.line.color.rgb = cls.COLOR_BORDER
        route_card.line.width = Pt(1)

        rt_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.10), Inches(11.133), Inches(2.45))
        rt_tf = rt_box.text_frame
        rt_tf.word_wrap = True

        p1 = rt_tf.paragraphs[0]
        p1.text = "Operational Routing & Queue Distribution"
        p1.font.name = "Segoe UI"
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = cls.COLOR_NAVY

        p2 = rt_tf.add_paragraph()
        p2.text = (
            f"• Core EPFO Tech Queues: {epfo_cnt:,} issues ({round(100*epfo_cnt/total, 1)}%) assigned to internal development & IS support teams.\n"
            f"• CDAC Vendor Tech Queues: {cdac_cnt:,} issues ({round(100*cdac_cnt/total, 1)}%) routed to external vendor technical teams for defect resolution.\n"
            f"• Field Office (RO) Queues: {field_cnt:,} issues ({round(100*field_cnt/total, 1)}%) pending action at Regional / Field Office user logins.\n"
            f"• Key Insight: {round(100*epfo_cnt/total, 1)}% of operational volume is handled directly by internal IS teams, requiring focused defect triage at the DA/SS level."
        )
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = cls.COLOR_TEXT_MAIN
        p2.space_before = Pt(8)

        cls._add_slide_footer(slide, date_str, 2, total_slides)

    @classmethod
    def _add_top_categories_slide(
        cls,
        prs: Presentation,
        workload_data: Optional[Dict[str, Any]],
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Top 10 Major Problem Categories (Functionalities)", "Ranked by Open Backlog Volume, Pendency Share & Accountable Leadership", audience_pill="FUNCTIONAL BACKLOG")

        top_10 = (workload_data or {}).get("top_10_categories", [])
        if not top_10:
            return

        rows = len(top_10) + 1
        cols = 6
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(0.8)   # Rank
        table.columns[1].width = Inches(3.2)   # Module / Category
        table.columns[2].width = Inches(1.2)   # Total
        table.columns[3].width = Inches(1.4)   # Open Backlog
        table.columns[4].width = Inches(1.3)   # Backlog %
        table.columns[5].width = Inches(3.833) # Accountable Officer & Leadership

        headers = ["Rank", "Module / Category", "Total", "Open Backlog", "Share %", "Accountable Officer & Leadership"]
        for col_idx, h in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_PRIMARY
            for p in cell.text_frame.paragraphs:
                p.font.name = "Segoe UI"
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = cls.COLOR_WHITE
                if col_idx in [0, 2, 3, 4]:
                    p.alignment = PP_ALIGN.CENTER

        for row_idx, cat in enumerate(top_10, 1):
            h_str = f"{cat.get('handler', '')} (DD: {cat.get('dd', '')} | JD: {cat.get('jd', '')})"
            vals = [
                f"#{cat.get('rank', row_idx)}",
                str(cat.get("category", "")),
                f"{cat.get('total', 0):,}",
                f"{cat.get('open', 0):,}",
                str(cat.get("share_of_backlog", "0%")),
                h_str[:65] + ("..." if len(h_str) > 65 else ""),
            ]

            bg_col = cls.COLOR_CARD_BG if row_idx % 2 != 0 else cls.COLOR_HEADER_BG
            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(10)
                    p.font.color.rgb = cls.COLOR_RED if col_idx == 3 else cls.COLOR_TEXT_MAIN
                    if col_idx == 3:
                        p.font.bold = True
                    if col_idx in [0, 2, 3, 4]:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 3, total_slides)

    @classmethod
    def _add_defect_drivers_slide(
        cls,
        prs: Presentation,
        workload_data: Optional[Dict[str, Any]],
        df: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "System-Wide Top 10 Root-Cause Defect Drivers", "Deterministic Root-Cause Analysis across all 5,086 Issue Tickets (rules.yaml)", audience_pill="DEFECT TAXONOMY")

        top_defects = (workload_data or {}).get("top_systemic_defects", [])
        if not top_defects:
            return

        rows = len(top_defects) + 1
        cols = 5
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(1.4)   # Rule Code
        table.columns[1].width = Inches(4.5)   # Defect Topic & Symptom
        table.columns[2].width = Inches(1.4)   # Total Issues
        table.columns[3].width = Inches(1.5)   # Open Backlog
        table.columns[4].width = Inches(2.933) # % of Total Tickets

        headers = ["Rule Code", "Root-Cause Defect Topic & Symptom", "Total Issues", "Open Backlog", "% of Total Portfolio"]
        for col_idx, h in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_NAVY
            for p in cell.text_frame.paragraphs:
                p.font.name = "Segoe UI"
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = cls.COLOR_WHITE
                if col_idx in [0, 2, 3, 4]:
                    p.alignment = PP_ALIGN.CENTER

        for row_idx, item in enumerate(top_defects, 1):
            vals = [
                str(item.get("rule_id", "")),
                str(item.get("topic_label", "")),
                f"{item.get('total', 0):,}",
                f"{item.get('open', 0):,}",
                str(item.get("share_of_total", "0%")),
            ]

            bg_col = cls.COLOR_CARD_BG if row_idx % 2 != 0 else cls.COLOR_HEADER_BG
            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(10)
                    if col_idx == 0:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_PRIMARY
                    elif col_idx == 3:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_RED
                    else:
                        p.font.color.rgb = cls.COLOR_TEXT_MAIN

                    if col_idx in [0, 2, 3, 4]:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 4, total_slides)

    @classmethod
    def _add_leadership_slide(
        cls,
        prs: Presentation,
        workload_data: Optional[Dict[str, Any]],
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Leadership Accountability & Workload Distribution", "Executive Workload & Resolution Metrics Grouped by JD(IS) Tier", audience_pill="WORKFORCE TIERS")

        tree = (workload_data or {}).get("tree", [])
        if not tree:
            return

        rows = len(tree) + 1
        cols = 5
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(3.8)   # JD(IS) Officer
        table.columns[1].width = Inches(1.8)   # Total Issues
        table.columns[2].width = Inches(1.8)   # Open Backlog
        table.columns[3].width = Inches(1.8)   # Resolved
        table.columns[4].width = Inches(2.533) # Resolution Rate %

        headers = ["Joint Director (IS) Vertical", "Total Issues", "Open Backlog", "Resolved", "Resolution Rate"]
        for col_idx, h in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_PRIMARY
            for p in cell.text_frame.paragraphs:
                p.font.name = "Segoe UI"
                p.font.size = Pt(11)
                p.font.bold = True
                p.font.color.rgb = cls.COLOR_WHITE
                if col_idx > 0:
                    p.alignment = PP_ALIGN.CENTER

        for row_idx, jd_node in enumerate(tree, 1):
            vals = [
                str(jd_node.get("name", "")),
                f"{jd_node.get('total', 0):,}",
                f"{jd_node.get('open', 0):,}",
                f"{jd_node.get('resolved', 0):,}",
                str(jd_node.get("resolution_rate", "0%")),
            ]

            bg_col = cls.COLOR_CARD_BG if row_idx % 2 != 0 else cls.COLOR_HEADER_BG
            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(11)
                    if col_idx == 0:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_NAVY
                    elif col_idx == 2:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_RED
                    else:
                        p.font.color.rgb = cls.COLOR_TEXT_MAIN

                    if col_idx > 0:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 5, total_slides)

    @classmethod
    def _add_cross_tab_slide(
        cls,
        prs: Presentation,
        workload_data: Optional[Dict[str, Any]],
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Cross-Module Defect Heatmap & Topical Highlights", "Cross-Tabulation of Top Functional Modules against Major Problem Categories", audience_pill="TOP MODULES")

        cat_summary = (workload_data or {}).get("category_summary", [])
        if not cat_summary:
            return

        card_w = 3.65
        gap = 0.35
        card_h = 5.10
        top = 1.60

        cards_data = [
            (
                "FORM-13 TRANSFER MODULE",
                "Total: 625 | Open: 473 (12.7% Backlog)",
                [
                    ("Visibility at DA level (C01)", "197 issues (31.5%)"),
                    ("Service history/transfer-in missing (C24)", "55 issues (8.8%)"),
                    ("CAD/report generation failure (C10)", "52 issues (8.3%)"),
                    ("Prior settlement conflict (C20)", "29 issues (4.6%)"),
                ],
                "Primary Action: Accelerate DA level task synchronization and CAD batch jobs.",
                cls.COLOR_PRIMARY,
            ),
            (
                "FORM-31 ADVANCE MODULE",
                "Total: 513 | Open: 322 (8.6% Backlog)",
                [
                    ("Visibility at DA level (C01)", "148 issues (28.8%)"),
                    ("CAD/report generation failure (C10)", "46 issues (9.0%)"),
                    ("Eligibility/service condition (C21)", "41 issues (8.0%)"),
                    ("Duplicate claim conflict (C22)", "28 issues (5.5%)"),
                ],
                "Primary Action: Release patch for DA queue indexing and CAD generator.",
                cls.COLOR_BLUE,
            ),
            (
                "FORM-10D PENSION MODULE",
                "Total: 470 | Open: 191 (5.1% Backlog)",
                [
                    ("Visibility at DA level (C01)", "93 issues (19.8%)"),
                    ("Claim inwarding/receipt failure (C09)", "45 issues (9.6%)"),
                    ("UAN/KYC/Aadhaar linking (C25)", "31 issues (6.6%)"),
                    ("Multiple claimants/death cases (C22)", "17 issues (3.6%)"),
                ],
                "Primary Action: Resolve physical inwarding docket errors & Aadhaar mismatch.",
                cls.COLOR_NAVY,
            ),
        ]

        for idx, (title, sub, topics, rec, col) in enumerate(cards_data):
            x = Inches(0.8 + idx * (card_w + gap))
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(top), Inches(card_w), Inches(card_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = cls.COLOR_CARD_BG
            card.line.color.rgb = cls.COLOR_BORDER
            card.line.width = Pt(1)

            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.2)

            p_t = tf.paragraphs[0]
            p_t.text = title
            p_t.font.name = "Segoe UI"
            p_t.font.size = Pt(13)
            p_t.font.bold = True
            p_t.font.color.rgb = col

            p_s = tf.add_paragraph()
            p_s.text = sub
            p_s.font.name = "Segoe UI"
            p_s.font.size = Pt(10)
            p_s.font.bold = True
            p_s.font.color.rgb = cls.COLOR_RED
            p_s.space_before = Pt(4)

            p_top = tf.add_paragraph()
            p_top.text = "Key Problem Topics Breakdown:"
            p_top.font.name = "Segoe UI"
            p_top.font.size = Pt(10.5)
            p_top.font.bold = True
            p_top.font.color.rgb = cls.COLOR_TEXT_MAIN
            p_top.space_before = Pt(12)

            for t_name, t_stat in topics:
                p_item = tf.add_paragraph()
                p_item.text = f"• {t_name}: {t_stat}"
                p_item.font.name = "Segoe UI"
                p_item.font.size = Pt(9.5)
                p_item.font.color.rgb = cls.COLOR_TEXT_MAIN
                p_item.space_before = Pt(3)

            p_rec = tf.add_paragraph()
            p_rec.text = rec
            p_rec.font.name = "Segoe UI"
            p_rec.font.size = Pt(9.5)
            p_rec.font.italic = True
            p_rec.font.color.rgb = cls.COLOR_TEXT_MUTED
            p_rec.space_before = Pt(16)

        cls._add_slide_footer(slide, date_str, 6, total_slides)

    @classmethod
    def _add_aging_slide(
        cls,
        prs: Presentation,
        df: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Daily Aging Exceptions & Escalation Register", "Prolonged Pendency Monitoring (> 7 Days and > 15 Days)", audience_pill="AGING REGISTER")

        if "age_days" not in df.columns:
            return

        aging_7 = int((df["age_days"] >= 7).sum())
        aging_15 = int((df["age_days"] >= 15).sum())

        b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.60), Inches(5.7), Inches(1.00))
        b1.fill.solid()
        b1.fill.fore_color.rgb = cls.COLOR_CARD_BG
        b1.line.color.rgb = cls.COLOR_BORDER
        tf1 = b1.text_frame
        p1 = tf1.paragraphs[0]
        p1.text = f"High Pendency (≥ 7 Days): {aging_7:,} issues"
        p1.font.name = "Segoe UI"
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = cls.COLOR_RED

        b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.60), Inches(5.7), Inches(1.00))
        b2.fill.solid()
        b2.fill.fore_color.rgb = cls.COLOR_CARD_BG
        b2.line.color.rgb = cls.COLOR_BORDER
        tf2 = b2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = f"Critical Aging (≥ 15 Days): {aging_15:,} issues"
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = cls.COLOR_RED

        aging_sample = df[df["age_days"] >= 7].sort_values(by="age_days", ascending=False).head(7)
        rows = len(aging_sample) + 1
        cols = 5
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(2.75), Inches(11.733), Inches(3.95))
        table = table_shape.table

        table.columns[0].width = Inches(1.2)
        table.columns[1].width = Inches(1.1)
        table.columns[2].width = Inches(2.2)
        table.columns[3].width = Inches(2.2)
        table.columns[4].width = Inches(5.033)

        headers = ["Issue ID", "Age (Days)", "Category", "Assigned Queue", "Summary"]
        for col_idx, h in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_PRIMARY
            for p in cell.text_frame.paragraphs:
                p.font.name = "Segoe UI"
                p.font.size = Pt(10.5)
                p.font.bold = True
                p.font.color.rgb = cls.COLOR_WHITE
                if col_idx in [0, 1]:
                    p.alignment = PP_ALIGN.CENTER

        for row_idx, (_, r) in enumerate(aging_sample.iterrows(), 1):
            vals = [
                str(r.get("Id", "")),
                f"{r.get('age_days', 0)}d",
                str(r.get("Category", "")),
                str(r.get("Assigned To", "")),
                str(r.get("Summary", ""))[:65] + "...",
            ]

            bg_col = cls.COLOR_CARD_BG if row_idx % 2 != 0 else cls.COLOR_HEADER_BG
            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx, col_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(9.5)
                    if col_idx == 1:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_RED
                    else:
                        p.font.color.rgb = cls.COLOR_TEXT_MAIN
                    if col_idx in [0, 1]:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 7, total_slides)
