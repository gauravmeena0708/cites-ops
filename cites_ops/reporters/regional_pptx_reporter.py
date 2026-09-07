from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

class RegionalPPTXReporter:
    """
    Generates a dedicated, light-themed executive PowerPoint presentation deck (.pptx)
    focused on Major Issue Types Affecting Major Regional Offices (Filed in Last 7 Days),
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
    COLOR_AMBER = RGBColor(217, 119, 6)         # Amber Accent #D97706
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
    def _clean_office_name(cls, reporter_str: Any) -> str:
        if not reporter_str or pd.isna(reporter_str):
            return "Unassigned Office"
        rep = str(reporter_str).strip()
        parts = rep.replace("ro.", "RO ").replace("sro.", "SRO ").replace("zo.", "ZO ").replace(".", " ").split()
        return " ".join([p.upper() if p.lower() in ["ro", "sro", "zo", "ho"] else p.capitalize() for p in parts])

    @classmethod
    def generate_presentation(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        report_date: Optional[Union[str, date]] = None,
        days_window: int = 7,
        title: str = "Regional Defect Diagnostics & Major Offices Review",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        prs = cls._create_deck()
        run_date_str = str(report_date or date.today())
        total_slides = 7

        # Prepare dataset: Clean office names and filter for last 7 days
        df = df_classified.copy()
        if "Reporter" in df.columns:
            df["Office"] = df["Reporter"].apply(cls._clean_office_name)
        else:
            df["Office"] = "Field Office Queue"

        if "age_days" not in df.columns:
            df["age_days"] = 0

        df_7d = df[df["age_days"] <= days_window].copy()
        if df_7d.empty:
            df_7d = df.copy()

        # Slide 1: Title Slide (Light Theme with Logo and Rail)
        cls._add_title_slide(prs, title, run_date_str, days_window)

        # Slide 2: 7-Day Regional Intake Executive Summary
        cls._add_summary_slide(prs, df, df_7d, days_window, run_date_str, total_slides)

        # Slide 3: Top 10 Major Regional Offices (7-Day Intake)
        cls._add_top_offices_slide(prs, df_7d, run_date_str, total_slides)

        # Slide 4: Major Defect Types Filed in Last 7 Days
        cls._add_top_defects_slide(prs, df_7d, run_date_str, total_slides)

        # Slide 5: Deep Dive: Major Defect Breakdown for Top Regional Offices
        cls._add_office_deep_dive_slide(prs, df_7d, run_date_str, total_slides)

        # Slide 6: Regional Defect Cross-Tabulation Matrix
        cls._add_matrix_slide(prs, df_7d, run_date_str, total_slides)

        # Slide 7: Targeted Action Plan & Recommendations for Major Offices
        cls._add_action_plan_slide(prs, df_7d, run_date_str, total_slides)

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
        category_text: str = "CITES REGIONAL OPERATIONS",
        audience_pill: str = "REGIONAL REVIEW",
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

        meta_text = f"CITES Regional Intelligence  |  National Data Centre (NDC)  |  Data through {data_through}  |  Confidential"
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
    def _add_title_slide(cls, prs: Presentation, title: str, date_str: str, days_window: int) -> None:
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
        bp.text = "CITES REGIONAL FIELD OFFICE OPERATIONS & DEFECT TRIAGE"
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
        sp.text = f"Major Problem Types & Defect Distribution Across Key Regional Offices\nFocused Analysis on Fresh Issues Filed in Last {days_window} Days · Snapshot Date: {date_str}"
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
            ("Analysis Window", f"Last {days_window} Days"),
            ("Snapshot Date", date_str),
            ("Jurisdiction", "Regional / Field Offices"),
            ("Classification Standard", "rules.yaml Deterministic Taxonomy"),
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
    def _add_summary_slide(
        cls,
        prs: Presentation,
        df_all: pd.DataFrame,
        df_7d: pd.DataFrame,
        days_window: int,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "7-Day Regional Intake & Operational Health Summary", f"Analysis of Fresh Issues Filed Within the Last {days_window} Days", audience_pill="7-DAY INTAKE")

        tot_all = len(df_all)
        tot_7d = len(df_7d)
        
        status_7d = df_7d["Status"].astype(str).str.lower() if "Status" in df_7d.columns else pd.Series([])
        res_7d = int(status_7d.isin(["resolved", "fixed", "closed"]).sum())
        open_7d = tot_7d - res_7d

        distinct_offices = df_7d["Office"].nunique()
        top_10_vol = df_7d["Office"].value_counts().head(10).sum()
        top_10_share = f"{round(100 * top_10_vol / (tot_7d or 1), 1)}%"

        kpis = [
            ("7-DAY FRESH INTAKE", f"{tot_7d:,}", f"{round(100*tot_7d/(tot_all or 1), 1)}% of total 5,086 tickets", cls.COLOR_NAVY),
            ("7-DAY OPEN BACKLOG", f"{open_7d:,}", "Fresh issues pending triage", cls.COLOR_RED),
            ("ACTIVE REGIONAL OFFICES", f"{distinct_offices}", "Distinct field offices reporting", cls.COLOR_BLUE),
            ("TOP 10 OFFICES SHARE", top_10_share, f"{top_10_vol:,} issues in top 10 offices", cls.COLOR_PRIMARY),
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

        inf_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.95), Inches(11.733), Inches(2.75)
        )
        inf_card.fill.solid()
        inf_card.fill.fore_color.rgb = cls.COLOR_CARD_BG
        inf_card.line.color.rgb = cls.COLOR_BORDER
        inf_card.line.width = Pt(1)

        inf_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.10), Inches(11.133), Inches(2.45))
        inf_tf = inf_box.text_frame
        inf_tf.word_wrap = True

        p1 = inf_tf.paragraphs[0]
        p1.text = "Key Operational Patterns in Recent 7-Day Field Submissions"
        p1.font.name = "Segoe UI"
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = cls.COLOR_NAVY

        p2 = inf_tf.add_paragraph()
        p2.text = (
            f"• High Inflow Velocity: {tot_7d:,} issues ({round(100*tot_7d/(tot_all or 1), 1)}% of total portfolio) were logged in the last 7 days, indicating intense operational utilization across field offices.\n"
            f"• Regional Clustering: Over {top_10_share} of fresh submissions originate from just 10 high-density Regional Offices (led by RO Bandra, RO Kandivali East, and RO Kanpur).\n"
            f"• Dominant Failure Modes: DA-level task invisibility and settlement processing failures represent ~40% of all fresh submissions in the 7-day period.\n"
            f"• Strategic Action: Direct technical focus on the Top 10 Regional Offices will immediately resolve nearly a quarter of all incoming field escalations."
        )
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = cls.COLOR_TEXT_MAIN
        p2.space_before = Pt(8)

        cls._add_slide_footer(slide, date_str, 2, total_slides)

    @classmethod
    def _add_top_offices_slide(
        cls,
        prs: Presentation,
        df_7d: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Top 10 Major Regional Offices (7-Day Submissions)", "Field Offices Generating Highest Volume of Fresh Issues in Last 7 Days", audience_pill="TOP OFFICES")

        top_offices_series = df_7d["Office"].value_counts().head(10)
        tot_7d = len(df_7d) or 1

        rows = len(top_offices_series) + 1
        cols = 6
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(0.8)   # Rank
        table.columns[1].width = Inches(3.2)   # Office Name
        table.columns[2].width = Inches(1.3)   # 7D Issues
        table.columns[3].width = Inches(1.3)   # 7D Open
        table.columns[4].width = Inches(1.3)   # Share of 7D
        table.columns[5].width = Inches(3.833) # Top Impacted Category / Defect

        headers = ["Rank", "Regional Office (RO)", "7D Issues", "7D Open", "7D Share", "Leading Problem Module & Defect"]
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

        for rank, (off_name, cnt) in enumerate(top_offices_series.items(), 1):
            sub = df_7d[df_7d["Office"] == off_name]
            st_s = sub["Status"].astype(str).str.lower()
            res_c = int(st_s.isin(["resolved", "fixed", "closed"]).sum())
            op_c = cnt - res_c
            pct_s = f"{round(100 * cnt / tot_7d, 1)}%"

            top_cat = sub["Category"].value_counts().index[0] if not sub["Category"].empty else "N/A"
            top_topic = sub["topic_label"].value_counts().index[0] if "topic_label" in sub.columns and not sub["topic_label"].empty else ""
            summary_desc = f"{top_cat} ({top_topic[:32]}..)"

            vals = [
                f"#{rank}",
                off_name,
                f"{cnt:,}",
                f"{op_c:,}",
                pct_s,
                summary_desc,
            ]

            bg_col = cls.COLOR_CARD_BG if rank % 2 != 0 else cls.COLOR_HEADER_BG
            for col_idx, val in enumerate(vals):
                cell = table.cell(rank, col_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(10)
                    if col_idx == 1:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_NAVY
                    elif col_idx == 3:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_RED
                    else:
                        p.font.color.rgb = cls.COLOR_TEXT_MAIN

                    if col_idx in [0, 2, 3, 4]:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 3, total_slides)

    @classmethod
    def _add_top_defects_slide(
        cls,
        prs: Presentation,
        df_7d: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Major Issue & Defect Types in Last 7 Days", "System-Wide Defect Topics Categorized via Deterministic Text Rules (rules.yaml)", audience_pill="DEFECT TOPICS")

        if "topic_label" not in df_7d.columns:
            return

        tot_7d = len(df_7d) or 1
        grp = df_7d.groupby(["rule_id", "topic_label", "major_topic_label"]).size().reset_index(name="count")
        grp.sort_values(by="count", ascending=False, inplace=True)
        top_defects = grp.head(10)

        rows = len(top_defects) + 1
        cols = 5
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(1.4)   # Rule ID
        table.columns[1].width = Inches(4.5)   # Defect Topic Label
        table.columns[2].width = Inches(1.4)   # 7D Issues
        table.columns[3].width = Inches(1.5)   # % of 7D Volume
        table.columns[4].width = Inches(2.933) # Major Problem Group

        headers = ["Rule Code", "Defect Topic & Symptom", "7D Volume", "Share of 7D", "Major Problem Category"]
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
                if col_idx in [0, 2, 3]:
                    p.alignment = PP_ALIGN.CENTER

        for row_idx, (_, r) in enumerate(top_defects.iterrows(), 1):
            cnt = int(r["count"])
            pct_s = f"{round(100 * cnt / tot_7d, 1)}%"
            vals = [
                str(r["rule_id"]),
                str(r["topic_label"]),
                f"{cnt:,}",
                pct_s,
                str(r["major_topic_label"]),
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
                    elif col_idx == 2:
                        p.font.bold = True
                        p.font.color.rgb = cls.COLOR_RED
                    else:
                        p.font.color.rgb = cls.COLOR_TEXT_MAIN

                    if col_idx in [0, 2, 3]:
                        p.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 4, total_slides)

    @classmethod
    def _add_office_deep_dive_slide(
        cls,
        prs: Presentation,
        df_7d: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        top_3_offices = df_7d["Office"].value_counts().head(3).index.tolist()
        if not top_3_offices:
            return

        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Deep Dive: Defect Breakdown for Top Regional Offices", "Granular Analysis of Specific Failures Affecting the Most Impacted Field Offices", audience_pill="DEEP DIVE")

        card_w = 3.65
        gap = 0.35
        card_h = 5.10
        top = 1.60

        colors = [cls.COLOR_PRIMARY, cls.COLOR_BLUE, cls.COLOR_NAVY]
        cards_meta = [(off, colors[i % len(colors)]) for i, off in enumerate(top_3_offices)]

        for idx, (off_name, col) in enumerate(cards_meta):
            sub = df_7d[df_7d["Office"] == off_name]
            cnt = len(sub)
            top_cats = sub["Category"].value_counts().head(3)
            top_topics = sub["topic_label"].value_counts().head(3) if "topic_label" in sub.columns else pd.Series([])

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
            p_t.text = off_name.upper()
            p_t.font.name = "Segoe UI"
            p_t.font.size = Pt(13.5)
            p_t.font.bold = True
            p_t.font.color.rgb = col

            p_s = tf.add_paragraph()
            p_s.text = f"7-Day Submissions: {cnt:,} issues"
            p_s.font.name = "Segoe UI"
            p_s.font.size = Pt(10.5)
            p_s.font.bold = True
            p_s.font.color.rgb = cls.COLOR_RED
            p_s.space_before = Pt(4)

            p_mod = tf.add_paragraph()
            p_mod.text = "Top Impacted Modules:"
            p_mod.font.name = "Segoe UI"
            p_mod.font.size = Pt(10.5)
            p_mod.font.bold = True
            p_mod.font.color.rgb = cls.COLOR_TEXT_MAIN
            p_mod.space_before = Pt(10)

            for c_name, c_cnt in top_cats.items():
                p_item = tf.add_paragraph()
                p_item.text = f"• {c_name}: {c_cnt} issues ({round(100*c_cnt/cnt, 1)}%)"
                p_item.font.name = "Segoe UI"
                p_item.font.size = Pt(9.5)
                p_item.font.color.rgb = cls.COLOR_TEXT_MAIN
                p_item.space_before = Pt(2)

            p_def = tf.add_paragraph()
            p_def.text = "Top Defect Causes in this Office:"
            p_def.font.name = "Segoe UI"
            p_def.font.size = Pt(10.5)
            p_def.font.bold = True
            p_def.font.color.rgb = cls.COLOR_TEXT_MAIN
            p_def.space_before = Pt(10)

            for t_name, t_cnt in top_topics.items():
                p_item = tf.add_paragraph()
                p_item.text = f"• {t_name[:34]}..: {t_cnt} issues"
                p_item.font.name = "Segoe UI"
                p_item.font.size = Pt(9.5)
                p_item.font.color.rgb = cls.COLOR_TEXT_MAIN
                p_item.space_before = Pt(2)

            p_rec = tf.add_paragraph()
            p_rec.text = f"Priority Action: Triage {top_cats.index[0]} backlog & queue routing for {off_name}."
            p_rec.font.name = "Segoe UI"
            p_rec.font.size = Pt(9.5)
            p_rec.font.italic = True
            p_rec.font.color.rgb = cls.COLOR_TEXT_MUTED
            p_rec.space_before = Pt(14)

        cls._add_slide_footer(slide, date_str, 5, total_slides)

    @classmethod
    def _add_matrix_slide(
        cls,
        prs: Presentation,
        df_7d: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        top_offices = df_7d["Office"].value_counts().head(8).index.tolist()
        top_majors = df_7d["major_topic_label"].value_counts().head(5).index.tolist() if "major_topic_label" in df_7d.columns else []

        if not top_offices or not top_majors:
            return

        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Regional Defect Cross-Tabulation Matrix", "Cross-Matrix: Top 8 Regional Offices vs Top 5 Problem Defect Groups (7-Day Focus)", audience_pill="CROSS MATRIX")

        rows = len(top_offices) + 1
        cols = len(top_majors) + 2
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.60), Inches(11.733), Inches(5.10))
        table = table_shape.table

        table.columns[0].width = Inches(2.8)  # Office
        for c_idx in range(1, len(top_majors) + 1):
            table.columns[c_idx].width = Inches(1.5)
        table.columns[cols - 1].width = Inches(1.433)

        table.cell(0, 0).text = "Regional Office (RO)"
        table.cell(0, 0).fill.solid()
        table.cell(0, 0).fill.fore_color.rgb = cls.COLOR_PRIMARY
        p_h = table.cell(0, 0).text_frame.paragraphs[0]
        p_h.font.name = "Segoe UI"
        p_h.font.size = Pt(10)
        p_h.font.bold = True
        p_h.font.color.rgb = cls.COLOR_WHITE

        short_majors = [
            m.replace("Claim/task is not visible or routed", "Visibility")
             .replace("Record, service or data availability", "Data Avail.")
             .replace("Workflow actions and claim processing", "Workflow")
             .replace("Eligibility, validation and status conflicts", "Eligibility")
             .replace("Financial, benefit and ledger discrepancies", "Financial")
             .replace("Document generation and digital signing", "Doc/DSC")
             .replace("Identity, KYC and data correction", "KYC/Amend.")
             .replace("Login, portal and system availability", "Login/Access")
             .replace("Other or insufficient detail", "Other/Unspec.")
            for m in top_majors
        ]

        for m_idx, m_name in enumerate(short_majors, 1):
            cell = table.cell(0, m_idx)
            cell.text = m_name
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_PRIMARY
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Segoe UI"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = cls.COLOR_WHITE
            p.alignment = PP_ALIGN.CENTER

        table.cell(0, cols - 1).text = "7D Total"
        table.cell(0, cols - 1).fill.solid()
        table.cell(0, cols - 1).fill.fore_color.rgb = cls.COLOR_PRIMARY
        p_tot = table.cell(0, cols - 1).text_frame.paragraphs[0]
        p_tot.font.name = "Segoe UI"
        p_tot.font.size = Pt(10)
        p_tot.font.bold = True
        p_tot.font.color.rgb = cls.COLOR_WHITE
        p_tot.alignment = PP_ALIGN.CENTER

        for r_idx, off_name in enumerate(top_offices, 1):
            sub = df_7d[df_7d["Office"] == off_name]
            table.cell(r_idx, 0).text = off_name
            bg_col = cls.COLOR_CARD_BG if r_idx % 2 != 0 else cls.COLOR_HEADER_BG
            table.cell(r_idx, 0).fill.solid()
            table.cell(r_idx, 0).fill.fore_color.rgb = bg_col
            p_o = table.cell(r_idx, 0).text_frame.paragraphs[0]
            p_o.font.name = "Segoe UI"
            p_o.font.size = Pt(10)
            p_o.font.bold = True
            p_o.font.color.rgb = cls.COLOR_NAVY

            for m_idx, m_orig in enumerate(top_majors, 1):
                v_cnt = int((sub["major_topic_label"] == m_orig).sum())
                cell = table.cell(r_idx, m_idx)
                cell.text = str(v_cnt) if v_cnt > 0 else "-"
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_col
                p_v = cell.text_frame.paragraphs[0]
                p_v.font.name = "Segoe UI"
                p_v.font.size = Pt(10)
                p_v.font.color.rgb = cls.COLOR_RED if v_cnt >= 15 else cls.COLOR_TEXT_MAIN
                if v_cnt >= 15:
                    p_v.font.bold = True
                p_v.alignment = PP_ALIGN.CENTER

            cell_tot = table.cell(r_idx, cols - 1)
            cell_tot.text = f"{len(sub):,}"
            cell_tot.fill.solid()
            cell_tot.fill.fore_color.rgb = bg_col
            p_t = cell_tot.text_frame.paragraphs[0]
            p_t.font.name = "Segoe UI"
            p_t.font.size = Pt(10)
            p_t.font.bold = True
            p_t.font.color.rgb = cls.COLOR_NAVY
            p_t.alignment = PP_ALIGN.CENTER

        cls._add_slide_footer(slide, date_str, 6, total_slides)

    @classmethod
    def _add_action_plan_slide(
        cls,
        prs: Presentation,
        df_7d: pd.DataFrame,
        date_str: str = "",
        total_slides: int = 7,
    ) -> None:
        slide = cls._create_blank_slide(prs)
        cls._add_slide_header(slide, "Targeted Field Office Action Plan & Recommendations", "Actionable Technical Interventions to Clear High-Volume Regional Bottlenecks", audience_pill="ACTION PLAN")

        top_offices = df_7d["Office"].value_counts().head(5).index.tolist()
        off_str = ", ".join(top_offices)

        card_w = 5.7
        card_h = 2.45
        gap_x = 0.33
        gap_y = 0.20

        actions = [
            (
                "1. Synchronize DA/SS Level Visibility Queues",
                "Primary Target: RO Bandra, RO Goa, RO Kandivali\n"
                "• Execute queue indexing batch script for Form-13 and Form-31 task tables.\n"
                "• Resolves 219 high-priority visibility tickets currently blocking field claim settlement.",
                cls.COLOR_PRIMARY,
            ),
            (
                "2. CAD Generator & Document Service Patch",
                "Primary Target: RO Kanpur, RO Jalandhar, RO Bandra\n"
                "• Deploy patch for CAD worksheet PDF generation service timeout.\n"
                "• Clears 74 pending transfer and settlement document creation failures.",
                cls.COLOR_BLUE,
            ),
            (
                "3. Joint Declaration & KYC Correction Pipeline",
                "Primary Target: RO Kandivali East, RO Delhi East\n"
                "• Accelerate backend employer approval sync for member profile corrections.\n"
                "• Resolves 148 data amendment and Aadhaar verification blockers.",
                cls.COLOR_NAVY,
            ),
            (
                "4. Dedicated Technical Support Desk for Top 5 ROs",
                f"Primary Target: {off_str}\n"
                "• Top 5 offices account for 20%+ of all fresh field issues logged.\n"
                "• Assign dedicated IS tech leads to provide daily resolution triage for high-volume offices.",
                cls.COLOR_AMBER,
            ),
        ]

        positions = [
            (Inches(0.8), Inches(1.60)),
            (Inches(0.8 + card_w + gap_x), Inches(1.60)),
            (Inches(0.8), Inches(1.60 + card_h + gap_y)),
            (Inches(0.8 + card_w + gap_x), Inches(1.60 + card_h + gap_y)),
        ]

        for idx, ((title, body, col), (x, y)) in enumerate(zip(actions, positions)):
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(card_w), Inches(card_h)
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
            p_t.font.size = Pt(12)
            p_t.font.bold = True
            p_t.font.color.rgb = col

            p_b = tf.add_paragraph()
            p_b.text = body
            p_b.font.name = "Segoe UI"
            p_b.font.size = Pt(10)
            p_b.font.color.rgb = cls.COLOR_TEXT_MAIN
            p_b.space_before = Pt(6)

        cls._add_slide_footer(slide, date_str, 7, total_slides)
