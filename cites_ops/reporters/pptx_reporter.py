from datetime import date
from pathlib import Path
from typing import Any, Dict, Optional, Union
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

class PPTXReporter:
    """
    Generates PowerPoint slide decks (.pptx) for executive management reviews.
    Replaces legacy PowerShell script automation with pure Python.
    """

    COLOR_NAVY = RGBColor(31, 78, 121)    # #1F4E79
    COLOR_DARK = RGBColor(38, 38, 38)     # #262626
    COLOR_MUTED = RGBColor(115, 115, 115) # #737373
    COLOR_WHITE = RGBColor(255, 255, 255)
    COLOR_ACCENT = RGBColor(192, 0, 0)    # Red alert accent

    @classmethod
    def generate_presentation(
        cls,
        df_classified: pd.DataFrame,
        output_path: Union[str, Path],
        report_date: Optional[Union[str, date]] = None,
        title: str = "CITES Operations Intelligence Review",
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        prs = Presentation()
        prs.slide_width = Inches(13.333)  # 16:9 Widescreen
        prs.slide_height = Inches(7.5)

        run_date_str = str(report_date or date.today())

        # Slide 1: Title
        cls._add_title_slide(prs, title, run_date_str)

        # Slide 2: Executive KPIs
        cls._add_kpi_slide(prs, df_classified, run_date_str)

        # Slide 3: Major Category Breakdown
        cls._add_category_slide(prs, df_classified)

        prs.save(out_file)
        return str(out_file)

    @classmethod
    def _add_title_slide(cls, prs: Presentation, title: str, date_str: str) -> None:
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout

        # Title
        tx_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.0))
        tf = tx_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = cls.COLOR_NAVY

        # Subtitle
        p2 = tf.add_paragraph()
        p2.text = f"Daily Decision-Support & Backlog Review · As of {date_str}"
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(20)
        p2.font.color.rgb = cls.COLOR_MUTED

    @classmethod
    def _add_kpi_slide(cls, prs: Presentation, df: pd.DataFrame, date_str: str) -> None:
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Header
        cls._add_slide_header(slide, "Operational Snapshot & Health Overview", f"Snapshot Date: {date_str}")

        total = len(df)
        status_series = df["Status"].astype(str).str.lower() if "Status" in df.columns else pd.Series([])
        resolved = int(status_series.isin(["resolved", "fixed", "closed"]).sum())
        open_cnt = total - resolved

        kpis = [
            ("Total Ingested", str(total), cls.COLOR_NAVY),
            ("Open Backlog", str(open_cnt), cls.COLOR_ACCENT),
            ("Resolved / Closed", str(resolved), cls.COLOR_NAVY),
            ("Resolution Rate", f"{round(100 * resolved / total, 1)}%" if total else "0%", cls.COLOR_NAVY),
        ]

        # Draw 4 KPI Cards
        left_start = 1.0
        card_w = 2.6
        gap = 0.3
        top = 2.0
        card_h = 2.0

        for idx, (label, val, col) in enumerate(kpis):
            x = Inches(left_start + idx * (card_w + gap))
            box = slide.shapes.add_textbox(x, Inches(top), Inches(card_w), Inches(card_h))
            tf = box.text_frame
            tf.word_wrap = True
            
            p_val = tf.paragraphs[0]
            p_val.text = val
            p_val.font.name = "Segoe UI"
            p_val.font.size = Pt(40)
            p_val.font.bold = True
            p_val.font.color.rgb = col
            p_val.alignment = PP_ALIGN.CENTER

            p_lbl = tf.add_paragraph()
            p_lbl.text = label
            p_lbl.font.name = "Segoe UI"
            p_lbl.font.size = Pt(14)
            p_lbl.font.color.rgb = cls.COLOR_MUTED
            p_lbl.alignment = PP_ALIGN.CENTER

    @classmethod
    def _add_category_slide(cls, prs: Presentation, df: pd.DataFrame) -> None:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        cls._add_slide_header(slide, "Top Problem Categories & Driver Breakdown", "Deterministic Text Analysis")

        if "major_topic_label" not in df.columns:
            return

        cat_counts = df["major_topic_label"].value_counts().head(8).reset_index()
        cat_counts.columns = ["Major Category", "Count"]

        rows = len(cat_counts) + 1
        cols = 3
        table_shape = slide.shapes.add_table(rows, cols, Inches(1.0), Inches(1.8), Inches(11.3), Inches(4.5))
        table = table_shape.table

        table.columns[0].width = Inches(6.5)
        table.columns[1].width = Inches(2.4)
        table.columns[2].width = Inches(2.4)

        headers = ["Major Category", "Total Issues", "% of Backlog"]
        total = len(df) or 1
        for col_idx, h in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = cls.COLOR_NAVY
            for p in cell.text_frame.paragraphs:
                p.font.name = "Segoe UI"
                p.font.size = Pt(13)
                p.font.bold = True
                p.font.color.rgb = cls.COLOR_WHITE

        for row_idx, r in cat_counts.iterrows():
            cnt = int(r["Count"])
            pct = f"{round(100 * cnt / total, 1)}%"
            vals = [str(r["Major Category"]), f"{cnt:,}", pct]

            for col_idx, val in enumerate(vals):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = val
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Segoe UI"
                    p.font.size = Pt(12)
                    p.font.color.rgb = cls.COLOR_DARK

    @classmethod
    def _add_slide_header(cls, slide, title: str, subtitle: str) -> None:
        box = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.3), Inches(1.0))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = cls.COLOR_NAVY

        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = "Segoe UI"
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = cls.COLOR_MUTED
