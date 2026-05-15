"""
output_generator.py
T.A.N.G.S - Output Generation Module

Three generator classes:
  ExcelGenerator       – in-memory .xlsx with 3 styled sheets
  ChartGenerator       – matplotlib PNG charts returned as BytesIO
  PowerPointGenerator  – .pptx with title, summary, and embedded chart slides

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import io
import math
import logging
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# Use non-interactive backend BEFORE pyplot is imported anywhere else.
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.figure import Figure

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# ======================================================================
# THEME DEFINITIONS
# ======================================================================
THEMES: Dict[str, Dict[str, str]] = {
    "corporate_blue": {
        "bg": "FFFFFF",
        "accent": "4472C4",
        "text": "1A1A1A",
    },
    "dark_mode": {
        "bg": "1A1A2E",
        "accent": "E94560",
        "text": "FFFFFF",
    },
    "clean_light": {
        "bg": "F8F9FA",
        "accent": "28A745",
        "text": "212529",
    },
    "academic_gold": {
        "bg": "FFFFFF",
        "accent": "B8860B",
        "text": "1A1A1A",
    },
}

# Chart color palette for pie charts
_PIE_COLORS = [
    "#4472C4", "#ED7D31", "#A9D18E", "#FF6B6B",
    "#7030A0", "#00B0F0", "#92D050", "#FFD700",
    "#FF8C00", "#20B2AA",
]


# ======================================================================
# EXCEL GENERATOR
# ======================================================================
class ExcelGenerator:
    """Generates an in-memory Excel workbook with three styled sheets."""

    # Openpyxl fill objects
    _HEADER_FILL_LIGHT = PatternFill("solid", fgColor="D9E1F2")   # Sheet 1
    _HEADER_FILL_BLUE = PatternFill("solid", fgColor="4472C4")    # Sheet 3
    _ROW_WHITE = PatternFill("solid", fgColor="FFFFFF")
    _ROW_GRAY = PatternFill("solid", fgColor="F2F2F2")
    _SECTION_FILL = PatternFill("solid", fgColor="EBF3FB")

    def generate(self, data: Dict[str, Any], doc_type: str) -> bytes:
        """Build workbook and return bytes.

        Args:
            data:     dict from process_result / process_timetable /
                      process_financial.
            doc_type: "result" | "timetable" | "financial"

        Returns:
            Raw .xlsx bytes suitable for download or disk write.
        """
        if data is None:
            data = {}
        doc_type = (doc_type or "").lower()

        wb = Workbook()

        # ---- Sheet 1: Raw Data ----
        ws1 = wb.active
        ws1.title = "Raw Data"
        main_df = data.get("dataframe", pd.DataFrame())
        self._write_df_sheet(ws1, main_df, blue_header=False)

        # ---- Sheet 2: Summary ----
        ws2 = wb.create_sheet("Summary")
        self._write_summary_sheet(ws2, data, doc_type)

        # ---- Sheet 3: Query Result ----
        ws3 = wb.create_sheet("Query Result")
        self._write_result_sheet(ws3, data, doc_type)

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    # ------------------------------------------------------------------
    def _write_df_sheet(
        self,
        ws,
        df: pd.DataFrame,
        blue_header: bool = False,
        start_row: int = 1,
        title_label: Optional[str] = None,
    ) -> int:
        """Write a DataFrame to a worksheet with styling.

        Returns the next empty row number after the written data.
        """
        if title_label:
            title_cell = ws.cell(row=start_row, column=1, value=title_label)
            title_cell.font = Font(bold=True, size=13, color="4472C4")
            start_row += 1

        if df is None or df.empty:
            cell = ws.cell(row=start_row, column=1, value="No data available.")
            cell.font = Font(italic=True, color="808080")
            return start_row + 1

        # ---- Header row ----
        if blue_header:
            hdr_fill = self._HEADER_FILL_BLUE
            hdr_font = Font(bold=True, size=11, color="FFFFFF")
        else:
            hdr_fill = self._HEADER_FILL_LIGHT
            hdr_font = Font(bold=True, size=11, color="1A1A1A")

        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(
                row=start_row, column=col_idx, value=str(col_name)
            )
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        # ---- Data rows ----
        for row_offset, row_tuple in enumerate(
            df.itertuples(index=False, name=None), 1
        ):
            row_idx = start_row + row_offset
            fill = (
                self._ROW_WHITE if row_idx % 2 == 0 else self._ROW_GRAY
            )
            for col_idx, v in enumerate(row_tuple, 1):
                cell = ws.cell(
                    row=row_idx,
                    column=col_idx,
                    value=self._cell_value(v),
                )
                cell.fill = fill
                cell.alignment = Alignment(
                    vertical="center", wrap_text=False
                )

        # ---- Auto-fit column widths ----
        self._auto_fit_columns(ws, df, header_row=start_row)

        # ---- Freeze the header row ----
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

        return start_row + len(df) + 1

    def _write_summary_sheet(
        self, ws, data: Dict[str, Any], doc_type: str
    ) -> None:
        """Write human-readable summary to Sheet 2."""
        bold_font = Font(bold=True, size=12)
        section_font = Font(bold=True, size=13, color="4472C4")
        value_font = Font(size=11)
        thin_border = Border(
            bottom=Side(style="thin", color="B0B0B0")
        )

        row = 1

        if doc_type == "result":
            summary = data.get("summary_stats", {}) or {}
            # Section title
            t = ws.cell(row=row, column=1, value="RESULT SUMMARY")
            t.font = section_font
            t.fill = self._SECTION_FILL
            ws.merge_cells(
                start_row=row, start_column=1, end_row=row, end_column=5
            )
            row += 2

            # Key metrics
            metrics = [
                ("Total Students",   summary.get("total_students", 0)),
                ("Passed",           summary.get("passed", 0)),
                ("Failed / ATKT",    summary.get("atkt", 0)),
                ("Overall Pass %",   f"{summary.get('overall_pass_percentage', 0.0):.2f}%"),
            ]
            for label, value in metrics:
                c1 = ws.cell(row=row, column=1, value=label)
                c1.font = bold_font
                c1.border = thin_border
                c2 = ws.cell(row=row, column=2, value=value)
                c2.font = value_font
                c2.border = thin_border
                row += 1
            row += 1

            # Per-subject table
            per_sub = summary.get("per_subject", {}) or {}
            fr_data = data.get("failure_rates", {}) or {}
            subjects = sorted(
                set(list(per_sub.keys()) + list(fr_data.keys()))
            )
            if subjects:
                t2 = ws.cell(row=row, column=1, value="SUBJECT STATISTICS")
                t2.font = section_font
                t2.fill = self._SECTION_FILL
                ws.merge_cells(
                    start_row=row, start_column=1,
                    end_row=row, end_column=5,
                )
                row += 1
                # Sub-headers
                sub_headers = ["Subject", "Average", "Max", "Min", "Pass %"]
                for ci, h in enumerate(sub_headers, 1):
                    c = ws.cell(row=row, column=ci, value=h)
                    c.font = Font(bold=True, size=11)
                    c.fill = self._HEADER_FILL_LIGHT
                    c.alignment = Alignment(horizontal="center")
                row += 1
                for i, subj in enumerate(subjects):
                    fill = (
                        self._ROW_WHITE if i % 2 == 0 else self._ROW_GRAY
                    )
                    ps = per_sub.get(subj, {})
                    fr = fr_data.get(subj, {})
                    pass_pct = 100.0 - fr.get("rate", 0.0)
                    vals = [
                        subj,
                        round(ps.get("average", 0), 2),
                        ps.get("max", 0),
                        ps.get("min", 0),
                        f"{pass_pct:.1f}%",
                    ]
                    for ci, v in enumerate(vals, 1):
                        c = ws.cell(row=row, column=ci, value=v)
                        c.fill = fill
                    row += 1

            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 14

        elif doc_type == "timetable":
            t = ws.cell(row=row, column=1, value="TIMETABLE SUMMARY")
            t.font = section_font
            t.fill = self._SECTION_FILL
            ws.merge_cells(
                start_row=row, start_column=1, end_row=row, end_column=6
            )
            row += 2

            wl_df = data.get("workload_df", pd.DataFrame())
            if wl_df is not None and not wl_df.empty:
                row = self._write_df_sheet(
                    ws, wl_df, blue_header=False, start_row=row,
                    title_label="Teacher Workload",
                )
            freq = data.get("subject_freq", {}) or {}
            if freq:
                row += 1
                t2 = ws.cell(row=row, column=1, value="Subject Frequency")
                t2.font = section_font
                row += 1
                for si, (subj, cnt) in enumerate(
                    sorted(freq.items(), key=lambda x: -x[1])
                ):
                    ws.cell(row=row, column=1, value=subj).font = bold_font
                    ws.cell(row=row, column=2, value=cnt)
                    row += 1

        elif doc_type == "financial":
            t = ws.cell(row=row, column=1, value="FINANCIAL KPIs")
            t.font = section_font
            t.fill = self._SECTION_FILL
            ws.merge_cells(
                start_row=row, start_column=1, end_row=row, end_column=3
            )
            row += 2

            kpis = data.get("kpis", {}) or {}
            for kpi_name, kpi_val in kpis.items():
                c1 = ws.cell(
                    row=row, column=1,
                    value=kpi_name.replace("_", " ").title(),
                )
                c1.font = bold_font
                c2 = ws.cell(row=row, column=2, value=kpi_val)
                c2.number_format = "#,##0.00"
                row += 1

            yoy = data.get("yoy_growth")
            if yoy is not None:
                row += 1
                c1 = ws.cell(row=row, column=1, value="YoY Growth")
                c1.font = bold_font
                c2 = ws.cell(row=row, column=2, value=f"{yoy:.2f}%")
                row += 1

            query_ans = data.get("query_answer", "")
            if query_ans:
                row += 1
                ws.cell(row=row, column=1, value="Query Answer").font = bold_font
                row += 1
                ans_cell = ws.cell(row=row, column=1, value=query_ans)
                ans_cell.alignment = Alignment(wrap_text=True)
                ws.row_dimensions[row].height = 60
                ws.merge_cells(
                    start_row=row, start_column=1,
                    end_row=row, end_column=4,
                )

            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 18

        else:
            ws.cell(row=1, column=1, value="No summary available for this document type.")

    def _write_result_sheet(
        self, ws, data: Dict[str, Any], doc_type: str
    ) -> None:
        """Write the query-result DataFrame to Sheet 3 with blue header."""
        result_df = self._get_result_df(data, doc_type)
        if result_df is None or result_df.empty:
            ws.cell(row=1, column=1, value="No query result data available.")
            ws.cell(row=1, column=1).font = Font(italic=True, color="808080")
            return

        # Section label
        label_map = {
            "result": "Query Result — Toppers / KT Students",
            "timetable": "Query Result — Teacher Workload",
            "financial": "Query Result — Financial KPIs",
        }
        self._write_df_sheet(
            ws, result_df, blue_header=True, start_row=1,
            title_label=label_map.get(doc_type, "Query Result"),
        )

        # Append the text answer if available
        query_ans = data.get("query_answer", "")
        if query_ans:
            last_row = len(result_df) + 4
            ws.cell(row=last_row, column=1, value="LLM Answer:").font = (
                Font(bold=True, size=11, color="4472C4")
            )
            ans_cell = ws.cell(row=last_row + 1, column=1, value=query_ans)
            ans_cell.alignment = Alignment(wrap_text=True)
            ws.row_dimensions[last_row + 1].height = 80
            try:
                ws.merge_cells(
                    start_row=last_row + 1, start_column=1,
                    end_row=last_row + 1,
                    end_column=min(len(result_df.columns), 6),
                )
            except Exception:
                pass

    def _get_result_df(
        self, data: Dict[str, Any], doc_type: str
    ) -> pd.DataFrame:
        """Return the most relevant DataFrame for the query result sheet."""
        if doc_type == "result":
            for key in ("topper_df", "kt_df", "dataframe"):
                df = data.get(key)
                if df is not None and not df.empty:
                    return df
        elif doc_type == "timetable":
            for key in ("workload_df", "dataframe"):
                df = data.get(key)
                if df is not None and not df.empty:
                    return df
        elif doc_type == "financial":
            for key in ("dataframe",):
                df = data.get(key)
                if df is not None and not df.empty:
                    return df
        # Final fallback
        df = data.get("dataframe")
        return df if df is not None else pd.DataFrame()

    @staticmethod
    def _cell_value(v: Any) -> Any:
        """Convert a cell value to an openpyxl-safe Python type."""
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        # Unwrap numpy scalars
        if hasattr(v, "item"):
            try:
                return v.item()
            except Exception:
                pass
        # Keep int / float as-is; convert everything else to str
        if isinstance(v, (int, float)):
            if math.isnan(v) if isinstance(v, float) else False:
                return ""
            return v
        if isinstance(v, list):
            return ", ".join(str(x) for x in v)
        return str(v)

    def _auto_fit_columns(
        self, ws, df: pd.DataFrame, header_row: int = 1
    ) -> None:
        """Set column widths based on content length, max 42 units."""
        for col_idx, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            # Header length
            max_len = len(str(col_name))
            # Data lengths (sample first 200 rows for speed)
            col_vals = df.iloc[:200, col_idx - 1]
            for v in col_vals:
                vlen = len(str(v)) if v is not None else 0
                if vlen > max_len:
                    max_len = vlen
            ws.column_dimensions[col_letter].width = min(max_len + 2, 42)


# ======================================================================
# CHART GENERATOR
# ======================================================================
class ChartGenerator:
    """Generates matplotlib charts and returns them as PNG BytesIO objects."""

    _BAR_COLOR = "#4472C4"
    _GRID_COLOR = "#E0E0E0"
    _BG_COLOR = "#FAFAFA"

    # ---- Public API --------------------------------------------------

    def bar_chart(
        self,
        df: pd.DataFrame,
        x_col: str,
        y_col: str,
        title: str = "Bar Chart",
    ) -> BytesIO:
        """Vertical bar chart: x_col on axis, y_col as bar heights."""
        if (
            df is None
            or df.empty
            or x_col not in df.columns
            or y_col not in df.columns
        ):
            return self._empty_chart(title)

        x_vals = df[x_col].astype(str).tolist()
        y_vals = pd.to_numeric(df[y_col], errors="coerce").fillna(0).tolist()
        n = len(x_vals)

        fig_w = max(8, min(16, n * 1.2))
        fig, ax = plt.subplots(figsize=(fig_w, 6))

        positions = range(n)
        bars = ax.bar(positions, y_vals, color=self._BAR_COLOR,
                      edgecolor="white", linewidth=0.5)

        # Value labels on bars
        for bar, val in zip(bars, y_vals):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max(y_vals) * 0.01,
                f"{val:.0f}",
                ha="center", va="bottom", fontsize=8, color="#333333",
            )

        ax.set_xticks(positions)
        rotate = len(x_vals) > 6
        ax.set_xticklabels(
            x_vals,
            rotation=45 if rotate else 0,
            ha="right" if rotate else "center",
            fontsize=9,
        )
        ax.set_title(title, fontsize=14, fontweight="bold", pad=15)
        ax.set_xlabel(str(x_col), fontsize=10)
        ax.set_ylabel(str(y_col), fontsize=10)
        ax.set_xlim(-0.5, n - 0.5)
        self._apply_clean_style(ax)

        plt.tight_layout()
        return self._fig_to_bytes(fig)

    def pie_chart(
        self,
        df: pd.DataFrame,
        labels_col: str,
        values_col: str,
        title: str = "Pie Chart",
    ) -> BytesIO:
        """Pie chart with percentage labels on slices."""
        if (
            df is None
            or df.empty
            or labels_col not in df.columns
            or values_col not in df.columns
        ):
            return self._empty_chart(title)

        labels = df[labels_col].astype(str).tolist()
        values = pd.to_numeric(df[values_col], errors="coerce").fillna(0).tolist()
        # Filter zero/negative values
        pairs = [(l, v) for l, v in zip(labels, values) if v > 0]
        if not pairs:
            return self._empty_chart(title)
        labels, values = zip(*pairs)

        colors = (_PIE_COLORS * ((len(labels) // len(_PIE_COLORS)) + 1))[
            : len(labels)
        ]

        fig, ax = plt.subplots(figsize=(9, 7))
        wedges, texts, autotexts = ax.pie(
            values,
            labels=labels,
            colors=colors,
            autopct="%1.1f%%",
            startangle=140,
            pctdistance=0.82,
            wedgeprops={"edgecolor": "white", "linewidth": 1.5},
        )
        for at in autotexts:
            at.set_fontsize(9)
            at.set_color("white")
            at.set_fontweight("bold")
        for t in texts:
            t.set_fontsize(9)

        ax.set_title(title, fontsize=14, fontweight="bold", pad=20)
        ax.axis("equal")

        plt.tight_layout()
        return self._fig_to_bytes(fig)

    def line_chart(
        self,
        df: pd.DataFrame,
        x_col: str,
        y_col: str,
        title: str = "Line Chart",
    ) -> BytesIO:
        """Line chart with marker dots."""
        if (
            df is None
            or df.empty
            or x_col not in df.columns
            or y_col not in df.columns
        ):
            return self._empty_chart(title)

        x_vals = df[x_col].astype(str).tolist()
        y_vals = pd.to_numeric(df[y_col], errors="coerce").fillna(0).tolist()
        n = len(x_vals)

        fig, ax = plt.subplots(figsize=(10, 5))
        positions = range(n)
        ax.plot(
            positions, y_vals,
            color=self._BAR_COLOR,
            linewidth=2,
            marker="o",
            markersize=7,
            markerfacecolor="white",
            markeredgecolor=self._BAR_COLOR,
            markeredgewidth=2,
        )
        # Value annotations
        for pos, val in zip(positions, y_vals):
            ax.annotate(
                f"{val:.1f}",
                (pos, val),
                textcoords="offset points",
                xytext=(0, 10),
                ha="center",
                fontsize=8,
                color="#333333",
            )

        ax.set_xticks(positions)
        rotate = n > 6
        ax.set_xticklabels(
            x_vals,
            rotation=45 if rotate else 0,
            ha="right" if rotate else "center",
            fontsize=9,
        )
        ax.set_title(title, fontsize=14, fontweight="bold", pad=15)
        ax.set_xlabel(str(x_col), fontsize=10)
        ax.set_ylabel(str(y_col), fontsize=10)
        self._apply_clean_style(ax)

        plt.tight_layout()
        return self._fig_to_bytes(fig)

    def failure_rate_chart(self, failure_rates: Dict[str, Any]) -> BytesIO:
        """Horizontal bar chart, colour-coded by severity."""
        if not failure_rates:
            return self._empty_chart("Subject Failure Rates")

        subjects = list(failure_rates.keys())
        rates = [
            float(failure_rates[s].get("rate", 0)) for s in subjects
        ]
        # Sort by rate descending so highest failure is at top
        pairs = sorted(zip(subjects, rates), key=lambda x: x[1], reverse=True)
        subjects, rates = zip(*pairs) if pairs else ([], [])
        subjects = list(subjects)
        rates = list(rates)

        def _color(r: float) -> str:
            if r < 20:
                return "#28A745"
            if r <= 40:
                return "#FFA500"
            return "#DC3545"

        colors = [_color(r) for r in rates]
        fig_h = max(4.0, len(subjects) * 0.65 + 1.5)
        fig, ax = plt.subplots(figsize=(10, fig_h))

        bars = ax.barh(
            subjects, rates,
            color=colors,
            edgecolor="white",
            linewidth=0.5,
            height=0.55,
        )

        # Percentage labels at end of each bar
        max_rate = max(rates) if rates else 100
        for bar, rate in zip(bars, rates):
            offset = max_rate * 0.01
            ax.text(
                bar.get_width() + offset,
                bar.get_y() + bar.get_height() / 2,
                f"{rate:.1f}%",
                va="center",
                fontsize=9,
                color="#333333",
            )

        ax.set_xlim(0, max_rate * 1.20 if max_rate > 0 else 100)
        ax.set_title(
            "Subject-wise Failure Rate", fontsize=14, fontweight="bold", pad=15
        )
        ax.set_xlabel("Failure Rate (%)", fontsize=10)

        # Legend
        legend_patches = [
            mpatches.Patch(color="#28A745", label="< 20%  Acceptable"),
            mpatches.Patch(color="#FFA500", label="20–40%  Concerning"),
            mpatches.Patch(color="#DC3545", label="> 40%  Critical"),
        ]
        ax.legend(
            handles=legend_patches,
            loc="lower right",
            fontsize=9,
            framealpha=0.9,
        )

        self._apply_clean_style(ax)
        ax.set_facecolor(self._BG_COLOR)
        plt.tight_layout()
        return self._fig_to_bytes(fig)

    # ---- Helpers -----------------------------------------------------

    def _apply_clean_style(self, ax) -> None:
        """Remove top/right spines, add light horizontal grid."""
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#CCCCCC")
        ax.spines["bottom"].set_color("#CCCCCC")
        ax.yaxis.grid(
            True, color=self._GRID_COLOR, linewidth=0.8, linestyle="-"
        )
        ax.set_axisbelow(True)
        ax.set_facecolor(self._BG_COLOR)
        ax.tick_params(axis="both", labelsize=9, colors="#555555")

    def _empty_chart(self, title: str = "Chart") -> BytesIO:
        """Return a placeholder PNG when data is unavailable."""
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.text(
            0.5, 0.5,
            "No data available",
            transform=ax.transAxes,
            ha="center", va="center",
            fontsize=16, color="#AAAAAA",
        )
        ax.set_title(title, fontsize=14, fontweight="bold", color="#888888")
        ax.axis("off")
        plt.tight_layout()
        return self._fig_to_bytes(fig)

    @staticmethod
    def _fig_to_bytes(fig: Figure) -> BytesIO:
        """Save figure to BytesIO PNG and close it."""
        buf = BytesIO()
        try:
            fig.savefig(buf, format="png", dpi=120, bbox_inches="tight",
                        facecolor="white")
        finally:
            plt.close(fig)
        buf.seek(0)
        return buf


# ======================================================================
# POWERPOINT GENERATOR
# ======================================================================
class PowerPointGenerator:
    """Generates an in-memory .pptx presentation with theme support."""

    # Slide canvas dimensions (widescreen 16:9)
    _W = Inches(13.333)
    _H = Inches(7.5)

    def generate(
        self,
        data: Dict[str, Any],
        chart_types: List[str],
        theme: str = "corporate_blue",
    ) -> bytes:
        """Build presentation and return bytes.

        Args:
            data:        dict from a processing module.
            chart_types: list of chart type strings, e.g.
                         ["failure_rate_chart", "bar_chart", "pie_chart"].
            theme:       key from THEMES dict.

        Returns:
            Raw .pptx bytes.
        """
        if data is None:
            data = {}
        theme_cfg = THEMES.get(theme, THEMES["corporate_blue"])
        doc_type = self._infer_doc_type(data)
        query = data.get("query", "")

        prs = Presentation()
        prs.slide_width = self._W
        prs.slide_height = self._H

        # Use the blank slide layout (index 6 in the built-in template)
        try:
            blank_layout = prs.slide_layouts[6]
        except IndexError:
            blank_layout = prs.slide_layouts[0]

        # ---- Slide 1: Title ----
        self._add_title_slide(prs, blank_layout, theme_cfg, doc_type, query)

        # ---- Slide 2: Summary ----
        self._add_summary_slide(prs, blank_layout, theme_cfg, data, doc_type)

        # ---- Slide 3+: Charts ----
        if not chart_types:
            # Placeholder if no charts requested
            self._add_placeholder_slide(
                prs, blank_layout, theme_cfg,
                "No Charts Requested",
                "Add chart_types list to generate chart slides.",
            )
        else:
            for ct in chart_types:
                # _add_chart_slide always adds exactly one slide; exceptions
                # are caught internally and rendered as error text.
                self._add_chart_slide(
                    prs, blank_layout, theme_cfg, data, ct, doc_type
                )

        buf = BytesIO()
        prs.save(buf)
        buf.seek(0)
        return buf.getvalue()

    # ------------------------------------------------------------------
    # SLIDE BUILDERS
    # ------------------------------------------------------------------

    def _add_title_slide(
        self,
        prs: Presentation,
        layout,
        theme_cfg: Dict[str, str],
        doc_type: str,
        query: str,
    ) -> None:
        slide = prs.slides.add_slide(layout)
        self._set_slide_bg(slide, theme_cfg["bg"])
        accent = theme_cfg["accent"]
        txt = theme_cfg["text"]

        # Top accent band
        self._add_rect(slide, 0, 0, 13.333, 1.6, accent, accent)

        # App title in accent band
        self._add_text_box(
            slide, "T.A.N.G.S",
            Inches(0.5), Inches(0.0), Inches(9), Inches(0.9),
            font_size=42, bold=True, color_hex="FFFFFF",
            align=PP_ALIGN.LEFT,
        )
        self._add_text_box(
            slide,
            "Transformative AI-Powered NLP & Generation Suite",
            Inches(0.5), Inches(0.9), Inches(12), Inches(0.55),
            font_size=13, bold=False, color_hex="FFFFFF",
            align=PP_ALIGN.LEFT,
        )

        # Document type heading
        doc_display = (doc_type or "Document").replace("_", " ").title()
        self._add_text_box(
            slide, f"Analysis Report: {doc_display}",
            Inches(1.0), Inches(2.3), Inches(11), Inches(0.85),
            font_size=28, bold=True, color_hex=accent,
            align=PP_ALIGN.LEFT,
        )

        # Generated date
        date_str = datetime.now().strftime("%B %d, %Y")
        self._add_text_box(
            slide, f"Generated on {date_str}",
            Inches(1.0), Inches(3.25), Inches(11), Inches(0.45),
            font_size=12, bold=False, color_hex=txt,
        )

        # Query line
        if query:
            q_display = f'Query: "{query}"'
            if len(q_display) > 110:
                q_display = q_display[:107] + '..."'
            self._add_text_box(
                slide, q_display,
                Inches(1.0), Inches(3.9), Inches(11), Inches(0.7),
                font_size=13, bold=False, color_hex=txt, italic=True,
            )

        # Bottom accent bar
        self._add_rect(slide, 0, 6.9, 13.333, 0.35, accent, accent)

        # Footer
        self._add_text_box(
            slide,
            "SAKEC Mumbai  |  B.Tech Final Year Project  |  2025",
            Inches(0.5), Inches(6.95), Inches(12), Inches(0.35),
            font_size=9, bold=False, color_hex="FFFFFF",
            align=PP_ALIGN.CENTER,
        )

    def _add_summary_slide(
        self,
        prs: Presentation,
        layout,
        theme_cfg: Dict[str, str],
        data: Dict[str, Any],
        doc_type: str,
    ) -> None:
        slide = prs.slides.add_slide(layout)
        self._set_slide_bg(slide, theme_cfg["bg"])
        accent = theme_cfg["accent"]
        txt = theme_cfg["text"]

        # Slide title
        self._add_text_box(
            slide, "Executive Summary",
            Inches(0.5), Inches(0.15), Inches(12), Inches(0.75),
            font_size=28, bold=True, color_hex=accent,
        )
        # Thin accent underline
        self._add_rect(slide, 0.5, 0.92, 12, 0.05, accent, accent)

        if doc_type == "result":
            summary = data.get("summary_stats", {}) or {}
            metrics = [
                ("Total Students",  str(summary.get("total_students", 0))),
                ("Passed",          str(summary.get("passed", 0))),
                ("Failed / ATKT",   str(summary.get("atkt", 0))),
                ("Pass Rate",       f"{summary.get('overall_pass_percentage', 0.0):.1f}%"),
            ]
            # 4 metric cards in a row
            for i, (label, value) in enumerate(metrics):
                col_x = 0.5 + i * 3.2
                # Big value
                self._add_text_box(
                    slide, value,
                    Inches(col_x), Inches(1.2), Inches(3.0), Inches(1.1),
                    font_size=40, bold=True, color_hex=accent,
                    align=PP_ALIGN.CENTER,
                )
                # Label
                self._add_text_box(
                    slide, label,
                    Inches(col_x), Inches(2.3), Inches(3.0), Inches(0.45),
                    font_size=13, bold=False, color_hex=txt,
                    align=PP_ALIGN.CENTER,
                )

            # Per-subject stats as compact text
            per_sub = summary.get("per_subject", {}) or {}
            fr_data = data.get("failure_rates", {}) or {}
            subjects = sorted(set(list(per_sub.keys()) + list(fr_data.keys())))
            if subjects:
                lines = ["Subject Performance (Average Mark | Pass %)"]
                for subj in subjects[:8]:  # cap at 8 for slide space
                    ps = per_sub.get(subj, {})
                    fr = fr_data.get(subj, {})
                    avg = ps.get("average", 0)
                    pass_pct = 100.0 - fr.get("rate", 0.0)
                    lines.append(
                        f"  {subj:<12}  avg={avg:.1f}   pass={pass_pct:.1f}%"
                    )
                self._add_text_box(
                    slide, "\n".join(lines),
                    Inches(0.5), Inches(3.0), Inches(12.3), Inches(3.8),
                    font_size=10, bold=False, color_hex=txt,
                    font_name="Consolas",
                )

        elif doc_type == "timetable":
            wl_df = data.get("workload_df", pd.DataFrame())
            freq = data.get("subject_freq", {}) or {}
            total_teachers = len(wl_df) if wl_df is not None else 0
            total_subjects = len(freq)
            total_lectures = (
                int(wl_df["total_weekly_lectures"].sum())
                if wl_df is not None
                and not wl_df.empty
                and "total_weekly_lectures" in wl_df.columns
                else 0
            )
            metrics = [
                ("Total Teachers",    str(total_teachers)),
                ("Unique Subjects",   str(total_subjects)),
                ("Weekly Lectures",   str(total_lectures)),
            ]
            for i, (label, value) in enumerate(metrics):
                col_x = 0.8 + i * 4.2
                self._add_text_box(
                    slide, value,
                    Inches(col_x), Inches(1.3), Inches(4.0), Inches(1.1),
                    font_size=44, bold=True, color_hex=accent,
                    align=PP_ALIGN.CENTER,
                )
                self._add_text_box(
                    slide, label,
                    Inches(col_x), Inches(2.4), Inches(4.0), Inches(0.45),
                    font_size=14, bold=False, color_hex=txt,
                    align=PP_ALIGN.CENTER,
                )
            # Top 6 teachers
            if wl_df is not None and not wl_df.empty:
                lines = ["Top Teachers by Weekly Lectures:"]
                for _, row in wl_df.head(6).iterrows():
                    name = row.get("teacher_name", "")
                    lects = row.get("total_weekly_lectures", 0)
                    subs = row.get("subjects", "")
                    lines.append(f"  {name:<28} {lects:>3} lectures   [{subs}]")
                self._add_text_box(
                    slide, "\n".join(lines),
                    Inches(0.5), Inches(3.2), Inches(12.3), Inches(3.5),
                    font_size=10, bold=False, color_hex=txt,
                    font_name="Consolas",
                )

        elif doc_type == "financial":
            kpis = data.get("kpis", {}) or {}
            if kpis:
                lines = ["Extracted Financial KPIs:"]
                for k, v in kpis.items():
                    lines.append(
                        f"  {k.replace('_', ' ').title():<25} {v:>15,.2f}"
                    )
                yoy = data.get("yoy_growth")
                if yoy is not None:
                    lines.append(
                        f"  {'YoY Growth':<25} {yoy:>14.2f}%"
                    )
                self._add_text_box(
                    slide, "\n".join(lines),
                    Inches(1.5), Inches(1.3), Inches(10), Inches(5.5),
                    font_size=16, bold=False, color_hex=txt,
                    font_name="Consolas",
                )
            else:
                self._add_text_box(
                    slide, "No financial KPIs extracted.",
                    Inches(1.5), Inches(2.5), Inches(10), Inches(2),
                    font_size=18, bold=False, color_hex="#888888",
                    align=PP_ALIGN.CENTER,
                )

        else:
            self._add_text_box(
                slide, "Summary data not available.",
                Inches(1), Inches(2), Inches(11), Inches(2),
                font_size=18, bold=False, color_hex="#888888",
                align=PP_ALIGN.CENTER,
            )

        # Footer accent line
        self._add_rect(slide, 0, 7.3, 13.333, 0.05, accent, accent)

    def _add_chart_slide(
        self,
        prs: Presentation,
        layout,
        theme_cfg: Dict[str, str],
        data: Dict[str, Any],
        chart_type: str,
        doc_type: str,
    ) -> None:
        # Slide is added unconditionally; any exception below renders as
        # error text inside this same slide (no duplicate slide is created).
        slide = prs.slides.add_slide(layout)
        self._set_slide_bg(slide, theme_cfg["bg"])
        accent = theme_cfg["accent"]
        txt = theme_cfg["text"]

        try:
            # ---- Resolve chart data ----
            df, x_col, y_col, chart_variant, chart_title = self._resolve_chart_data(
                chart_type, data, doc_type
            )

            # ---- Slide title ----
            self._add_text_box(
                slide, chart_title,
                Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.75),
                font_size=24, bold=True, color_hex=accent,
            )
            self._add_rect(slide, 0.4, 0.93, 12.5, 0.05, accent, accent)

            # ---- Generate chart ----
            cg = ChartGenerator()
            if chart_variant == "failure_rate":
                chart_buf = cg.failure_rate_chart(
                    data.get("failure_rates", {}) or {}
                )
            elif chart_variant == "pie":
                chart_buf = cg.pie_chart(
                    df, x_col or "x", y_col or "y", chart_title
                )
            elif chart_variant == "line":
                chart_buf = cg.line_chart(
                    df, x_col or "x", y_col or "y", chart_title
                )
            else:
                chart_buf = cg.bar_chart(
                    df, x_col or "x", y_col or "y", chart_title
                )

            # ---- Embed chart image ----
            slide.shapes.add_picture(
                chart_buf,
                Inches(0.6), Inches(1.1), Inches(12.1), Inches(4.8),
            )

            # ---- Insight text (2 lines) ----
            insight = self._generate_insight(chart_type, data)
            self._add_text_box(
                slide, insight,
                Inches(0.6), Inches(6.1), Inches(12.1), Inches(1.1),
                font_size=10, bold=False, color_hex=txt,
            )

        except Exception as e:
            logger.warning(
                "Chart slide generation failed for %s: %s", chart_type, e
            )
            self._add_text_box(
                slide,
                f"Chart type: {chart_type}\n"
                f"Could not generate chart: {e}\n"
                "Please verify the data dictionary contains the expected keys.",
                Inches(1.5), Inches(2.5), Inches(10), Inches(3),
                font_size=13, bold=False, color_hex="#CC0000",
                align=PP_ALIGN.LEFT,
            )

        # Footer accent line (always)
        self._add_rect(slide, 0, 7.3, 13.333, 0.05, accent, accent)

    def _add_placeholder_slide(
        self,
        prs: Presentation,
        layout,
        theme_cfg: Dict[str, str],
        title: str,
        message: str,
    ) -> None:
        slide = prs.slides.add_slide(layout)
        self._set_slide_bg(slide, theme_cfg["bg"])
        self._add_text_box(
            slide, title,
            Inches(1), Inches(2), Inches(11), Inches(1),
            font_size=24, bold=True, color_hex=theme_cfg["accent"],
            align=PP_ALIGN.CENTER,
        )
        self._add_text_box(
            slide, message,
            Inches(1.5), Inches(3.3), Inches(10), Inches(1.5),
            font_size=14, bold=False, color_hex=theme_cfg["text"],
            align=PP_ALIGN.CENTER,
        )

    # ------------------------------------------------------------------
    # CHART DATA RESOLVER
    # ------------------------------------------------------------------

    def _resolve_chart_data(
        self, chart_type: str, data: Dict[str, Any], doc_type: str
    ) -> Tuple[pd.DataFrame, Optional[str], Optional[str], str, str]:
        """Return (df, x_col, y_col, chart_variant, title) for a chart type.

        chart_variant: "bar" | "pie" | "line" | "failure_rate"
        """
        doc_type = (doc_type or "").lower()
        empty = pd.DataFrame()

        if chart_type == "failure_rate_chart":
            return (
                empty, None, None, "failure_rate",
                "Subject-wise Failure Rate",
            )

        if chart_type in ("bar_chart", "toppers", "bar"):
            if doc_type == "result":
                _td = data.get("topper_df")
                df = (
                    _td
                    if (_td is not None and not _td.empty)
                    else data.get("dataframe", empty)
                )
                if df is not None and not df.empty:
                    x = "name" if "name" in df.columns else df.columns[0]
                    y = "total" if "total" in df.columns else None
                    if y is None:
                        # Pick first numeric column
                        for c in df.columns:
                            if pd.api.types.is_numeric_dtype(df[c]):
                                y = c
                                break
                    if y:
                        return df, x, y, "bar", "Top Students by Total Marks"
            if doc_type == "timetable":
                df = data.get("workload_df", empty)
                if df is not None and not df.empty:
                    return (
                        df,
                        "teacher_name",
                        "total_weekly_lectures",
                        "bar",
                        "Teacher Weekly Workload",
                    )
            if doc_type == "financial":
                kpis = data.get("kpis", {}) or {}
                if kpis:
                    df = pd.DataFrame(
                        list(kpis.items()), columns=["Metric", "Value"]
                    )
                    return df, "Metric", "Value", "bar", "Financial KPIs"
            return empty, None, None, "bar", "Bar Chart"

        if chart_type in ("pie_chart", "pie"):
            if doc_type == "result":
                summary = data.get("summary_stats", {}) or {}
                passed = summary.get("passed", 0)
                atkt = summary.get("atkt", 0)
                if passed + atkt > 0:
                    df = pd.DataFrame({
                        "Status": ["Passed", "Failed / ATKT"],
                        "Count":  [passed, atkt],
                    })
                    return df, "Status", "Count", "pie", "Pass vs Fail Distribution"
            if doc_type == "timetable":
                freq = data.get("subject_freq", {}) or {}
                if freq:
                    top10 = sorted(freq.items(), key=lambda x: -x[1])[:10]
                    df = pd.DataFrame(top10, columns=["Subject", "Count"])
                    return df, "Subject", "Count", "pie", "Subject Distribution"
            return empty, None, None, "pie", "Pie Chart"

        if chart_type in ("line_chart", "line"):
            # Line charts are most useful for financial trends
            df = data.get("dataframe", empty)
            if df is not None and not df.empty and len(df.columns) >= 2:
                # Find a numeric column for y-axis
                x = df.columns[0]
                y = None
                for c in df.columns[1:]:
                    if pd.api.types.is_numeric_dtype(df[c]):
                        y = c
                        break
                if y:
                    return df, x, y, "line", "Trend Analysis"
            return empty, None, None, "line", "Trend Analysis"

        # Unknown chart type – attempt generic bar
        df = data.get("dataframe", empty)
        if df is not None and not df.empty and len(df.columns) >= 2:
            x = df.columns[0]
            for c in df.columns[1:]:
                if pd.api.types.is_numeric_dtype(df[c]):
                    return df, x, c, "bar", chart_type.replace("_", " ").title()
        return empty, None, None, "bar", chart_type.replace("_", " ").title()

    # ------------------------------------------------------------------
    # INSIGHT GENERATOR (rule-based, 2 lines)
    # ------------------------------------------------------------------

    def _generate_insight(
        self, chart_type: str, data: Dict[str, Any]
    ) -> str:
        try:
            if chart_type == "failure_rate_chart":
                fr = data.get("failure_rates", {}) or {}
                if not fr:
                    return (
                        "No failure rate data available.\n"
                        "Ensure the result PDF was parsed correctly."
                    )
                worst_subj, worst_info = max(
                    fr.items(), key=lambda x: x[1].get("rate", 0)
                )
                critical = [
                    s for s, v in fr.items() if v.get("rate", 0) > 40
                ]
                line1 = (
                    f"Highest failure: {worst_subj} at "
                    f"{worst_info.get('rate', 0):.1f}%"
                )
                line2 = (
                    f"{len(critical)} subject(s) exceed 40% failure threshold "
                    f"— immediate academic intervention recommended."
                    if critical
                    else "All subjects within acceptable failure range (≤ 40%)."
                )
                return f"{line1}\n{line2}"

            if chart_type in ("bar_chart", "toppers", "bar"):
                topper_df = data.get("topper_df")
                if topper_df is not None and not topper_df.empty:
                    top_row = topper_df.iloc[0]
                    name = top_row.get("name", "N/A")
                    total = top_row.get("total", 0)
                    line1 = f"Top Performer: {name}  (Total ESE Score: {total:.0f})"
                    line2 = (
                        f"Showing top {len(topper_df)} students ranked by "
                        f"cumulative ESE marks."
                    )
                    return f"{line1}\n{line2}"
                wl_df = data.get("workload_df")
                if wl_df is not None and not wl_df.empty:
                    top = wl_df.iloc[0]
                    line1 = (
                        f"Busiest: {top.get('teacher_name', 'N/A')} — "
                        f"{top.get('total_weekly_lectures', 0)} weekly lectures."
                    )
                    line2 = (
                        f"Total teaching staff: {len(wl_df)}. "
                        f"Workload shown in descending order."
                    )
                    return f"{line1}\n{line2}"

            if chart_type in ("pie_chart", "pie"):
                summary = data.get("summary_stats", {}) or {}
                pass_pct = summary.get("overall_pass_percentage", 0)
                line1 = f"Overall pass rate: {pass_pct:.1f}%"
                line2 = (
                    "Passed segment represents students with all ESE ≥ 40. "
                    "ATKT students require supplementary examination."
                )
                return f"{line1}\n{line2}"

            if chart_type in ("line_chart", "line"):
                return (
                    "Trend analysis based on extracted document data.\n"
                    "Values represent key metrics across the dataset."
                )

            # Fallback
            return (
                "Analysis based on extracted document data.\n"
                "Powered by T.A.N.G.S — SAKEC Mumbai, 2025."
            )

        except Exception as e:
            logger.warning("Insight generation failed: %s", e)
            return (
                "Analysis based on extracted document data.\n"
                "Powered by T.A.N.G.S — SAKEC Mumbai, 2025."
            )

    # ------------------------------------------------------------------
    # SLIDE UTILITY METHODS
    # ------------------------------------------------------------------

    @staticmethod
    def _infer_doc_type(data: Dict[str, Any]) -> str:
        """Guess document type from data dict keys."""
        if not data:
            return ""
        explicit = data.get("doc_type", "")
        if explicit:
            return explicit
        if "summary_stats" in data and "topper_df" in data:
            return "result"
        if "workload_df" in data:
            return "timetable"
        if "kpis" in data:
            return "financial"
        return ""

    def _set_slide_bg(self, slide, color_hex: str) -> None:
        """Fill slide background with a solid hex color."""
        try:
            bg = slide.background
            fill = bg.fill
            fill.solid()
            r, g, b = self._hex_to_rgb(color_hex)
            fill.fore_color.rgb = RGBColor(r, g, b)
        except Exception as e:
            logger.warning("Could not set slide background: %s", e)

    def _add_rect(
        self,
        slide,
        left_in: float,
        top_in: float,
        width_in: float,
        height_in: float,
        fill_hex: str,
        line_hex: Optional[str] = None,
    ) -> None:
        """Add a solid filled rectangle (integer shape type 1 = RECTANGLE)."""
        try:
            shape = slide.shapes.add_shape(
                1,                      # 1 = MSO_AUTO_SHAPE_TYPE.RECTANGLE
                Inches(left_in),
                Inches(top_in),
                Inches(width_in),
                Inches(height_in),
            )
            r, g, b = self._hex_to_rgb(fill_hex)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(r, g, b)
            if line_hex:
                lr, lg, lb = self._hex_to_rgb(line_hex)
                shape.line.color.rgb = RGBColor(lr, lg, lb)
            else:
                # Hide border by matching fill colour
                shape.line.color.rgb = RGBColor(r, g, b)
        except Exception as e:
            logger.debug("Could not add rectangle shape: %s", e)

    def _add_text_box(
        self,
        slide,
        text: str,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
        font_size: int = 12,
        bold: bool = False,
        italic: bool = False,
        color_hex: str = "1A1A1A",
        align: PP_ALIGN = PP_ALIGN.LEFT,
        font_name: str = "Calibri",
    ):
        """Add a text box with a single consistent format.

        Multi-line text (containing \\n) is split into separate paragraphs.
        """
        if text is None:
            text = ""
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        r, g, b = self._hex_to_rgb(color_hex)

        lines = str(text).split("\n")
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if line.strip():
                run = p.add_run()
                run.text = line
                run.font.size = Pt(font_size)
                run.font.bold = bold
                run.font.italic = italic
                run.font.name = font_name
                run.font.color.rgb = RGBColor(r, g, b)
        return txBox

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
        """Convert a 6-digit hex string (with or without #) to (R, G, B)."""
        h = str(hex_color).lstrip("#").strip()
        if len(h) != 6:
            return 0, 0, 0
        try:
            return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        except ValueError:
            return 0, 0, 0
