"""
pdf_extractor.py
T.A.N.G.S - Hybrid PDF Extraction Engine

Three-tier extraction strategy:
  Tier 1: pdfplumber table extraction (clean digital PDFs e.g. FE_time_table.pdf,
          B_TECH_CYSE_MAY_2025_SEM-II.pdf)
  Tier 2: pdfplumber raw text + regex (semi-structured fallback)
  Tier 3: pdf2image + pytesseract OCR with auto-rotation (scanned PDFs
          e.g. Btech2.pdf)

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import os
import re
import logging
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import pdfplumber
from fuzzywuzzy import fuzz, process

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# Canonical column names used for fuzzy normalization of extracted headers.
CANONICAL_COLUMNS = [
    # Result columns
    "Roll No", "Roll Number", "Seat No", "Name", "Student Name",
    "ESE", "OR", "GR", "CR", "GP", "Total", "Result",
    "KT", "KT Count", "ATKT",
    # Timetable columns
    "Period", "Time", "Day", "Slot",
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
    "Subject", "Subject Code", "Faculty", "Teacher", "Room",
    # Financial columns
    "Revenue", "Profit", "Earnings", "Quarter", "Year",
    "Balance", "Period", "Amount",
]


class HybridExtractor:
    """Three-tier hybrid PDF extractor for academic, timetable, and
    financial PDFs. Designed for the SAKEC corpus but generalizes to
    similar layouts."""

    def __init__(self, fuzzy_threshold: int = 70, ocr_dpi: int = 300):
        self.fuzzy_threshold = fuzzy_threshold
        self.ocr_dpi = ocr_dpi

    # ------------------------------------------------------------------
    # PUBLIC API
    # ------------------------------------------------------------------
    def extract(self, file_path: str) -> Dict[str, Any]:
        """Run the three-tier extraction pipeline.

        Returns a dict with keys:
            tables      - list[pd.DataFrame]
            raw_text    - str
            method_used - "tier1_pdfplumber" | "tier2_regex" | "tier3_ocr" | "failed"
            doc_type    - "result" | "timetable" | "financial"
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"PDF not found: {file_path}")

        result: Dict[str, Any] = {
            "tables": [],
            "raw_text": "",
            "method_used": "none",
            "doc_type": "unknown",
        }

        # Always get text first (cheap, useful for legend parsing
        # and doc-type detection even if higher tier wins).
        raw_text = self._extract_text_pdfplumber(file_path)
        result["raw_text"] = raw_text

        # ----- TIER 1 -----
        try:
            tables = self._tier1_pdfplumber(file_path)
            total_rows = sum(len(df) for df in tables)
            if total_rows >= 2:
                tables = [self._postprocess_df(df) for df in tables]
                tables = [self.normalize_teacher_codes(df, raw_text) for df in tables]
                result["tables"] = tables
                result["method_used"] = "tier1_pdfplumber"
                result["doc_type"] = self.detect_doc_type(raw_text)
                logger.info(
                    "Tier 1 succeeded: %d tables, %d total rows",
                    len(tables), total_rows,
                )
                return result
            logger.info(
                "Tier 1 yielded only %d rows; falling through to Tier 2",
                total_rows,
            )
        except Exception as e:
            logger.warning("Tier 1 failed: %s", e)

        # ----- TIER 2 -----
        try:
            regex_df, matched_fields = self._tier2_regex(raw_text)
            if matched_fields >= 3:
                # Try to normalize teacher codes in the regex df as well
                regex_df = self.normalize_teacher_codes(regex_df, raw_text)
                result["tables"] = [regex_df] if not regex_df.empty else []
                result["method_used"] = "tier2_regex"
                result["doc_type"] = self.detect_doc_type(raw_text)
                logger.info(
                    "Tier 2 succeeded with %d matched field types", matched_fields,
                )
                return result
            logger.info(
                "Tier 2 matched only %d field types; falling through to Tier 3",
                matched_fields,
            )
        except Exception as e:
            logger.warning("Tier 2 failed: %s", e)

        # ----- TIER 3 -----
        try:
            ocr_text, ocr_tables = self._tier3_ocr(file_path)
            final_text = ocr_text if ocr_text else raw_text
            result["raw_text"] = final_text
            # Try to normalize teacher codes against the OCR text
            ocr_tables = [
                self.normalize_teacher_codes(t, final_text) for t in ocr_tables
            ]
            result["tables"] = ocr_tables
            result["method_used"] = "tier3_ocr"
            result["doc_type"] = self.detect_doc_type(final_text)
            logger.info(
                "Tier 3 OCR succeeded: %d tables, %d chars text",
                len(ocr_tables), len(final_text),
            )
            return result
        except Exception as e:
            logger.error("Tier 3 OCR failed: %s", e)

        # All tiers failed
        result["method_used"] = "failed"
        result["doc_type"] = (
            self.detect_doc_type(raw_text) if raw_text else "unknown"
        )
        return result

    # ------------------------------------------------------------------
    # TIER 1: pdfplumber tables
    # ------------------------------------------------------------------
    def _tier1_pdfplumber(self, file_path: str) -> List[pd.DataFrame]:
        """Extract all tables from all pages using pdfplumber."""
        all_tables: List[pd.DataFrame] = []
        with pdfplumber.open(file_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                try:
                    page_tables = page.extract_tables()
                except Exception as e:
                    logger.warning(
                        "Page %d table extract failed: %s", page_idx, e,
                    )
                    page_tables = []
                for tbl in page_tables:
                    if not tbl or len(tbl) < 1:
                        continue
                    header_raw = tbl[0]
                    body = tbl[1:] if len(tbl) > 1 else []
                    header = [self._clean_cell(c) for c in header_raw]
                    header = self._make_unique_columns(header)
                    if body:
                        # Pad/truncate body rows to header width
                        n = len(header)
                        norm_body = []
                        for r in body:
                            if len(r) < n:
                                norm_body.append(list(r) + [""] * (n - len(r)))
                            elif len(r) > n:
                                norm_body.append(list(r)[:n])
                            else:
                                norm_body.append(list(r))
                        df = pd.DataFrame(norm_body, columns=header)
                    else:
                        df = pd.DataFrame(columns=header)
                    if not df.empty:
                        all_tables.append(df)
        return all_tables

    def _postprocess_df(self, df: pd.DataFrame) -> pd.DataFrame:
        """Strip whitespace, forward-fill merged cells, fuzzy-normalize
        column names."""
        # Replace empty strings with NA
        df = df.replace(r"^\s*$", pd.NA, regex=True)
        # Forward-fill (vertical) to handle merged cells in the same column
        for col in df.columns:
            try:
                df[col] = df[col].ffill()
            except Exception:
                pass
        # Strip cell whitespace
        try:
            df = df.applymap(lambda v: v.strip() if isinstance(v, str) else v)
        except Exception:
            # pandas 2.1+ deprecates applymap; map works on DataFrame there
            try:
                df = df.map(lambda v: v.strip() if isinstance(v, str) else v)
            except Exception:
                pass
        # Normalize headers via fuzzy match against canonical set
        new_cols = []
        seen: Dict[str, int] = {}
        for col in df.columns:
            normalized = self._fuzzy_normalize_column(str(col))
            count = seen.get(normalized, 0)
            seen[normalized] = count + 1
            new_cols.append(
                normalized if count == 0 else f"{normalized}_{count}"
            )
        df.columns = new_cols
        return df

    def _fuzzy_normalize_column(self, col_name: str) -> str:
        if not col_name or not isinstance(col_name, str):
            return "col"
        cleaned = re.sub(r"\s+", " ", col_name).strip()
        if not cleaned:
            return "col"
        match = process.extractOne(cleaned, CANONICAL_COLUMNS, scorer=fuzz.ratio)
        if match and match[1] >= self.fuzzy_threshold:
            return match[0]
        return cleaned

    @staticmethod
    def _clean_cell(value: Any) -> str:
        if value is None:
            return ""
        s = str(value).replace("\n", " ").strip()
        s = re.sub(r"\s+", " ", s)
        return s

    @staticmethod
    def _make_unique_columns(cols: List[str]) -> List[str]:
        seen: Dict[str, int] = {}
        out: List[str] = []
        for i, c in enumerate(cols):
            base = c if c else f"col_{i}"
            n = seen.get(base, 0)
            seen[base] = n + 1
            out.append(base if n == 0 else f"{base}_{n}")
        return out

    # ------------------------------------------------------------------
    # TIER 2: regex on raw text
    # ------------------------------------------------------------------
    def _tier2_regex(self, raw_text: str) -> Tuple[pd.DataFrame, int]:
        if not raw_text:
            return pd.DataFrame(), 0

        roll_pattern = re.compile(r"[A-Z0-9]{10,20}")
        score_pattern = re.compile(r"\b\d{1,3}\.\d{1,2}\b|\b\d{1,3}\b")
        passfail_pattern = re.compile(
            r"\b(PASS|FAIL|A\.T\.KT|ATKT)\b", re.IGNORECASE,
        )
        subject_pattern = re.compile(r"\b[A-Z]{2,6}\d{3,6}\b")

        rows: List[Dict[str, str]] = []
        lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
        for ln in lines:
            roll_match = roll_pattern.search(ln)
            scores = score_pattern.findall(ln)
            pf_match = passfail_pattern.search(ln)
            subj_match = subject_pattern.findall(ln)

            if not roll_match:
                continue

            hits = sum([
                1 if roll_match else 0,
                1 if scores else 0,
                1 if pf_match else 0,
                1 if subj_match else 0,
            ])
            if hits >= 2:
                rows.append({
                    "roll_no": roll_match.group(0),
                    "scores": ",".join(scores),
                    "pass_fail": (
                        pf_match.group(0).upper() if pf_match else ""
                    ),
                    "subjects": ",".join(subj_match),
                    "raw_line": ln,
                })

        # Count distinct field types matched globally (the spec metric)
        global_hits = 0
        if roll_pattern.search(raw_text):
            global_hits += 1
        if score_pattern.search(raw_text):
            global_hits += 1
        if passfail_pattern.search(raw_text):
            global_hits += 1
        if subject_pattern.search(raw_text):
            global_hits += 1

        df = pd.DataFrame(rows) if rows else pd.DataFrame()
        return df, global_hits

    # ------------------------------------------------------------------
    # TIER 3: OCR for scanned PDFs
    # ------------------------------------------------------------------
    def _tier3_ocr(self, file_path: str) -> Tuple[str, List[pd.DataFrame]]:
        if not OCR_AVAILABLE or not PDF2IMAGE_AVAILABLE:
            raise RuntimeError(
                "pytesseract and pdf2image are required for OCR (Tier 3). "
                "Install with: pip install pytesseract pdf2image "
                "and ensure tesseract + poppler are on PATH."
            )
        images = convert_from_path(file_path, dpi=self.ocr_dpi)
        all_text_parts: List[str] = []
        all_tables: List[pd.DataFrame] = []
        for idx, img in enumerate(images):
            rotated = self._auto_rotate(img)
            try:
                text = pytesseract.image_to_string(
                    rotated, config="--oem 3 --psm 6",
                )
            except Exception as e:
                logger.warning("OCR page %d failed: %s", idx, e)
                text = ""
            all_text_parts.append(text)
            page_table = self._text_to_table_rows(text)
            if page_table is not None and not page_table.empty:
                all_tables.append(page_table)
        return "\n".join(all_text_parts), all_tables

    def _auto_rotate(self, img: "Image.Image") -> "Image.Image":
        """Use tesseract OSD (--psm 0) to detect rotation and correct it."""
        try:
            osd = pytesseract.image_to_osd(img, config="--psm 0")
            rot_match = re.search(r"Rotate:\s*(\d+)", osd)
            if rot_match:
                rotation = int(rot_match.group(1))
                if rotation % 360 != 0:
                    # PIL Image.rotate is CCW; tesseract Rotate value is CW
                    # needed to correct. So rotate by -rotation degrees.
                    return img.rotate(-rotation, expand=True)
        except Exception as e:
            logger.warning("OSD rotation detection failed: %s", e)
        return img

    def _text_to_table_rows(self, text: str) -> Optional[pd.DataFrame]:
        """Best-effort: split OCR text into table rows by multi-space gaps."""
        if not text:
            return None
        rows: List[List[str]] = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = re.split(r"\s{2,}|\t+|\|", line)
            parts = [p.strip() for p in parts if p.strip()]
            if len(parts) >= 3:
                rows.append(parts)
        if not rows:
            return None
        max_cols = max(len(r) for r in rows)
        padded = [r + [""] * (max_cols - len(r)) for r in rows]
        cols = [f"col_{i}" for i in range(max_cols)]
        return pd.DataFrame(padded, columns=cols)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _extract_text_pdfplumber(file_path: str) -> str:
        try:
            parts: List[str] = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    try:
                        t = page.extract_text() or ""
                        parts.append(t)
                    except Exception:
                        continue
            return "\n".join(parts)
        except Exception as e:
            logger.warning("Raw text extraction failed: %s", e)
            return ""

    def detect_doc_type(self, text: str) -> str:
        """Classify the document as result / timetable / financial.

        Default: 'result' if no signal.
        """
        if not text:
            return "result"
        t = text.lower()
        result_kw = [
            "result", "marks", "pass", "fail", "roll",
            "semester", "examination", "grade", "ese",
        ]
        timetable_kw = [
            "period", "time", "monday", "tuesday", "lecture",
            "slot", "faculty", "room",
        ]
        financial_kw = [
            "revenue", "profit", "balance", "earnings",
            "fiscal", "quarterly",
        ]

        res = sum(1 for k in result_kw if k in t)
        tt = sum(1 for k in timetable_kw if k in t)
        fin = sum(1 for k in financial_kw if k in t)

        scores = {"result": res, "timetable": tt, "financial": fin}
        best = max(scores, key=scores.get)
        if scores[best] == 0:
            return "result"
        return best

    def normalize_teacher_codes(
        self, df: pd.DataFrame, legend_text: str,
    ) -> pd.DataFrame:
        """Replace teacher codes like '/MK' in DataFrame cells with full
        names parsed from the SAKEC legend text (e.g. '/MK: Manjusha
        Kulkarni')."""
        if df is None or df.empty or not legend_text:
            return df

        mapping = self._parse_teacher_legend(legend_text)
        if not mapping:
            return df

        def replace_codes(value: Any) -> Any:
            if not isinstance(value, str):
                return value
            new_val = value
            # Replace longer codes first to avoid prefix collisions.
            for code in sorted(mapping.keys(), key=len, reverse=True):
                if code in new_val:
                    # Preserve the separator: '/MK' -> '/Manjusha Kulkarni'
                    new_val = new_val.replace(code, "/" + mapping[code])
            return new_val

        out_df = df.copy()
        for col in out_df.columns:
            try:
                out_df[col] = out_df[col].apply(replace_codes)
            except Exception:
                continue
        return out_df

    @staticmethod
    def _parse_teacher_legend(text: str) -> Dict[str, str]:
        """Parse '/XX: Full Name' and '/XX - Full Name' pairs from text."""
        mapping: Dict[str, str] = {}
        # Match "/CODE : Name" where name runs until next "/CODE" marker
        # or end of line. CODE is 1-5 uppercase letters.
        pattern = re.compile(
            r"/([A-Z]{1,5})\s*[:\-\u2013]\s*([^/\n\r]+)"
        )
        for m in pattern.finditer(text):
            code = m.group(1).strip()
            name = m.group(2).strip().rstrip(",;.")
            # Defensive: collapse internal whitespace
            name = re.sub(r"\s+", " ", name)
            if code and name and 2 <= len(name) <= 60:
                # Only accept names that look like a person (contain letters)
                if re.search(r"[A-Za-z]{2,}", name):
                    mapping["/" + code] = name
        return mapping
