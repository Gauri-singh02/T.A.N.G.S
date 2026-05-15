"""
academic_module.py
T.A.N.G.S - Academic Result Analytics Module

Parses SAKEC B.Tech result PDFs (format: per-subject ESE/OR/GR/CR/GP),
computes toppers, KT students, failure rates, summary statistics, and
answers natural-language queries via the Groq LLM API.

Targets:
  - B_TECH_CYSE_MAY_2025_SEM-II.pdf  (63 students across 3 pages)
  - Generalizes to other SAKEC result PDFs with the same column scheme.

Pass mark for ESE = 40 (SAKEC convention).

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import re
import math
import logging
from typing import Dict, List, Optional, Any

import pandas as pd


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# SAKEC convention: ESE pass mark.
PASS_MARK = 40

# Columns that are NOT subject mark columns.
_RESERVED_COLS = {
    "roll_no", "name", "Result", "result", "kt_count",
    "failed_subjects", "total", "Total",
}


# ======================================================================
# PARSING
# ======================================================================
def parse_student_data(
    tables: List[pd.DataFrame], raw_text: str,
) -> pd.DataFrame:
    """Parse SAKEC result tables/text into a clean student DataFrame.

    Output schema:
        roll_no (str), name (str),
        <Subject_1> (float, ESE mark), ..., <Subject_N> (float, ESE mark),
        Result (str: 'PASS' / 'A.T.KT'),
        kt_count (int),
        total (float, sum of ESE marks).
    """
    df_from_tables: Optional[pd.DataFrame] = None
    if tables:
        try:
            df_from_tables = _parse_from_tables(tables, raw_text)
        except Exception as e:
            logger.warning("Table-based parse failed: %s", e)
            df_from_tables = None

    if df_from_tables is not None and not df_from_tables.empty:
        return _finalize_df(df_from_tables)

    # Fallback to raw text
    try:
        df_from_text = _parse_from_text(raw_text)
    except Exception as e:
        logger.warning("Text-based parse failed: %s", e)
        df_from_text = pd.DataFrame()

    return _finalize_df(df_from_text)


def _finalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure required columns exist, coerce subject cols to numeric,
    compute 'total' if missing."""
    if df is None or df.empty:
        return pd.DataFrame(
            columns=["roll_no", "name", "Result", "kt_count", "total"]
        )
    if "roll_no" not in df.columns:
        df["roll_no"] = ""
    if "name" not in df.columns:
        df["name"] = ""
    if "Result" not in df.columns:
        df["Result"] = ""
    if "kt_count" not in df.columns:
        df["kt_count"] = 0

    subject_cols = _subject_columns(df)
    for c in subject_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    if "total" not in df.columns:
        if subject_cols:
            df["total"] = df[subject_cols].sum(axis=1, skipna=True)
        else:
            df["total"] = 0.0
    else:
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0.0)

    # kt_count integer
    df["kt_count"] = pd.to_numeric(df["kt_count"], errors="coerce").fillna(0).astype(int)
    # Clean roll_no & name as strings
    df["roll_no"] = df["roll_no"].astype(str).str.strip()
    df["name"] = df["name"].astype(str).map(_clean_name)
    # Normalize Result
    df["Result"] = df["Result"].astype(str).str.upper().str.strip()
    df["Result"] = df["Result"].replace({"ATKT": "A.T.KT", "FAIL": "A.T.KT"})

    return df.reset_index(drop=True)


def _subject_columns(df: pd.DataFrame) -> List[str]:
    """Return likely subject (mark) columns: not reserved, numeric-coercible."""
    cols = []
    for c in df.columns:
        if c in _RESERVED_COLS:
            continue
        # Treat as a subject column if any value is numeric or coercible
        try:
            coerced = pd.to_numeric(df[c], errors="coerce")
            if coerced.notna().any():
                cols.append(c)
        except Exception:
            continue
    return cols


def _parse_from_tables(
    tables: List[pd.DataFrame], raw_text: str,
) -> Optional[pd.DataFrame]:
    """Inspect tables and extract student records."""
    scored = []
    for t in tables:
        if t is None or t.empty:
            continue
        sc = _score_table_as_result(t)
        if sc > 0:
            scored.append((sc, t))

    if not scored:
        return None

    all_records: List[Dict[str, Any]] = []
    subjects_seen: List[str] = []

    for _, t in sorted(scored, key=lambda x: -x[0]):
        recs, subs = _extract_records_from_table(t, raw_text)
        # Track first non-empty subject list
        if subs and not subjects_seen:
            subjects_seen = list(subs)
        all_records.extend(recs)

    if not all_records:
        return None

    df = pd.DataFrame(all_records)

    # Reorder columns: roll_no, name, subjects (in discovered order),
    # Result, kt_count, total, then anything else
    ordered: List[str] = ["roll_no", "name"]
    for s in subjects_seen:
        if s in df.columns and s not in ordered:
            ordered.append(s)
    for c in df.columns:
        if c not in ordered and c not in {"Result", "kt_count", "total"}:
            ordered.append(c)
    for tail in ("Result", "kt_count", "total"):
        if tail in df.columns:
            ordered.append(tail)
    df = df[[c for c in ordered if c in df.columns]]

    # Dedupe by roll_no across multi-page tables
    if "roll_no" in df.columns:
        df = df.drop_duplicates(subset=["roll_no"], keep="first").reset_index(drop=True)
    return df


def _score_table_as_result(df: pd.DataFrame) -> int:
    """Heuristic: how likely this table is a SAKEC result table."""
    if df is None or df.empty:
        return 0
    score = 0
    roll_re = re.compile(r"^\d{8,12}$")
    try:
        first_col = df.iloc[:, 0].astype(str).tolist()
    except Exception:
        first_col = []
    has_rolls = sum(1 for v in first_col if roll_re.match(v.strip()))
    if has_rolls > 0:
        score += 10 + has_rolls
    try:
        flat = " ".join(df.astype(str).values.flatten()).upper()
    except Exception:
        flat = ""
    if "ESE" in flat:
        score += 5
    if "PASS" in flat or "A.T.KT" in flat or "ATKT" in flat:
        score += 5
    if any(k in flat for k in ("ROLL", "SEAT", "NAME")):
        score += 2
    return score


def _extract_records_from_table(
    df: pd.DataFrame, raw_text: str,
):
    """Extract (records, subjects) from a single result-like table."""
    records: List[Dict[str, Any]] = []
    subjects = _discover_subjects(df, raw_text)

    # Try direct column-based extraction first if columns are clearly named.
    ese_cols = [c for c in df.columns if "ESE" in str(c).upper()]
    roll_col = _find_column_loose(df, ["roll", "seat"])
    name_col = _find_column_loose(df, ["name"])
    result_col = _find_column_loose(df, ["result"])
    kt_col = _find_column_loose(df, ["kt"])

    if ese_cols and roll_col is not None:
        # Map ESE columns to subjects (in column order)
        if subjects and len(subjects) == len(ese_cols):
            subj_to_col = dict(zip(subjects, ese_cols))
        elif subjects and len(subjects) > 0:
            # Best-effort 1-1 in declaration order, truncate to min len.
            n = min(len(subjects), len(ese_cols))
            subj_to_col = dict(zip(subjects[:n], ese_cols[:n]))
            # Fill remaining ESE cols with generic names
            for i, col in enumerate(ese_cols[n:], start=n + 1):
                subj_to_col[f"Subject_{i}"] = col
            subjects = list(subj_to_col.keys())
        else:
            subj_to_col = {
                f"Subject_{i + 1}": col for i, col in enumerate(ese_cols)
            }
            subjects = list(subj_to_col.keys())

        for _, row in df.iterrows():
            roll = _cell(row, roll_col)
            if not re.match(r"^\d{8,12}$", roll):
                continue
            rec: Dict[str, Any] = {
                "roll_no": roll,
                "name": _clean_name(_cell(row, name_col)) if name_col else "",
                "Result": _cell(row, result_col).upper() if result_col else "",
            }
            if kt_col is not None:
                try:
                    rec["kt_count"] = int(float(_cell(row, kt_col)))
                except (ValueError, TypeError):
                    rec["kt_count"] = 0
            else:
                rec["kt_count"] = 0
            for subj, col in subj_to_col.items():
                try:
                    rec[subj] = float(_cell(row, col))
                except (ValueError, TypeError):
                    rec[subj] = None
            records.append(rec)

    # If column-based produced no records, fall back to row-positional heuristic.
    if not records:
        records = _row_heuristic_extract(df, subjects)

    return records, subjects


def _discover_subjects(df: pd.DataFrame, raw_text: str) -> List[str]:
    """Discover subject codes from headers or raw text. Returns a list
    like ['BTCYS401', 'BTCYS402', ...] preserving order."""
    subjects: List[str] = []
    code_re = re.compile(r"\b[A-Z]{2,6}\d{2,6}\b")

    # From column names
    for c in df.columns:
        c_str = str(c)
        m = code_re.search(c_str)
        if m and m.group(0) not in subjects:
            subjects.append(m.group(0))

    # From first row (sometimes the merged header lives there)
    if not subjects and not df.empty:
        try:
            header_row = " ".join(df.iloc[0].astype(str).tolist())
        except Exception:
            header_row = ""
        for m in code_re.findall(header_row):
            if m not in subjects:
                subjects.append(m)

    # From raw_text near top
    if not subjects and raw_text:
        for m in code_re.findall(raw_text[:5000]):
            if m not in subjects and not m.startswith("SAKEC"):
                subjects.append(m)

    # Cap to a sensible upper bound (semester typically <= 8 subjects)
    return subjects[:8]


def _find_column_loose(
    df: pd.DataFrame, keywords: List[str],
) -> Optional[str]:
    for c in df.columns:
        c_str = str(c).lower()
        for kw in keywords:
            if kw in c_str:
                return c
    return None


def _cell(row: pd.Series, col: Optional[str]) -> str:
    if col is None:
        return ""
    val = row.get(col)
    if pd.isna(val):
        return ""
    return str(val).strip()


def _row_heuristic_extract(
    df: pd.DataFrame, subjects: List[str],
) -> List[Dict[str, Any]]:
    """Positional fallback: per row, find roll, name, then assume the
    numeric values come in groups of 4 per subject (ESE, OR, CR, GP --
    GR is a letter and is skipped during numeric parsing)."""
    records: List[Dict[str, Any]] = []
    n_subjects = len(subjects)

    for _, row in df.iterrows():
        values: List[str] = []
        for v in row.values:
            if pd.isna(v):
                values.append("")
            else:
                values.append(str(v).strip())

        # Find roll
        roll = None
        roll_idx = -1
        for i, v in enumerate(values):
            if re.match(r"^\d{8,12}$", v):
                roll = v
                roll_idx = i
                break
        if not roll:
            continue

        # Find name (next alpha-heavy non-result cell)
        name = ""
        name_idx = roll_idx
        for j in range(roll_idx + 1, len(values)):
            v = values[j]
            if not v:
                continue
            if re.match(r"^(PASS|FAIL|ATKT|A\.T\.KT)$", v, re.IGNORECASE):
                break
            if re.search(r"[A-Za-z]{2,}", v) and not re.search(r"\d", v):
                name = v
                name_idx = j
                break

        # Collect numeric & result tokens after name
        numbers: List[float] = []
        result = ""
        for k in range(name_idx + 1, len(values)):
            v = values[k]
            if not v:
                continue
            if re.match(r"^(PASS|FAIL|ATKT|A\.T\.KT)$", v, re.IGNORECASE):
                result = v.upper().replace("ATKT", "A.T.KT")
                if result == "FAIL":
                    result = "A.T.KT"
                continue
            try:
                numbers.append(float(v))
            except ValueError:
                # likely a grade letter; ignore
                continue

        rec: Dict[str, Any] = {
            "roll_no": roll,
            "name": _clean_name(name),
            "Result": result,
        }

        # Map ESE (every 4th number) to subjects
        if n_subjects > 0:
            for i, subj in enumerate(subjects):
                ese_idx = i * 4
                rec[subj] = numbers[ese_idx] if ese_idx < len(numbers) else None
            # KT count: first small-int tail number
            after = n_subjects * 4
            kt = 0
            for x in numbers[after:]:
                if 0 <= x <= 10 and x == int(x):
                    kt = int(x)
                    break
            rec["kt_count"] = kt
        else:
            rec["kt_count"] = 0

        records.append(rec)

    return records


def _parse_from_text(raw_text: str) -> pd.DataFrame:
    """Plain-regex fallback: parse student rows directly from raw text."""
    if not raw_text:
        return pd.DataFrame()

    # Subject discovery
    code_re = re.compile(r"\b[A-Z]{2,6}\d{2,6}\b")
    subjects: List[str] = []
    for c in code_re.findall(raw_text[:5000]):
        if c not in subjects and not c.startswith("SAKEC"):
            subjects.append(c)
    subjects = subjects[:8]
    n_subjects = len(subjects)

    records: List[Dict[str, Any]] = []
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(\d{8,12})\s+(.+)$", line)
        if not m:
            continue
        roll = m.group(1)
        rest = m.group(2)

        pf = re.search(
            r"\b(PASS|A\.T\.KT|ATKT|FAIL)\b", rest, re.IGNORECASE,
        )
        if pf:
            result = pf.group(1).upper().replace("ATKT", "A.T.KT")
            if result == "FAIL":
                result = "A.T.KT"
            before = rest[:pf.start()].strip()
        else:
            result = ""
            before = rest

        # Name: leading alphabetic run before first digit
        name_m = re.match(r"^([A-Z][A-Z\s\.\-]+?)(?=\s+\d|\s+[A-Z]+\d|\s*$)", before)
        name = _clean_name(name_m.group(1)) if name_m else ""

        numbers: List[float] = []
        for tok in re.findall(r"\d+\.\d+|\d+", rest):
            try:
                numbers.append(float(tok))
            except ValueError:
                continue

        rec: Dict[str, Any] = {
            "roll_no": roll,
            "name": name,
            "Result": result,
        }
        if n_subjects > 0:
            for i, subj in enumerate(subjects):
                ese_idx = i * 4
                rec[subj] = numbers[ese_idx] if ese_idx < len(numbers) else None
            after = n_subjects * 4
            kt = 0
            for x in numbers[after:]:
                if 0 <= x <= 10 and x == int(x):
                    kt = int(x)
                    break
            rec["kt_count"] = kt
        else:
            rec["kt_count"] = 0
        records.append(rec)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    if "roll_no" in df.columns:
        df = df.drop_duplicates(subset=["roll_no"], keep="first").reset_index(drop=True)
    return df


def _clean_name(name: Any) -> str:
    """Normalize a student name: collapse whitespace, strip junk chars."""
    if name is None:
        return ""
    if isinstance(name, float) and math.isnan(name):
        return ""
    s = str(name)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # Drop trailing punctuation noise
    s = s.strip(" ,.;:-")
    return s


# ======================================================================
# ANALYTICS
# ======================================================================
def get_toppers(
    df: pd.DataFrame,
    subject: Optional[str] = None,
    top_n: int = 10,
) -> pd.DataFrame:
    """Return the top-N students. If subject is None, sort by total
    marks descending. Otherwise sort by that subject's ESE marks."""
    if df is None or df.empty:
        return pd.DataFrame()

    work = df.copy()
    subject_cols = _subject_columns(work)

    if subject is None:
        if "total" in work.columns:
            sort_col = "total"
        elif "Total" in work.columns:
            sort_col = "Total"
        else:
            work["total"] = work[subject_cols].sum(axis=1, skipna=True) if subject_cols else 0
            sort_col = "total"
        sorted_df = (
            work.sort_values(sort_col, ascending=False, na_position="last")
            .head(top_n)
            .reset_index(drop=True)
        )
    else:
        matched_col = None
        for c in subject_cols:
            if str(c).lower() == subject.lower() or subject.lower() in str(c).lower():
                matched_col = c
                break
        if matched_col is None:
            logger.warning(
                "Subject %r not found among %s", subject, subject_cols,
            )
            return pd.DataFrame()
        sorted_df = (
            work.sort_values(matched_col, ascending=False, na_position="last")
            .head(top_n)
            .reset_index(drop=True)
        )

    sorted_df.insert(0, "rank", range(1, len(sorted_df) + 1))
    return sorted_df


def get_kt_students(
    df: pd.DataFrame, pass_mark: int = PASS_MARK,
) -> pd.DataFrame:
    """Return students who have any subject ESE < pass_mark or whose
    Result is A.T.KT. Adds 'failed_subjects' and 'kt_count' columns."""
    if df is None or df.empty:
        return pd.DataFrame()

    subject_cols = _subject_columns(df)
    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        failed: List[str] = []
        for c in subject_cols:
            val = row.get(c)
            if pd.notna(val):
                try:
                    if float(val) < pass_mark:
                        failed.append(c)
                except (ValueError, TypeError):
                    continue
        result_val = str(row.get("Result", "")).upper().strip()
        is_atkt = result_val in {"A.T.KT", "ATKT", "FAIL"}
        if failed or is_atkt:
            new_row = row.to_dict()
            new_row["failed_subjects"] = failed
            new_row["kt_count"] = len(failed) if failed else max(
                int(new_row.get("kt_count", 0) or 0), 1 if is_atkt else 0,
            )
            rows.append(new_row)

    if not rows:
        return pd.DataFrame()

    result_df = pd.DataFrame(rows)
    if "kt_count" in result_df.columns:
        result_df = result_df.sort_values(
            "kt_count", ascending=False,
        ).reset_index(drop=True)
    return result_df


def failure_rate_per_subject(
    df: pd.DataFrame, pass_mark: int = PASS_MARK,
) -> Dict[str, Dict[str, float]]:
    """For each subject column return {failed, rate, passed}."""
    if df is None or df.empty:
        return {}

    out: Dict[str, Dict[str, float]] = {}
    subject_cols = _subject_columns(df)
    for c in subject_cols:
        vals = pd.to_numeric(df[c], errors="coerce").dropna()
        if len(vals) == 0:
            continue
        failed = int((vals < pass_mark).sum())
        passed = int((vals >= pass_mark).sum())
        total = len(vals)
        rate = (failed / total) * 100.0 if total > 0 else 0.0
        out[str(c)] = {
            "failed": failed,
            "rate": round(rate, 2),
            "passed": passed,
        }
    return out


def get_summary_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """Return overall + per-subject summary statistics."""
    if df is None or df.empty:
        return {
            "total_students": 0,
            "passed": 0,
            "failed": 0,
            "atkt": 0,
            "per_subject": {},
            "overall_pass_percentage": 0.0,
        }

    total = len(df)
    subject_cols = _subject_columns(df)

    if "Result" in df.columns:
        res_series = df["Result"].astype(str).str.upper().str.strip()
        passed = int((res_series == "PASS").sum())
        atkt = int(res_series.isin(["A.T.KT", "ATKT", "FAIL"]).sum())
    else:
        passed = 0
        atkt = 0
        for _, row in df.iterrows():
            failed_any = False
            for c in subject_cols:
                v = row.get(c)
                if pd.notna(v):
                    try:
                        if float(v) < PASS_MARK:
                            failed_any = True
                            break
                    except (ValueError, TypeError):
                        continue
            if failed_any:
                atkt += 1
            else:
                passed += 1

    failed = atkt

    per_subject: Dict[str, Dict[str, float]] = {}
    for c in subject_cols:
        vals = pd.to_numeric(df[c], errors="coerce").dropna()
        if len(vals) == 0:
            continue
        per_subject[str(c)] = {
            "max": float(vals.max()),
            "min": float(vals.min()),
            "average": round(float(vals.mean()), 2),
        }

    pass_pct = (passed / total * 100.0) if total > 0 else 0.0

    return {
        "total_students": total,
        "passed": passed,
        "failed": failed,
        "atkt": atkt,
        "per_subject": per_subject,
        "overall_pass_percentage": round(pass_pct, 2),
    }


def aggregate_score(scores: List[Any]) -> float:
    """Return the arithmetic mean of the given scores. Handles None / NaN."""
    if not scores:
        return 0.0
    valid: List[float] = []
    for s in scores:
        if s is None:
            continue
        if isinstance(s, float) and math.isnan(s):
            continue
        try:
            valid.append(float(s))
        except (TypeError, ValueError):
            continue
    if not valid:
        return 0.0
    return sum(valid) / len(valid)


# ======================================================================
# LLM QUERY
# ======================================================================
def answer_query_with_llm(
    df: pd.DataFrame, query: str, groq_client,
) -> str:
    """Send a natural-language query about the result data to Groq's
    llama3-8b-8192 and return the answer text."""
    if df is None or df.empty:
        return "No student data available to answer the query."
    if not query or not str(query).strip():
        return "Empty query."

    # Build dataframe head context
    try:
        context_head = df.head(20).to_string(index=False)
    except Exception as e:
        context_head = f"(Could not render dataframe head: {e})"

    summary = get_summary_stats(df)
    # Compact per-subject string
    per_sub_lines = []
    for subj, stats in summary.get("per_subject", {}).items():
        per_sub_lines.append(
            f"  - {subj}: max={stats['max']}, min={stats['min']}, "
            f"avg={stats['average']}"
        )
    per_sub_str = "\n".join(per_sub_lines) if per_sub_lines else "  (none)"

    summary_str = (
        f"Total students: {summary['total_students']}\n"
        f"Passed: {summary['passed']}\n"
        f"Failed / A.T.KT: {summary['atkt']}\n"
        f"Overall pass percentage: {summary['overall_pass_percentage']}%\n"
        f"Per-subject stats:\n{per_sub_str}\n"
    )

    system_msg = (
        "You are an academic analyst for SAKEC Mumbai engineering college. "
        "Analyze student result data and answer queries about toppers, "
        "failures, KT students, and performance statistics. Be precise "
        "and concise."
    )
    user_msg = (
        f"Student data (first 20 rows):\n{context_head}\n\n"
        f"Summary statistics:\n{summary_str}\n"
        f"User query: {query}"
    )

    try:
        response = groq_client.chat.completions.create(
            model="llama3-8b-8192",
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            max_tokens=500,
        )
        # Defensive access
        if hasattr(response, "choices") and response.choices:
            choice = response.choices[0]
            if hasattr(choice, "message") and hasattr(choice.message, "content"):
                content = choice.message.content
                if content:
                    return str(content).strip()
        return "LLM returned empty response."
    except Exception as e:
        logger.error("Groq LLM call failed: %s", e)
        return f"LLM query failed: {e}"


# ======================================================================
# ORCHESTRATION
# ======================================================================
def process_result(
    extractor_output: Dict[str, Any],
    query: str,
    groq_client,
) -> Dict[str, Any]:
    """Orchestrate the full result-processing pipeline.

    Returns:
        {
            "dataframe":     pd.DataFrame,
            "query_answer":  str,
            "summary_stats": dict,
            "topper_df":     pd.DataFrame,
            "kt_df":         pd.DataFrame,
            "failure_rates": dict,
        }
    """
    tables = extractor_output.get("tables", []) if extractor_output else []
    raw_text = extractor_output.get("raw_text", "") if extractor_output else ""

    df = parse_student_data(tables, raw_text)

    summary = get_summary_stats(df)
    toppers = get_toppers(df, subject=None, top_n=10)
    kts = get_kt_students(df, pass_mark=PASS_MARK)
    fail_rates = failure_rate_per_subject(df, pass_mark=PASS_MARK)

    # Route the query through the LLM (with structured stats in the prompt).
    # We always defer to the LLM for the final natural-language answer so
    # the user sees a polished response; the structured outputs are also
    # returned for the caller to render directly.
    query_lower = (query or "").lower()
    answer_text: str

    if df.empty:
        answer_text = (
            "Could not parse any student records from the document. "
            "Please verify the PDF is a SAKEC-format result sheet."
        )
    else:
        # If the query targets a specific subject's toppers, augment context
        if any(k in query_lower for k in ("topper", "top ", "rank", "highest")):
            subj = _extract_subject_from_query(query, df)
            subject_toppers = get_toppers(df, subject=subj, top_n=10)
            # Override toppers with subject-specific version if found
            if subj and not subject_toppers.empty:
                toppers = subject_toppers
        answer_text = answer_query_with_llm(df, query, groq_client)

    return {
        "dataframe": df,
        "query_answer": answer_text,
        "summary_stats": summary,
        "topper_df": toppers,
        "kt_df": kts,
        "failure_rates": fail_rates,
    }


def _extract_subject_from_query(
    query: str, df: pd.DataFrame,
) -> Optional[str]:
    """Try to extract a subject name from the query by matching against
    DataFrame subject columns (case-insensitive substring match)."""
    if df is None or df.empty or not query:
        return None
    subject_cols = _subject_columns(df)
    qlow = query.lower()
    # Match longer column names first to avoid false partial hits
    for c in sorted(subject_cols, key=lambda x: -len(str(x))):
        if str(c).lower() in qlow:
            return c
    return None
