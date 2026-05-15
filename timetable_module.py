"""
timetable_module.py
T.A.N.G.S - Timetable Parsing & Analytics Module

Parses SAKEC timetable PDFs (per-period rows, per-day columns) into a
long-format DataFrame, computes teacher workload, subject frequency,
free slots, and answers natural-language queries via Groq LLM.

Cell format (SAKEC convention):
    "SubjectCode /TeacherCode RoomNo"   e.g.  "DE&MI /MK 108"
    "Subject1 /T1 / Subject2 /T2"        e.g.  "DE /MK/SR 108" (team)
    "A: ML /SG 305 / B: CN /SC 306"      (batch split)

Targets:
  - FE_time_table.pdf   (digital, clean, FE B.Tech Sem-II)
  - Btech2.pdf          (scanned, rotated, BE Sem-VIII)

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import re
import logging
from typing import Dict, List, Optional, Any, Tuple

import pandas as pd
from fuzzywuzzy import fuzz


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# ---- Day name normalization ------------------------------------------
DAY_NORM: Dict[str, str] = {
    "mon": "Monday", "monday": "Monday",
    "tue": "Tuesday", "tues": "Tuesday", "tuesday": "Tuesday",
    "wed": "Wednesday", "weds": "Wednesday", "wednesday": "Wednesday",
    "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday",
    "thursday": "Thursday",
    "fri": "Friday", "friday": "Friday",
    "sat": "Saturday", "saturday": "Saturday",
}

ORDERED_DAYS: List[str] = [
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
]

# SAKEC branch keywords
_BRANCH_KEYWORDS: List[Tuple[str, List[str]]] = [
    ("Computer", ["computer engineering", "comp engg", "cse"]),
    ("IT", ["information technology", "info tech"]),
    ("AIDS", ["artificial intelligence and data science",
              "ai & ds", "aids"]),
    ("CYSE", ["cyber security", "cyse", "cy se"]),
    ("ENTC", ["electronics and telecommunication", "extc", "entc"]),
    ("MECH", ["mechanical engineering"]),
    ("FE", ["first year", "fy b.tech", "fy btech"]),
]


# ======================================================================
# LEGEND & BRANCH HELPERS
# ======================================================================
def _parse_legend(raw_text: str) -> Dict[str, str]:
    """Parse '/XX: Full Name' pairs from raw text using the spec pattern."""
    if not raw_text:
        return {}
    pattern = re.compile(
        r"/([A-Z]+):\s*([A-Za-z\s\.]+?)(?=\s*/[A-Z]+:|$)"
    )
    mapping: Dict[str, str] = {}
    for m in pattern.finditer(raw_text):
        code = m.group(1).strip()
        name = m.group(2).strip().rstrip(",;.")
        name = re.sub(r"\s+", " ", name)
        if code and name and 2 <= len(name) <= 60:
            if re.search(r"[A-Za-z]{2,}", name):
                mapping["/" + code] = name
    return mapping


def _detect_branch(raw_text: str) -> str:
    """Heuristic: detect SAKEC branch / programme from raw text."""
    if not raw_text:
        return ""
    t = raw_text.lower()
    for code, terms in _BRANCH_KEYWORDS:
        for term in terms:
            if term in t:
                return code
    sem_m = re.search(
        r"sem(?:ester)?\s*[-]?\s*([IVX]+|\d+)",
        raw_text, re.IGNORECASE,
    )
    if sem_m:
        return f"Sem-{sem_m.group(1).strip()}"
    return ""


def _resolve_teacher(teacher_code: str, legend: Dict[str, str]) -> str:
    """Resolve '/MK' (and team '/MK/SR') to full names using legend."""
    if not teacher_code:
        return ""
    parts = re.findall(r"/[A-Z]+", teacher_code)
    if not parts:
        return teacher_code
    names: List[str] = []
    for p in parts:
        names.append(legend.get(p, p))
    return ", ".join(names)


# ======================================================================
# TABLE & COLUMN DISCOVERY
# ======================================================================
def _normalize_day(s: Any) -> Optional[str]:
    if s is None:
        return None
    text = str(s).lower().strip()
    text = re.sub(r"[^a-z]", "", text)
    if not text:
        return None
    return DAY_NORM.get(text)


def _find_day_columns(df: pd.DataFrame) -> Dict[Any, str]:
    """Return dict mapping df column -> canonical day name."""
    out: Dict[Any, str] = {}
    for c in df.columns:
        norm = _normalize_day(c)
        if norm and norm not in out.values():
            out[c] = norm
    return out


def _find_period_column(df: pd.DataFrame) -> Optional[Any]:
    """Find the period-number column."""
    for c in df.columns:
        c_str = str(c).lower()
        if "period" in c_str or "slot" in c_str:
            return c
        if c_str.strip() in {"no", "no.", "sr", "sr.", "sr no", "s.no"}:
            return c
    # Fallback: first column with mostly small numeric values 1-12
    for c in df.columns:
        try:
            vals = pd.to_numeric(df[c], errors="coerce").dropna()
            if len(vals) >= 3 and vals.max() <= 12 and vals.min() >= 1:
                return c
        except Exception:
            continue
    return None


def _find_time_column(df: pd.DataFrame) -> Optional[Any]:
    """Find the time column (cells like '8:00-9:00')."""
    time_re = re.compile(r"\d{1,2}[:\.]\d{2}")
    for c in df.columns:
        c_str = str(c).lower()
        if "time" in c_str:
            return c
    for c in df.columns:
        try:
            hits = df[c].astype(str).str.contains(time_re).sum()
            if hits >= 3:
                return c
        except Exception:
            continue
    return None


def _find_timetable_tables(
    tables: List[pd.DataFrame],
) -> List[pd.DataFrame]:
    """Return tables whose columns contain >=3 day names."""
    results: List[pd.DataFrame] = []
    for t in tables:
        if t is None or t.empty:
            continue
        day_matches = sum(
            1 for c in t.columns if _normalize_day(c) is not None
        )
        if day_matches >= 3:
            results.append(t)
    return results


# ======================================================================
# CELL PARSING
# ======================================================================
_RECESS_RE = re.compile(r"\b(RECESS|LUNCH|BREAK)\b", re.IGNORECASE)
_BATCH_SPLIT_RE = re.compile(r"\s*\b([AB])\s*[:\)]\s*")
_CELL_FULL_RE = re.compile(
    r"^([A-Za-z0-9&\-\s]+?)\s*((?:/[A-Z]+)+)\s*([A-Za-z0-9\-]*)$"
)
_TEACHER_FALLBACK_RE = re.compile(r"(/[A-Z]+(?:/[A-Z]+)*)")


def _parse_cell(cell: Any) -> List[Tuple[str, str, str]]:
    """Parse a timetable cell.

    Returns a list of (subject, teacher_code, room) tuples.
    Returns [] for RECESS / LUNCH / BREAK / empty cells.
    Returns multiple tuples for split batches (A:.../B:...) or where
    several entries are stacked in the same cell.
    """
    if cell is None:
        return []
    s = str(cell).strip()
    if not s:
        return []
    if _RECESS_RE.search(s):
        return []

    # Flatten newlines for primary parse but keep original for newline-split
    s_flat = re.sub(r"[\n\r]+", " ", s)
    s_flat = re.sub(r"\s+", " ", s_flat).strip()

    parts: List[str] = []

    # Split by batch markers if present (A: ... B: ...)
    if re.search(r"\b[AB]\s*[:\)]\s*\w", s_flat):
        chunks = _BATCH_SPLIT_RE.split(s_flat)
        # chunks alternates: [pre, "A", content1, "B", content2, ...]
        i = 1
        while i < len(chunks):
            if i + 1 < len(chunks):
                piece = chunks[i + 1].strip()
                if piece:
                    parts.append(piece)
            i += 2

    # If no batch parts found, try newline-based split
    if not parts:
        nl_parts = [p.strip() for p in s.split("\n") if p.strip()]
        if (len(nl_parts) > 1
                and sum(1 for p in nl_parts if re.search(r"/[A-Z]+", p)) >= 2):
            parts = nl_parts
        else:
            parts = [s_flat]

    results: List[Tuple[str, str, str]] = []
    for raw in parts:
        p = raw.strip()
        if not p:
            continue
        if _RECESS_RE.search(p):
            continue
        m = _CELL_FULL_RE.match(p)
        if m:
            subject = re.sub(r"\s+", " ", m.group(1)).strip()
            teacher = m.group(2).strip()
            room = m.group(3).strip()
            if subject:
                results.append((subject, teacher, room))
            continue
        # Loose fallback: locate /XX anywhere
        tm = _TEACHER_FALLBACK_RE.search(p)
        if tm:
            teacher = tm.group(1)
            before = p[:tm.start()].strip()
            after = p[tm.end():].strip()
            subject = re.sub(r"\s+", " ", before).strip()
            room_m = re.search(r"([A-Za-z0-9\-]+)", after)
            room = room_m.group(1) if room_m else ""
            if subject:
                results.append((subject, teacher, room))

    return results


# ======================================================================
# RECORD EXTRACTION
# ======================================================================
def _extract_period(row: pd.Series, period_col: Optional[Any]) -> Optional[int]:
    if period_col is None:
        return None
    try:
        val = row.get(period_col)
        if pd.isna(val):
            return None
        m = re.search(r"\d+", str(val))
        if m:
            return int(m.group(0))
    except Exception:
        return None
    return None


def _extract_time(row: pd.Series, time_col: Optional[Any]) -> str:
    if time_col is None:
        return ""
    try:
        val = row.get(time_col)
        if pd.isna(val):
            return ""
        return re.sub(r"\s+", " ", str(val)).strip()
    except Exception:
        return ""


def _parse_timetable_table(
    tt_df: pd.DataFrame,
    legend: Dict[str, str],
    branch: str,
) -> List[Dict[str, Any]]:
    """Extract records from a single timetable DataFrame."""
    records: List[Dict[str, Any]] = []

    work = tt_df.copy()
    # Replace empty strings with NA so ffill works
    work = work.replace(r"^\s*$", pd.NA, regex=True)
    # Forward-fill merged cells. limit=1 = handle the common 2-period lab
    # case without over-propagating into truly-free slots that follow.
    work = work.ffill(limit=1)

    period_col = _find_period_column(work)
    time_col = _find_time_column(work)
    day_cols = _find_day_columns(work)
    if not day_cols:
        return records

    for _, row in work.iterrows():
        period = _extract_period(row, period_col)
        time_val = _extract_time(row, time_col)
        for col_name, day_name in day_cols.items():
            cell = row.get(col_name)
            if cell is None or pd.isna(cell):
                continue
            entries = _parse_cell(cell)
            for subject, teacher_code, room in entries:
                teacher_full = _resolve_teacher(teacher_code, legend)
                records.append({
                    "teacher": teacher_full,
                    "subject": subject,
                    "day": day_name,
                    "slot": period if period is not None else None,
                    "time": time_val,
                    "room": room,
                    "branch": branch,
                })
    return records


def _parse_timetable_from_text(
    raw_text: str, legend: Dict[str, str], branch: str,
) -> List[Dict[str, Any]]:
    """Best-effort fallback: extract subject/teacher/room triples from
    raw text alone (used when tables are unusable, e.g. scanned PDFs)."""
    if not raw_text:
        return []
    pattern = re.compile(
        r"\b([A-Z]{2,8}(?:[&\-][A-Z0-9]{1,5})?)\s*"
        r"(/[A-Z]+(?:/[A-Z]+)*)\s*"
        r"([A-Z0-9]{2,5})?"
    )
    records: List[Dict[str, Any]] = []
    seen: set = set()
    for m in pattern.finditer(raw_text):
        subject = m.group(1)
        teacher_code = m.group(2)
        room = m.group(3) or ""
        # Filter false positives (e.g., legend codes themselves)
        if subject.startswith("SAKEC"):
            continue
        key = (subject, teacher_code, room)
        if key in seen:
            continue
        seen.add(key)
        records.append({
            "teacher": _resolve_teacher(teacher_code, legend),
            "subject": subject,
            "day": "",
            "slot": None,
            "time": "",
            "room": room,
            "branch": branch,
        })
    return records


# ======================================================================
# TEACHER DEDUPLICATION
# ======================================================================
def _dedup_teachers(
    records: List[Dict[str, Any]], threshold: int = 85,
) -> List[Dict[str, Any]]:
    """Use FuzzyWuzzy ratio to unify teacher name variants."""
    if not records:
        return records
    canonical_map: Dict[str, str] = {}
    seen_names: List[str] = []

    def get_canonical(name: str) -> str:
        if not name:
            return name
        if name in canonical_map:
            return canonical_map[name]
        best_match = None
        best_score = 0
        for existing in seen_names:
            score = fuzz.ratio(name.lower(), existing.lower())
            if score > best_score:
                best_score = score
                best_match = existing
        if best_score >= threshold and best_match:
            canonical_map[name] = best_match
            return best_match
        seen_names.append(name)
        canonical_map[name] = name
        return name

    for r in records:
        r["teacher"] = get_canonical(r.get("teacher", ""))
    return records


# ======================================================================
# PUBLIC API
# ======================================================================
def parse_timetable(
    tables: List[pd.DataFrame], raw_text: str,
) -> pd.DataFrame:
    """Parse SAKEC timetable into a long-format DataFrame.

    Output columns:
        teacher (str), subject (str), day (str), slot (int|None),
        time (str), room (str), branch (str).
    """
    legend = _parse_legend(raw_text)
    branch = _detect_branch(raw_text)

    all_records: List[Dict[str, Any]] = []

    tt_tables = _find_timetable_tables(tables) if tables else []
    for tt_df in tt_tables:
        try:
            all_records.extend(_parse_timetable_table(tt_df, legend, branch))
        except Exception as e:
            logger.warning("Timetable table parse failed: %s", e)

    # If tables produced very little, augment from raw text
    if len(all_records) < 5 and raw_text:
        text_records = _parse_timetable_from_text(raw_text, legend, branch)
        existing_pairs = {
            (r.get("teacher", ""), r.get("subject", ""), r.get("day", ""))
            for r in all_records
        }
        for tr in text_records:
            key = (tr["teacher"], tr["subject"], tr["day"])
            if key not in existing_pairs:
                all_records.append(tr)
                existing_pairs.add(key)

    if not all_records:
        return pd.DataFrame(columns=[
            "teacher", "subject", "day", "slot", "time", "room", "branch",
        ])

    all_records = _dedup_teachers(all_records, threshold=85)

    df = pd.DataFrame(all_records)
    # Ensure column order
    cols = ["teacher", "subject", "day", "slot", "time", "room", "branch"]
    for c in cols:
        if c not in df.columns:
            df[c] = "" if c != "slot" else None
    df = df[cols]
    # Coerce slot to nullable Int64
    try:
        df["slot"] = pd.to_numeric(df["slot"], errors="coerce").astype("Int64")
    except Exception:
        pass
    # Drop exact duplicates
    df = df.drop_duplicates().reset_index(drop=True)
    return df


def teacher_workload(df: pd.DataFrame) -> pd.DataFrame:
    """One row per teacher with weekly workload metrics."""
    if df is None or df.empty or "teacher" not in df.columns:
        return pd.DataFrame(columns=[
            "teacher_name", "total_weekly_lectures",
            "subjects", "rooms", "days_active", "branches",
        ])

    working = df.copy()
    # Drop blank teachers and RECESS rows (already excluded but defensive)
    working = working[working["teacher"].astype(str).str.strip() != ""]
    if working.empty:
        return pd.DataFrame(columns=[
            "teacher_name", "total_weekly_lectures",
            "subjects", "rooms", "days_active", "branches",
        ])

    def _join_unique(series: pd.Series) -> str:
        items = sorted({
            str(v).strip() for v in series
            if v is not None and str(v).strip()
        })
        return ", ".join(items)

    grouped = working.groupby("teacher", dropna=False).agg(
        total_weekly_lectures=("subject", "count"),
        subjects=("subject", _join_unique),
        rooms=("room", _join_unique),
        days_active=("day", _join_unique),
        branches=("branch", _join_unique),
    ).reset_index()
    grouped = grouped.rename(columns={"teacher": "teacher_name"})
    grouped = grouped.sort_values(
        "total_weekly_lectures", ascending=False,
    ).reset_index(drop=True)
    return grouped


def subject_frequency(df: pd.DataFrame) -> Dict[str, int]:
    """Return {subject: total weekly count}."""
    if df is None or df.empty or "subject" not in df.columns:
        return {}
    series = df["subject"].astype(str).str.strip()
    series = series[series != ""]
    if series.empty:
        return {}
    counts = series.value_counts()
    return {str(k): int(v) for k, v in counts.items()}


def branch_schedule(df: pd.DataFrame, branch: str) -> pd.DataFrame:
    """Filter rows by branch (case-insensitive substring match)."""
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    if "branch" not in df.columns or not branch:
        return df
    mask = df["branch"].astype(str).str.lower().str.contains(
        branch.lower(), na=False,
    )
    return df[mask].reset_index(drop=True)


def get_free_slots(df: pd.DataFrame) -> Dict[str, List[int]]:
    """Return {day: [free period numbers]} for Monday-Saturday."""
    if df is None or df.empty:
        return {d: [] for d in ORDERED_DAYS}

    # Determine the full period range from the data
    try:
        slots_numeric = pd.to_numeric(df.get("slot"), errors="coerce").dropna()
        if len(slots_numeric) > 0:
            min_p = int(slots_numeric.min())
            max_p = int(slots_numeric.max())
            if max_p < min_p:
                min_p, max_p = 1, 8
        else:
            min_p, max_p = 1, 8
    except Exception:
        min_p, max_p = 1, 8

    all_periods = set(range(min_p, max_p + 1))
    out: Dict[str, List[int]] = {}
    for day in ORDERED_DAYS:
        try:
            day_rows = df[df["day"].astype(str) == day]
            used = set()
            for s in day_rows.get("slot", []):
                try:
                    if pd.notna(s):
                        used.add(int(s))
                except (ValueError, TypeError):
                    continue
            out[day] = sorted(all_periods - used)
        except Exception:
            out[day] = sorted(all_periods)
    return out


# ======================================================================
# LLM QUERY
# ======================================================================
def answer_query_with_llm(
    df: pd.DataFrame, query: str, groq_client,
) -> str:
    """Send a natural-language timetable query to Groq llama3-8b-8192."""
    if df is None or df.empty:
        return "No timetable data available to answer the query."
    if not query or not str(query).strip():
        return "Empty query."

    try:
        context = df.to_string(index=False)
    except Exception as e:
        context = f"(Could not render dataframe: {e})"
    # Cap context to keep prompt within model limits
    if len(context) > 6000:
        context = context[:6000] + "\n... (truncated)"

    system_msg = (
        "You are a timetable coordinator at SAKEC Mumbai. Answer queries "
        "about class schedules, teacher workloads, room allocations, and "
        "lecture frequencies."
    )
    user_msg = (
        f"Timetable data (long-format, one row per slot):\n{context}\n\n"
        f"User query: {query}"
    )

    try:
        response = groq_client.chat.completions.create(
            model="llama3-8b-8192",
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            max_tokens=400,
        )
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
def process_timetable(
    extractor_output: Dict[str, Any], query: str, groq_client,
) -> Dict[str, Any]:
    """End-to-end timetable processing pipeline.

    Returns:
        {
            "dataframe":       pd.DataFrame,
            "query_answer":    str,
            "workload_df":     pd.DataFrame,
            "subject_freq":    dict,
            "teacher_legend":  dict,
        }
    """
    tables = extractor_output.get("tables", []) if extractor_output else []
    raw_text = extractor_output.get("raw_text", "") if extractor_output else ""

    df = parse_timetable(tables, raw_text)
    workload = teacher_workload(df)
    freq = subject_frequency(df)
    legend = _parse_legend(raw_text)

    if df.empty:
        answer = (
            "Could not parse any timetable entries from the document. "
            "Please verify the PDF is a SAKEC-format timetable."
        )
    else:
        answer = answer_query_with_llm(df, query, groq_client)

    return {
        "dataframe": df,
        "query_answer": answer,
        "workload_df": workload,
        "subject_freq": freq,
        "teacher_legend": legend,
    }
