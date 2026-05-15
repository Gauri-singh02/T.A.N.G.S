"""
financial_module.py
T.A.N.G.S - Financial Document Analytics Module

Extracts financial KPIs (revenue, net profit, EPS, margin, growth) from
PDFs/tables, computes YoY growth, and answers natural-language queries
via Groq LLM.

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import re
import math
import logging
from typing import Dict, List, Any, Optional

import pandas as pd


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# Regex patterns per the spec. All applied with re.IGNORECASE.
REGEX_PATTERNS: Dict[str, str] = {
    "revenue": r"(?:revenue|turnover|net sales)[^\d]*?([\d,]+\.?\d*)",
    "net_profit": r"(?:net profit|net income|pat)[^\d]*?([\d,]+\.?\d*)",
    "eps": r"(?:eps|earnings per share)[^\d]*?([\d.]+)",
    "margin": r"(?:margin)[^\d]*?([\d.]+)\s*%",
    "growth": r"(?:growth|yoy)[^\d]*?([-\d.]+)\s*%",
}


# ======================================================================
# CORE EXTRACTION
# ======================================================================
def _to_float(raw: str) -> Optional[float]:
    """Convert a captured numeric string like '1,234.56' to float."""
    if raw is None:
        return None
    s = str(raw).strip().replace(",", "")
    if not s or s in {"-", "."}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def extract_kpis_from_text(raw_text: str) -> Dict[str, float]:
    """Apply REGEX_PATTERNS to free text and return the first match per KPI."""
    if not raw_text:
        return {}
    out: Dict[str, float] = {}
    for kpi, pattern in REGEX_PATTERNS.items():
        try:
            m = re.search(pattern, raw_text, re.IGNORECASE)
        except re.error as e:
            logger.warning("Regex error for %s: %s", kpi, e)
            continue
        if not m:
            continue
        val = _to_float(m.group(1))
        if val is not None:
            out[kpi] = val
    return out


def extract_kpis_from_tables(
    tables: List[pd.DataFrame],
) -> Dict[str, float]:
    """Flatten each table's cells, then apply REGEX_PATTERNS."""
    if not tables:
        return {}
    out: Dict[str, float] = {}
    for tbl in tables:
        if tbl is None or tbl.empty:
            continue
        try:
            # Include column headers so labels next to value columns match
            header_text = " ".join(str(c) for c in tbl.columns)
            body_text = " ".join(tbl.astype(str).values.flatten())
            flat = header_text + " " + body_text
        except Exception as e:
            logger.warning("Failed to flatten table: %s", e)
            continue
        for kpi, pattern in REGEX_PATTERNS.items():
            if kpi in out:
                # Keep the first table's match (deterministic)
                continue
            try:
                m = re.search(pattern, flat, re.IGNORECASE)
            except re.error:
                continue
            if not m:
                continue
            val = _to_float(m.group(1))
            if val is not None:
                out[kpi] = val
    return out


def yoy_growth(current: Any, previous: Any) -> float:
    """Year-over-Year growth rate: (Mt - Mt-1) / Mt-1 * 100.

    Returns 0.0 if previous is None / NaN / 0.
    """
    if current is None or previous is None:
        return 0.0
    try:
        c = float(current)
        p = float(previous)
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(c) or math.isnan(p) or p == 0:
        return 0.0
    return ((c - p) / p) * 100.0


# ======================================================================
# LLM QUERY
# ======================================================================
def answer_query_with_llm(
    raw_text: str, query: str, groq_client,
) -> str:
    """Send a financial query + the document text to Groq llama3-8b-8192."""
    if not raw_text:
        return "No document text available to answer the query."
    if not query or not str(query).strip():
        return "Empty query."

    # Cap context so we stay within model limits
    context = raw_text
    if len(context) > 8000:
        context = context[:8000] + "\n... (truncated)"

    system_msg = (
        "You are a financial analyst. Extract and analyze KPIs from "
        "financial documents."
    )
    user_msg = (
        f"Financial document text:\n{context}\n\n"
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
def process_financial(
    extractor_output: Dict[str, Any], query: str, groq_client,
) -> Dict[str, Any]:
    """End-to-end financial document processing pipeline.

    Returns:
        {
            "dataframe":    pd.DataFrame,   # one-row KPI table
            "query_answer": str,
            "kpis":         dict,           # merged regex hits
            "yoy_growth":   float | None,
        }
    """
    tables = extractor_output.get("tables", []) if extractor_output else []
    raw_text = extractor_output.get("raw_text", "") if extractor_output else ""

    # Pull KPIs from both sources, prefer the table value when both exist
    kpis_text = extract_kpis_from_text(raw_text)
    kpis_tables = extract_kpis_from_tables(tables)
    merged: Dict[str, float] = dict(kpis_text)
    merged.update(kpis_tables)

    # Compute YoY if we have an explicit growth hit, otherwise None
    yoy: Optional[float] = merged.get("growth")

    if merged:
        df = pd.DataFrame([merged])
    else:
        df = pd.DataFrame()

    if not raw_text and not merged:
        answer = "Could not extract any financial data from the document."
    else:
        answer = answer_query_with_llm(raw_text, query, groq_client)

    return {
        "dataframe": df,
        "query_answer": answer,
        "kpis": merged,
        "yoy_growth": yoy,
    }
