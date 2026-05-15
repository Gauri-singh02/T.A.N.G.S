"""
query_router.py
T.A.N.G.S - Query Intent Router

Routes natural-language queries to one of the supported processing
intents (topper / kt / workload / subject_stats / summary / kpi).

Strategy:
  1. Keyword match against INTENT_PATTERNS.
  2. If any keyword hits, prefer doc-type-specific intents among the
     top-scoring ones.
  3. If nothing hits, fall back to TF-IDF cosine similarity (sklearn).
  4. If max cosine similarity < 0.25, return 'llm_fallback'.

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
import re
import logging
from typing import Dict, List

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False


logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


# ----------------------------------------------------------------------
# Intent keyword patterns
# ----------------------------------------------------------------------
INTENT_PATTERNS: Dict[str, List[str]] = {
    "topper": [
        "top", "topper", "highest marks", "rank",
        "best student", "first", "scored most", "maximum marks",
    ],
    "kt": [
        "kt", "backlog", "fail", "failed", "atkt",
        "arrear", "not passed", "detained",
    ],
    "workload": [
        "workload", "load", "how many lectures",
        "classes per week", "schedule of", "timetable of",
    ],
    "subject_stats": [
        "average", "pass percentage", "failure rate",
        "how many passed", "statistics", "performance in",
    ],
    "summary": [
        "summarize", "summary", "overview",
        "total students", "overall result", "how many students",
    ],
    "kpi": [
        "revenue", "profit", "eps", "margin",
        "growth", "earnings", "financial",
    ],
}

# Doc-type -> ordered priority list (higher priority listed first)
_DOC_PRIORITY: Dict[str, List[str]] = {
    "result": ["topper", "kt", "subject_stats", "summary"],
    "timetable": ["workload"],
    "financial": ["kpi"],
}

# Suggested queries per doc type
_SUGGESTED: Dict[str, List[str]] = {
    "result": [
        "Who topped overall?",
        "List all KT students",
        "What is pass % in Applied Physics?",
        "Show subject wise failure rate",
        "Who failed in more than 2 subjects?",
    ],
    "timetable": [
        "Workload of Manjusha Kulkarni?",
        "What subjects on Monday?",
        "List all teachers",
        "Which room has most lectures?",
        "Free slots on Friday?",
    ],
    "financial": [
        "What is the revenue this year?",
        "Show net profit and EPS",
        "Calculate year over year growth",
        "What is the profit margin?",
        "Summarize the key financial KPIs",
    ],
}

# TF-IDF similarity threshold for fallback routing
_TFIDF_THRESHOLD: float = 0.25


# ----------------------------------------------------------------------
# Routing
# ----------------------------------------------------------------------
def route_query(query: str, doc_type: str) -> str:
    """Return the routed intent name, or 'llm_fallback' when uncertain.

    Args:
        query:    The user's natural-language query.
        doc_type: One of 'result', 'timetable', 'financial' (or other).
    """
    if not query or not str(query).strip():
        return "llm_fallback"

    q = str(query).lower().strip()
    # Normalize '%' to the word 'percentage' so queries like 'pass %'
    # match the 'pass percentage' keyword in subject_stats.
    q = q.replace("%", " percentage ")
    q = re.sub(r"\s+", " ", q).strip()
    doc_type_norm = (doc_type or "").lower().strip()
    priorities = _DOC_PRIORITY.get(doc_type_norm, [])

    # 1) Keyword-based scoring.
    # Weight each match by the keyword's length so that a long, specific
    # phrase like "failure rate" outranks a short generic substring like
    # "fail" (which is itself a substring of "failure"). This avoids the
    # 'kt' intent stealing queries that are clearly about subject stats.
    scores: Dict[str, int] = {}
    for intent, keywords in INTENT_PATTERNS.items():
        s = 0
        for kw in keywords:
            if kw in q:
                s += len(kw)
        if s > 0:
            scores[intent] = s

    if scores:
        max_score = max(scores.values())
        top_intents = [i for i, sc in scores.items() if sc == max_score]
        # Among top-scoring intents, prefer the doc-type priority order.
        for p in priorities:
            if p in top_intents:
                return p
        # Otherwise pick deterministically by INTENT_PATTERNS declaration order.
        for intent in INTENT_PATTERNS.keys():
            if intent in top_intents:
                return intent
        return top_intents[0]

    # 2) TF-IDF cosine similarity fallback
    if SKLEARN_AVAILABLE:
        try:
            intents = list(INTENT_PATTERNS.keys())
            docs = [
                " ".join(INTENT_PATTERNS[i]) for i in intents
            ] + [q]
            vectorizer = TfidfVectorizer().fit(docs)
            vecs = vectorizer.transform(docs)
            sims = cosine_similarity(vecs[-1], vecs[:-1])[0]
            if len(sims) == 0:
                return "llm_fallback"
            max_sim = float(sims.max())
            if max_sim < _TFIDF_THRESHOLD:
                return "llm_fallback"
            best_idx = int(sims.argmax())
            best_intent = intents[best_idx]
            # Bias by doc-type priority on near-ties
            if priorities:
                # If a priority intent is within 0.05 of best, prefer it
                for p in priorities:
                    if p in intents:
                        p_idx = intents.index(p)
                        if float(sims[p_idx]) >= max_sim - 0.05:
                            return p
            return best_intent
        except Exception as e:
            logger.warning("TF-IDF routing failed: %s", e)
            return "llm_fallback"

    # 3) sklearn not available, give up gracefully
    return "llm_fallback"


def get_suggested_queries(doc_type: str) -> List[str]:
    """Return 5 example queries appropriate for the given doc type."""
    if not doc_type:
        return []
    key = str(doc_type).lower().strip()
    return list(_SUGGESTED.get(key, []))
