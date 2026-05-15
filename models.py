"""
models.py
T.A.N.G.S - Transformative AI-Powered NLP & Generation Suite
Data models for the query-driven document processing pipeline.

Author: SAKEC Mumbai - Final Year B.Tech Project
Python: 3.9+
"""
from dataclasses import dataclass, field
from typing import Dict, List, Any


@dataclass
class StudentRecord:
    """A single student's academic result record."""
    roll_no: str
    name: str
    marks: Dict[str, float] = field(default_factory=dict)
    total: float = 0.0
    pass_fail: str = ""
    kt_count: int = 0


@dataclass
class TeacherSchedule:
    """A single teacher schedule entry (one slot on one day)."""
    name: str
    subject: str
    day: str
    slot: str
    room: str
    branch: str


@dataclass
class FinancialKPI:
    """A single financial KPI extracted from a financial document."""
    metric_name: str
    value: float
    period: str
    unit: str


@dataclass
class ProcessingResult:
    """The full result of running the pipeline on a document."""
    doc_type: str
    method_used: str
    dataframe: Any
    query_answer: str
    output_files: List[str] = field(default_factory=list)
