"""
CITES Operations Intelligence Framework
Enterprise Issue Triage, Workforce Accountability & Support Chat Mining Toolkit.
"""

__version__ = "1.0.0"

from .core.classifier import IssueClassifier
from .core.entity_matcher import EntityMatcher
from .core.chat_parser import ChatParser, ChatKnowledgeExtractor
from .core.workforce import WorkforceMapper
from .core.ingest import IngestValidator
from .reporters.excel_reporter import ExcelReporter
from .reporters.pptx_reporter import PPTXReporter
from .reporters.docx_reporter import DocxReporter
from .reporters.html_reporter import HTMLReporter

__all__ = [
    "IssueClassifier",
    "EntityMatcher",
    "ChatParser",
    "ChatKnowledgeExtractor",
    "WorkforceMapper",
    "IngestValidator",
    "ExcelReporter",
    "PPTXReporter",
    "DocxReporter",
    "HTMLReporter",
]
