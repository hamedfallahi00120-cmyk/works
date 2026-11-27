"""Multi-format local document ingestion for Local Office AI."""

from .models import ExtractedDocument, ExtractedTable
from .reader import DocumentReader, SUPPORTED_DOCUMENT_SUFFIXES

__all__ = ["DocumentReader", "ExtractedDocument", "ExtractedTable", "SUPPORTED_DOCUMENT_SUFFIXES"]
