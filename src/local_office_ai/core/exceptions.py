"""Typed exceptions used across the engine."""

from __future__ import annotations


class EngineError(Exception):
    """Base class for all engine errors."""


class WorkbookAccessError(EngineError):
    """The workbook could not be opened (missing, locked, permission denied)."""


class UnsupportedWorkbookError(EngineError):
    """The file is not a supported Excel workbook."""


class CorruptWorkbookError(EngineError):
    """The workbook exists but its structure cannot be parsed."""


class StructureDetectionError(EngineError):
    """No trustworthy table structure could be detected."""


class NoFinancialTableError(StructureDetectionError):
    """No table with financial meaning was found in the workbook."""


class NormalizationError(EngineError):
    """A value could not be normalized and the caller asked for strict mode."""


class AIServiceError(EngineError):
    """Ollama or the model could not be reached / returned an error."""


class AIResponseError(AIServiceError):
    """The model answered, but the answer is unusable or contradicts the data."""


class ComfyUIError(EngineError):
    """ComfyUI image generation failed or is unavailable."""


class ProcessingAborted(EngineError):
    """The operator or a guard stopped the run."""


class ReportFormatError(EngineError):
    """The requested output format is unsupported or does not match the writer."""


class ReportTemplateError(EngineError):
    """A DOCX report template is missing, unreadable or structurally unusable."""
