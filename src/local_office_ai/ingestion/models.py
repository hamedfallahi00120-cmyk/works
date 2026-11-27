"""Normalized records shared by all document readers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class ExtractedTable:
    source: Path
    location: str
    title: str
    headers: list[str]
    rows: list[list[Any]]
    truncated: bool = False
    warnings: list[str] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def column_count(self) -> int:
        return len(self.headers)


@dataclass
class ExtractedDocument:
    source: Path
    kind: str
    title: str
    text_blocks: list[tuple[str, str]] = field(default_factory=list)
    tables: list[ExtractedTable] = field(default_factory=list)
    page_count: int = 0
    slide_count: int = 0
    warnings: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def extracted_characters(self) -> int:
        return sum(len(text) for _, text in self.text_blocks) + sum(
            sum(len(str(value or "")) for value in row)
            for table in self.tables
            for row in table.rows
        )

    @property
    def extracted_rows(self) -> int:
        return sum(table.row_count for table in self.tables)

    def compact_text(self, limit: int = 8000) -> str:
        pieces: list[str] = []
        remaining = max(0, int(limit))
        for location, text in self.text_blocks:
            if remaining <= 0:
                break
            chunk = f"[{location}] {text.strip()}"
            if not chunk.strip():
                continue
            chunk = chunk[:remaining]
            pieces.append(chunk)
            remaining -= len(chunk)
        for table in self.tables:
            if remaining <= 0:
                break
            lines = [f"[جدول: {table.title} | {table.location}]", " | ".join(table.headers)]
            for row in table.rows[:30]:
                lines.append(" | ".join(str(value or "") for value in row))
            chunk = "\n".join(lines)[:remaining]
            pieces.append(chunk)
            remaining -= len(chunk)
        return "\n\n".join(pieces)
