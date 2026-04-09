"""Abstract base scanner interface."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Optional

from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import ScanResult


class BaseScanner(ABC):
    """Format-agnostic scanner interface.

    Subclasses implement scan() for each document format:
    - DocxScanner (Phase 1)
    - PptxScanner (Phase 2)
    - PdfScanner (Phase 3)
    """

    @abstractmethod
    def scan(
        self,
        file_path: str,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> ScanResult:
        """Scan a document and return accessibility issues found."""
        ...
