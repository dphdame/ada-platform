"""Abstract base fixer interface."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Optional

from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import FixResult, ScanResult


class BaseFixer(ABC):
    """Format-agnostic fixer interface.

    Subclasses implement fix() for each document format:
    - DocxFixer (Phase 1)
    - PptxFixer (Phase 2)
    - PdfFixer (Phase 3)
    """

    @abstractmethod
    def fix(
        self,
        file_path: str,
        output_path: str,
        scan_result: ScanResult,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> FixResult:
        """Apply safe fixes and return a FixResult."""
        ...
