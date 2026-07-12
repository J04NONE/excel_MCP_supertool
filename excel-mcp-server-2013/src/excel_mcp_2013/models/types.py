"""Tipos compartidos del servidor."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class ExcelSessionState:
    process_id: Optional[int] = None
    excel_version: str = ""
    is_64bit: bool = False
    active_workbook_count: int = 0
    status: str = "initialized"
