from pathlib import Path
from typing import Dict, List
import pandas as pd


_EXCEL_PATH = Path(__file__).resolve().parents[1] / "strategic_insight.xlsx"


class GoogleSheetsFallback:
    """Minimal shim that mimics subset of existing google_sheets service."""

    def __init__(self, workbook_path: Path):
        self.workbook = workbook_path

    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        if not self.workbook.exists():
            raise FileNotFoundError(f"Workbook not found: {self.workbook}")
        try:
            return pd.read_excel(self.workbook, sheet_name=sheet_name)
        except ValueError:
            # sheet missing -> empty frame
            return pd.DataFrame()

    def sheet_to_df(self, sheet_name: str) -> pd.DataFrame:
        return self._load_sheet(sheet_name)

    def get_dna(self) -> pd.DataFrame:
        return self._load_sheet("dna")

    def get_roadmap(self) -> pd.DataFrame:
        return self._load_sheet("roadmap")

    def get_operation_health(self) -> pd.DataFrame:
        return self._load_sheet("operation_health")


def _load_excel_cache() -> Dict[str, pd.DataFrame]:
    if not _EXCEL_PATH.exists():
        return {}
    try:
        return pd.read_excel(_EXCEL_PATH, sheet_name=None)
    except Exception:
        return {}


def sheet_to_df(sheet_name: str) -> pd.DataFrame:
    cache = _load_excel_cache()
    return cache.get(sheet_name, pd.DataFrame())


def get_dna() -> pd.DataFrame:
    return sheet_to_df("dna")


def get_roadmap() -> pd.DataFrame:
    return sheet_to_df("roadmap")


def get_operation_health() -> pd.DataFrame:
    return sheet_to_df("operation_health")
