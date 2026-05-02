"""Export DataFrames to downloadable bytes (CSV or Excel)."""

from __future__ import annotations

from enum import Enum
from io import BytesIO

import pandas as pd

from .constants import CSV_ENCODING


class ExportFormat(str, Enum):
    CSV = "csv"
    XLSX = "xlsx"


def _export_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Replace nulls with blank strings so every column shows empty cells, not NaN."""
    return df.fillna("")


def dataframe_to_bytes(df: pd.DataFrame, fmt: ExportFormat) -> tuple[bytes, str, str]:
    """
    Serialize a DataFrame for Streamlit download_button.

    Returns
    -------
    data : bytes
    mime_type : str
    default_filename : str
    """
    export_df = _export_frame(df)

    if fmt == ExportFormat.CSV:
        buffer = BytesIO()
        export_df.to_csv(buffer, index=False, encoding=CSV_ENCODING)
        return buffer.getvalue(), "text/csv; charset=utf-8", "export.csv"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Sheet1")
    return buffer.getvalue(), (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), "export.xlsx"
