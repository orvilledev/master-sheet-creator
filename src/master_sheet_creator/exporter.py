"""Export DataFrames to downloadable bytes (CSV or Excel)."""

from __future__ import annotations

from enum import Enum
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

from .constants import CSV_ENCODING, LONG_NUMERIC_ID_COLUMNS


class ExportFormat(str, Enum):
    CSV = "csv"
    XLSX = "xlsx"


def _export_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Replace nulls with blank strings so every column shows empty cells, not NaN."""
    return df.fillna("")


def _plain_id_string(value: object) -> str:
    """Convert cell values to full digit/string form — avoids Excel scientific notation."""
    if value == "":
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass

    if isinstance(value, str):
        return value.strip()

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (np.integer, int)):
        return str(int(value))

    if isinstance(value, (np.floating, float)):
        fv = float(value)
        if np.isfinite(fv):
            r = round(fv)
            if abs(fv - r) < 1e-6:
                return str(int(r))
        return str(value)

    return str(value).strip()


def _apply_long_id_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in LONG_NUMERIC_ID_COLUMNS:
        if col in out.columns:
            out[col] = out[col].map(_plain_id_string)
    return out


def _set_xlsx_text_format(workbook, *, column_names: list[str]) -> None:
    ws = workbook["Sheet1"]
    name_to_idx = {name: i + 1 for i, name in enumerate(column_names)}
    for col_name in LONG_NUMERIC_ID_COLUMNS:
        if col_name not in name_to_idx:
            continue
        letter = get_column_letter(name_to_idx[col_name])
        for row in range(2, ws.max_row + 1):
            ws[f"{letter}{row}"].number_format = "@"


def dataframe_to_bytes(df: pd.DataFrame, fmt: ExportFormat) -> tuple[bytes, str, str]:
    """
    Serialize a DataFrame for Streamlit download_button.

    Returns
    -------
    data : bytes
    mime_type : str
    default_filename : str
    """
    export_df = _apply_long_id_text_columns(_export_frame(df))
    cols_list = list(export_df.columns)

    if fmt == ExportFormat.CSV:
        buffer = BytesIO()
        export_df.to_csv(buffer, index=False, encoding=CSV_ENCODING)
        return buffer.getvalue(), "text/csv; charset=utf-8", "export.csv"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Sheet1")
        _set_xlsx_text_format(writer.book, column_names=cols_list)

    return buffer.getvalue(), (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), "export.xlsx"
