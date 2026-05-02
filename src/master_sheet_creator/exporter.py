"""Export DataFrames to downloadable bytes (CSV or Excel)."""

from __future__ import annotations

from enum import Enum
from io import BytesIO

import pandas as pd
from openpyxl.utils import get_column_letter

from .constants import CSV_ENCODING, LONG_NUMERIC_ID_COLUMNS
from .id_format import coerce_dataframe_long_ids, plain_id_string

# Invisible LTR mark — Excel keeps the cell as text; digits stay visible without E+ notation.
_XLSX_TEXT_PREFIX = "\u200e"


class ExportFormat(str, Enum):
    CSV = "csv"
    XLSX = "xlsx"


def _export_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Replace nulls with blank strings so every column shows empty cells, not NaN."""
    return df.fillna("")


def _rewrite_openpyxl_long_id_cells(workbook, *, column_names: list[str]) -> None:
    """
    Open Excel sheet and force ID columns to real text cells.

    pandas/to_excel can still write numeric types; overwriting values here guarantees
    Excel displays full digits (plus ``number_format = '@'``).
    """
    ws = workbook["Sheet1"]
    name_to_idx = {name: i + 1 for i, name in enumerate(column_names)}

    for col_name in LONG_NUMERIC_ID_COLUMNS:
        if col_name not in name_to_idx:
            continue
        letter = get_column_letter(name_to_idx[col_name])
        ws[f"{letter}1"].number_format = "@"

        for row in range(2, ws.max_row + 1):
            cell = ws[f"{letter}{row}"]
            raw = cell.value
            if raw is None or raw == "":
                cell.value = ""
            else:
                text = plain_id_string(raw)
                # Prefix stops Excel from treating digit-only strings as numbers on open.
                cell.value = _XLSX_TEXT_PREFIX + text if text else ""
            cell.number_format = "@"


def _csv_text_for_excel_open(value: object) -> str:
    """
    Excel's CSV importer treats long digit fields as numbers (scientific notation).

    A text formula like ``="199230130516"`` forces the full value as text in Excel.
    """
    s = plain_id_string(value)
    if s == "":
        return ""
    escaped = s.replace('"', '""')
    return f'="{escaped}"'


def dataframe_to_bytes(df: pd.DataFrame, fmt: ExportFormat) -> tuple[bytes, str, str]:
    """
    Serialize a DataFrame for Streamlit download_button.

    Returns
    -------
    data : bytes
    mime_type : str
    default_filename : str
    """
    export_df = coerce_dataframe_long_ids(_export_frame(df))
    cols_list = list(export_df.columns)

    if fmt == ExportFormat.CSV:
        csv_df = export_df.copy()
        for col in LONG_NUMERIC_ID_COLUMNS:
            if col in csv_df.columns:
                csv_df[col] = csv_df[col].map(_csv_text_for_excel_open)
        buffer = BytesIO()
        csv_df.to_csv(buffer, index=False, encoding=CSV_ENCODING)
        return buffer.getvalue(), "text/csv; charset=utf-8", "export.csv"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Sheet1")
        _rewrite_openpyxl_long_id_cells(writer.book, column_names=cols_list)

    return buffer.getvalue(), (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), "export.xlsx"
