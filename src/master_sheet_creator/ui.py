"""Streamlit UI composition for Master Sheet Creator."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from .columns import dataframe_fixed_output, preset_columns_present
from .constants import OUTPUT_COLUMN_ORDER, STREAMLIT_UPLOAD_TYPES
from .exporter import ExportFormat, dataframe_to_bytes
from .loader import FileParseError, UnsupportedFileError, load_uploaded_file


def _dataframe_unique_columns_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    Streamlit (PyArrow) rejects duplicate column labels; the exported Excel may repeat headers.

    Second and later occurrences get `` (2)``, `` (3)``, … appended so previews render.
    """
    seen: dict[str, int] = {}
    names: list[str] = []
    for col in df.columns:
        key = str(col)
        if key not in seen:
            seen[key] = 1
            names.append(key)
        else:
            seen[key] += 1
            names.append(f"{key} ({seen[key]})")
    out = df.copy()
    out.columns = names
    return out


def render_app() -> None:
    """Configure page and render the full application."""
    st.set_page_config(
        page_title="Master Sheet Creator",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.title("Master Sheet Creator")
    st.caption(
        "Upload your NetSuite-style export (.xls XML, .xlsx, or .csv). "
        "The download is always **Excel (.xlsx)** matching your **format.xlsx** layout (122 columns); "
        "fields not in the upload are left blank. **Every row** is included — previews are shortened."
    )
    _render_upload_flow()


def _render_upload_flow() -> None:
    uploaded = st.file_uploader(
        "Upload spreadsheet",
        type=STREAMLIT_UPLOAD_TYPES,
        help="Supports CSV, Excel .xlsx, and .xls (including Excel 2003 XML saved as .xls).",
    )

    if uploaded is None:
        st.info("Upload an export file to begin.")
        return

    try:
        df = load_uploaded_file(uploaded.name, uploaded.getvalue())
    except UnsupportedFileError as err:
        st.error(str(err))
        return
    except FileParseError as err:
        st.error(str(err))
        return

    present, missing_preset = preset_columns_present(OUTPUT_COLUMN_ORDER, df.columns)
    n_rows, n_cols = len(df), len(df.columns)
    st.info(
        f"**{n_rows:,}** rows × **{n_cols}** columns in your upload. "
        f"The downloaded file will include **all {n_rows:,} rows** and **{len(OUTPUT_COLUMN_ORDER)}** output columns."
    )

    st.subheader("Preview (raw upload — first 50 rows only)")
    st.dataframe(df.head(50), use_container_width=True)

    with st.expander("Template columns matched from your upload"):
        st.write(
            f"**{len(present)} / {len(OUTPUT_COLUMN_ORDER)}** template columns appear in this file "
            "(by exact header name)."
        )
        if missing_preset:
            st.warning(
                f"**{len(missing_preset)}** template column(s) are not in the upload "
                "(they will be blank in the export)."
            )
            with st.expander("Show missing template column names"):
                st.text("\n".join(missing_preset))
        else:
            st.success("All template columns are present in the upload.")

    output_df = dataframe_fixed_output(df, OUTPUT_COLUMN_ORDER)

    st.subheader("Export preview (first 25 rows only)")
    st.dataframe(
        _dataframe_unique_columns_for_display(output_df.head(25)),
        use_container_width=True,
    )

    data, mime, default_name = dataframe_to_bytes(output_df, ExportFormat.XLSX)
    st.download_button(
        "Download Excel (.xlsx)",
        data=data,
        file_name=default_name,
        mime=mime,
    )
