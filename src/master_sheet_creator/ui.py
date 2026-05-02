"""Streamlit UI composition for Master Sheet Creator."""

from __future__ import annotations

import streamlit as st

from .columns import subset_columns
from .constants import STREAMLIT_UPLOAD_TYPES
from .exporter import ExportFormat, dataframe_to_bytes
from .loader import FileParseError, UnsupportedFileError, load_uploaded_file


def render_app() -> None:
    """Configure page and render the full application."""
    st.set_page_config(
        page_title="Master Sheet Creator",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.title("Master Sheet Creator")
    st.caption(
        "Upload a spreadsheet, pick the columns you need, and download a new file."
    )
    _render_upload_flow()


def _render_upload_flow() -> None:
    uploaded = st.file_uploader(
        "Upload spreadsheet",
        type=STREAMLIT_UPLOAD_TYPES,
        help="Supported formats: CSV (.csv) and Excel (.xlsx).",
    )

    if uploaded is None:
        st.info("Upload a CSV or Excel file to begin.")
        return

    try:
        df = load_uploaded_file(uploaded.name, uploaded.getvalue())
    except UnsupportedFileError as err:
        st.error(str(err))
        return
    except FileParseError as err:
        st.error(str(err))
        return

    st.subheader("Preview (first 50 rows)")
    st.dataframe(df.head(50), use_container_width=True)

    all_columns = list(df.columns)
    selected = st.multiselect(
        "Columns to include in the output",
        options=all_columns,
        default=all_columns,
        format_func=str,
    )

    if not selected:
        st.warning("Choose at least one column to export.")
        return

    try:
        output_df = subset_columns(df, selected)
    except KeyError as err:
        st.error(str(err))
        return

    fmt = st.radio(
        "Output format",
        options=list(ExportFormat),
        format_func=lambda f: "CSV (.csv)"
        if f == ExportFormat.CSV
        else "Excel (.xlsx)",
        horizontal=True,
    )

    data, mime, default_name = dataframe_to_bytes(output_df, fmt)
    st.download_button(
        "Download formatted file",
        data=data,
        file_name=default_name,
        mime=mime,
    )
