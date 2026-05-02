"""Shared constants for file handling and UI."""

UPLOAD_EXTENSIONS = {".csv", ".xlsx"}
STREAMLIT_UPLOAD_TYPES = ["csv", "xlsx"]

# UTF-8 with BOM helps Excel on Windows recognize UTF-8 CSV correctly.
CSV_ENCODING = "utf-8-sig"
