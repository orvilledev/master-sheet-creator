"""Load uploaded spreadsheet files into pandas DataFrames."""

from __future__ import annotations

from io import BytesIO

import pandas as pd

from .constants import UPLOAD_EXTENSIONS


class UnsupportedFileError(ValueError):
    """Raised when the file extension is not supported."""


class FileParseError(RuntimeError):
    """Raised when bytes cannot be parsed as the declared spreadsheet format."""


def extension_for(filename: str) -> str:
    lower = filename.lower()
    for ext in sorted(UPLOAD_EXTENSIONS, key=len, reverse=True):
        if lower.endswith(ext):
            return ext
    return ""


def load_uploaded_file(filename: str, raw_bytes: bytes) -> pd.DataFrame:
    """
    Parse an uploaded file into a DataFrame.

    Parameters
    ----------
    filename : str
        Original filename (used only to detect format).
    raw_bytes : bytes
        Full file contents.
    """
    ext = extension_for(filename)
    if ext not in UPLOAD_EXTENSIONS:
        raise UnsupportedFileError(
            f"Unsupported file type; allowed: {', '.join(sorted(UPLOAD_EXTENSIONS))}"
        )

    buffer = BytesIO(raw_bytes)
    try:
        if ext == ".csv":
            return pd.read_csv(buffer)
        if ext == ".xlsx":
            return pd.read_excel(buffer, engine="openpyxl")
    except Exception as err:
        raise FileParseError(f"Failed to parse file as {ext}: {err}") from err

    raise UnsupportedFileError(f"No loader registered for {ext!r}")
