"""Keep long IDs as plain text end-to-end (no Excel scientific notation / float damage)."""

from __future__ import annotations

import numpy as np
import pandas as pd

from .constants import LONG_NUMERIC_ID_COLUMNS


def plain_id_string(value: object) -> str:
    """Normalize one cell to a full digit/string (never scientific notation)."""
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


def coerce_dataframe_long_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Force configured ID columns to plain Python strings (object dtype)."""
    if df.empty:
        return df
    out = df.copy()
    for col in LONG_NUMERIC_ID_COLUMNS:
        if col not in out.columns:
            continue
        out[col] = out[col].map(plain_id_string).astype(object)
    return out
