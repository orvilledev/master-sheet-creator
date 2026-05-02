"""Column selection helpers."""

from __future__ import annotations

import pandas as pd


def subset_columns(df: pd.DataFrame, selected: list[str]) -> pd.DataFrame:
    """Return df with only ``selected`` columns, preserving order."""
    missing = [c for c in selected if c not in df.columns]
    if missing:
        raise KeyError(f"Unknown columns: {missing}")
    return df.loc[:, list(selected)]
