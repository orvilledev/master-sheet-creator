"""Column selection helpers."""

from __future__ import annotations

from collections.abc import Sequence

import pandas as pd


def subset_columns(df: pd.DataFrame, selected: list[str]) -> pd.DataFrame:
    """Return df with only ``selected`` columns, preserving order."""
    missing = [c for c in selected if c not in df.columns]
    if missing:
        raise KeyError(f"Unknown columns: {missing}")
    return df.loc[:, list(selected)]


def dataframe_fixed_output(df: pd.DataFrame, preset: Sequence[str]) -> pd.DataFrame:
    """
    Build a frame with columns ``preset`` in order.

    Columns missing from ``df`` are filled with ``pd.NA``.
    """
    out = pd.DataFrame(index=df.index)
    for name in preset:
        out[name] = df[name] if name in df.columns else pd.Series(pd.NA, index=df.index)
    return out


def preset_columns_present(preset: Sequence[str], df_columns: Sequence) -> tuple[list[str], list[str]]:
    """Return (present, missing) preset names relative to ``df`` column labels (as strings)."""
    col_set = {str(c) for c in df_columns}
    present = [n for n in preset if n in col_set]
    missing = [n for n in preset if n not in col_set]
    return present, missing
