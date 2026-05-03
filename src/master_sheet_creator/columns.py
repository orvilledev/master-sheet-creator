"""Column selection helpers."""

from __future__ import annotations

from collections.abc import Sequence

import pandas as pd

from .constants import OUTPUT_COLUMN_ALIASES, OUTPUT_HEADER_DISPLAY


def subset_columns(df: pd.DataFrame, selected: list[str]) -> pd.DataFrame:
    """Return df with only ``selected`` columns, preserving order."""
    missing = [c for c in selected if c not in df.columns]
    if missing:
        raise KeyError(f"Unknown columns: {missing}")
    return df.loc[:, list(selected)]


def _source_column_for_preset(df: pd.DataFrame, preset_name: str) -> str | None:
    for key in OUTPUT_COLUMN_ALIASES.get(preset_name, (preset_name,)):
        if key in df.columns:
            return str(key)
    return None


def dataframe_fixed_output(df: pd.DataFrame, preset: Sequence[str]) -> pd.DataFrame:
    """
    Build a frame with **every** ``preset`` column in order.

    Columns missing from ``df``, or present but all-null, still appear; missing
    columns are filled with empty strings so exports always include the full layout.
    """
    series_list: list[pd.Series] = []
    for name in preset:
        label = OUTPUT_HEADER_DISPLAY.get(name, name)
        src = _source_column_for_preset(df, name)
        if src is not None:
            s = df[src].copy()
            s.name = label
            series_list.append(s)
        else:
            series_list.append(
                pd.Series("", index=df.index, dtype=object, name=label)
            )
    if not series_list:
        return pd.DataFrame(index=df.index)
    return pd.concat(series_list, axis=1)


def preset_columns_present(preset: Sequence[str], df_columns: Sequence) -> tuple[list[str], list[str]]:
    """Return (present, missing) preset names relative to ``df`` column labels (as strings)."""
    col_set = {str(c) for c in df_columns}
    present: list[str] = []
    missing: list[str] = []
    for n in preset:
        aliases = OUTPUT_COLUMN_ALIASES.get(n, (n,))
        if any(a in col_set for a in aliases):
            present.append(n)
        else:
            missing.append(n)
    return present, missing
