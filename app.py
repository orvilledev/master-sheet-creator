"""
Streamlit entrypoint for Master Sheet Creator.

Run: streamlit run app.py
"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if _SRC.exists() and str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from master_sheet_creator.ui import render_app


def main() -> None:
    render_app()


if __name__ == "__main__":
    main()
