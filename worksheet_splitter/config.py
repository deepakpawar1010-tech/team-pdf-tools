from __future__ import annotations
from pathlib import Path
import tempfile

FIRST_PAGE_BANNER_MARGIN: float = 45.0
SAME_PAGE_BOTTOM_MARGIN: float = 12.0

WORKSHEET_OUTPUT_DIR: Path = Path(tempfile.gettempdir()) / "worksheet_splitter_outputs"
