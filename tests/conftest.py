from __future__ import annotations

import sys
from pathlib import Path

# Ensure repo-root modules (e.g., lic_dsf_input_setters.py) are importable in tests.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

