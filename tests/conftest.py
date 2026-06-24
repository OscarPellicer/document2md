from __future__ import annotations

import sys
from pathlib import Path


PACKAGE_ROOT = Path(__file__).resolve().parents[1] / "skills" / "office2md" / "scripts"
sys.path.insert(0, str(PACKAGE_ROOT))

