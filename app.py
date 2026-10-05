"""Main entrypoint for Viedoc EDC QC Automation Application."""

import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.gui import launch

if __name__ == "__main__":
    launch()
