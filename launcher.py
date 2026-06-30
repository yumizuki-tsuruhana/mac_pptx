"""
Standalone entry point for PyInstaller .app bundle.

PyInstaller can't use relative imports (from .converter import ...),
so this launcher uses absolute imports to bootstrap the GUI.
"""

import sys
import os

if getattr(sys, "frozen", False):
    base = sys._MEIPASS
    sys.path.insert(0, base)

from mac_pptx_converter.gui import main

if __name__ == "__main__":
    main()
