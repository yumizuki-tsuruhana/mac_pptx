"""
Entry point for the .app bundle and PyInstaller.

Sets up sys.path so the mac_pptx_converter package can be found
regardless of where this script lives (repo root, .app bundle, frozen).
"""

import sys
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from mac_pptx_converter.gui import main

if __name__ == "__main__":
    main()
