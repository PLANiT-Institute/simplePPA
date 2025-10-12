#!/usr/bin/env python3
"""
Simple launcher for the PPA Analysis GUI
"""

import sys
import os

# Add current directory to path to import gui_app
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui_app import main

if __name__ == "__main__":
    print("Starting PPA Analysis GUI...")
    main()