#!/usr/bin/env python3
"""
Form Filler Application
Ajman University - Main Store Forms Generator

This application generates PDF forms for:
- Acknowledgment of Receipt
- Asset Transfer Form (ATF)

Features:
- Dynamic row addition for assets/items
- Auto-fill current date
- Digital signature with name and timestamp
"""

import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.gui.main_window import MainWindow


def main():
    """Main entry point"""
    print("Starting Ajman University Form Filler...")
    print("=" * 50)

    app = MainWindow()
    app.run()

    print("Application closed.")


if __name__ == "__main__":
    main()
