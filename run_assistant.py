"""
Main Launcher for Kruti Dev 010 Drafting & Spellcheck Assistant for MS Word.
"""

import os
import sys

# Ensure current directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_assistant.assistant_app import launch_app

if __name__ == "__main__":
    print("Launching Kruti Dev 010 Smart Assistant for MS Word...")
    launch_app()

