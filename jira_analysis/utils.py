"""Utility functions for file handling and data preparation."""

import os
import glob
from typing import Optional

def find_input_file() -> str:
    """Find the most recent Excel file in the input directory.
    
    Returns:
        str: Path to the most recent Excel file
        
    Raises:
        FileNotFoundError: If no Excel files are found in the input directory
    """
    files = glob.glob("input/*.xlsx")
    if not files:
        raise FileNotFoundError(
            "No Excel files found in the input directory. "
            "Please place your Jira export file in the input folder."
        )
    # Get the most recently modified file
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def ensure_directory(directory: str) -> None:
    """Ensure a directory exists, create it if it doesn't.
    
    Args:
        directory: Path to the directory
    """
    os.makedirs(directory, exist_ok=True)
