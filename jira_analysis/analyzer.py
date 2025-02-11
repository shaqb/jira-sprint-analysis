"""Core analysis functionality for Jira sprint data."""

import pandas as pd
import numpy as np
from typing import Dict, Any

def calculate_percentage_diff(original: float, actual: float) -> float:
    """Calculate percentage difference between original and actual values.
    
    Args:
        original: Original estimate value
        actual: Actual time spent
    
    Returns:
        float: Percentage difference
    """
    if pd.isna(original) or pd.isna(actual) or original == 0:
        return 0
    return ((actual - original) / original) * 100

def process_discipline_data(df: pd.DataFrame, discipline_mapping: Dict[str, Dict[str, str]]) -> Dict[str, Any]:
    """Process data for a specific discipline.
    
    Args:
        df: DataFrame containing the Jira sprint data
        discipline_mapping: Dictionary mapping disciplines to their column names
    
    Returns:
        dict: Processed statistics for the discipline
    """
    result = {
        "Complete Data Points": 0,
        "Partial Data Points": 0,
        "Missing Data Points": 0,
        "Original Estimate (Total)": 0,
        "AI Estimate (Total)": 0,
        "Actual Time (Total)": 0,
    }

    # Create sets to track unique tickets
    complete_tickets = set()
    partial_tickets = set()
    missing_tickets = set()

    for _, row in df.iterrows():
        original_col = discipline_mapping["original"]
        ai_col = discipline_mapping["ai"]
        actual_col = discipline_mapping["actual"]

        has_data = False
        if pd.notna(row[original_col]):
            has_data = True
            result["Original Estimate (Total)"] += float(row[original_col])
        if pd.notna(row[ai_col]):
            has_data = True
            result["AI Estimate (Total)"] += float(row[ai_col])
        if pd.notna(row[actual_col]):
            has_data = True
            result["Actual Time (Total)"] += float(row[actual_col])

        if has_data:
            if pd.notna(row[original_col]) and pd.notna(row[ai_col]) and pd.notna(row[actual_col]):
                complete_tickets.add(row["Key"])
            else:
                partial_tickets.add(row["Key"])
        else:
            missing_tickets.add(row["Key"])

    result["Complete Data Points"] = len(complete_tickets)
    result["Partial Data Points"] = len(partial_tickets)
    result["Missing Data Points"] = len(missing_tickets)

    return result
