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
    if pd.isna(original) or pd.isna(actual):
        return 0
    return ((actual - original) / original) * 100 if original != 0 else 0

def calculate_improvement(original_total: float, ai_total: float, actual_total: float) -> Dict[str, float]:
    """Calculate percentage improvements in estimation accuracy.
    
    Args:
        original_total: Total of original estimates
        ai_total: Total of AI estimates
        actual_total: Total of actual time spent
    
    Returns:
        dict: Dictionary containing improvement percentages
    """
    if original_total == 0 or actual_total == 0:
        return {
            "original_vs_actual": 0,
            "ai_vs_actual": 0,
            "overall_improvement": 0
        }
    
    # Calculate percentage deviations from actual
    original_deviation = ((original_total - actual_total) / actual_total) * 100
    ai_deviation = ((ai_total - actual_total) / actual_total) * 100
    
    # Calculate absolute error percentages
    original_error = abs(original_deviation)
    ai_error = abs(ai_deviation)
    
    # Calculate improvement as reduction in error percentage
    # If original error was 0, there can't be any improvement
    if original_error == 0:
        improvement = 0
    else:
        # Calculate how much the error was reduced
        error_reduction = original_error - ai_error
        # Convert to percentage of original error
        improvement = (error_reduction / original_error) * 100
        # Cap at 100%
        improvement = min(100, max(0, improvement))
    
    return {
        "original_vs_actual": original_deviation,
        "ai_vs_actual": ai_deviation,
        "overall_improvement": improvement
    }

def process_discipline_data(df: pd.DataFrame, discipline_mapping: Dict[str, str]) -> Dict[str, Any]:
    """Process data for a specific discipline.
    
    Args:
        df: DataFrame containing the Jira sprint data
        discipline_mapping: Dictionary mapping disciplines to their column names
    
    Returns:
        dict: Processed statistics for the discipline including missing data details
    """
    result = {
        "Complete Data Points": 0,
        "Partial Data Points": 0,
        "Missing Data Points": 0,
        "Original Estimate (Total)": 0,
        "AI Estimate (Total)": 0,
        "Actual Time (Total)": 0,
        "Original vs Actual (%)": 0,
        "AI vs Actual (%)": 0,
        "Estimation Improvement (%)": 0,
        "Missing Data Details": []  # List of dicts with missing data info
    }

    # Create sets to track unique tickets
    complete_tickets = set()
    partial_tickets = set()
    missing_tickets = set()

    # First pass: categorize tickets and track missing data
    for _, row in df.iterrows():
        original_col = discipline_mapping["original"]
        ai_col = discipline_mapping["ai"]
        actual_col = discipline_mapping["actual"]

        # Consider a value present only if it's not null
        has_original = pd.notna(row[original_col])
        has_ai = pd.notna(row[ai_col])
        has_actual = pd.notna(row[actual_col])
        
        # Track any ticket that has at least one field populated
        if has_original or has_ai or has_actual:
            if has_original and has_ai and has_actual:
                complete_tickets.add(row["Key"])
            else:
                partial_tickets.add(row["Key"])
                # Track missing fields for this ticket
                missing_fields = []
                present_fields = []
                values = {}
                
                # Check original estimate
                if not has_original:
                    missing_fields.append("Original Estimate")
                else:
                    value = float(row[original_col])
                    present_fields.append("Original Estimate")
                    values["Original Estimate"] = value
                    
                # Check AI estimate
                if not has_ai:
                    missing_fields.append("AI Estimate")
                else:
                    value = float(row[ai_col])
                    present_fields.append("AI Estimate")
                    values["AI Estimate"] = value
                    
                # Check actual time
                if not has_actual:
                    missing_fields.append("Actual Time")
                else:
                    value = float(row[actual_col])
                    present_fields.append("Actual Time")
                    values["Actual Time"] = value
                
                # Only add to missing data details if we have at least one non-zero value
                if values:
                    result["Missing Data Details"].append({
                        "Ticket": row["Key"],
                        "Assignee": row.get("Assignee", "Unassigned"),
                        "Missing Fields": missing_fields,
                        "Present Fields": present_fields,
                        "Values": values
                    })
        else:
            missing_tickets.add(row["Key"])

    # Second pass: calculate totals only for complete tickets
    for _, row in df.iterrows():
        if row["Key"] in complete_tickets:
            original_col = discipline_mapping["original"]
            ai_col = discipline_mapping["ai"]
            actual_col = discipline_mapping["actual"]
            
            result["Original Estimate (Total)"] += float(row[original_col])
            result["AI Estimate (Total)"] += float(row[ai_col])
            result["Actual Time (Total)"] += float(row[actual_col])

    result["Complete Data Points"] = len(complete_tickets)
    result["Partial Data Points"] = len(partial_tickets)
    result["Missing Data Points"] = len(missing_tickets)

    # Only calculate improvements if we have complete data points
    if result["Complete Data Points"] > 0:
        improvements = calculate_improvement(
            result["Original Estimate (Total)"],
            result["AI Estimate (Total)"],
            result["Actual Time (Total)"]
        )
        result["Original vs Actual (%)"] = round(improvements["original_vs_actual"], 1)
        result["AI vs Actual (%)"] = round(improvements["ai_vs_actual"], 1)
        result["Estimation Improvement (%)"] = round(improvements["overall_improvement"], 1)

    return result
