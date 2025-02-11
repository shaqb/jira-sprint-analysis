"""Tests for the analyzer module."""

import unittest
import pandas as pd
from jira_analysis.analyzer import process_discipline_data, calculate_percentage_diff

class TestAnalyzer(unittest.TestCase):
    def setUp(self):
        """Set up test data."""
        self.test_data = pd.DataFrame({
            "Key": ["TEST-1", "TEST-2", "TEST-3"],
            "QA Original Estimate based on old way": [2.0, 3.0, None],
            "QA Revised Estimate based on AI": [1.5, 2.5, None],
            "QA Actual Time based on AI": [1.8, 2.8, None],
        })
        
        self.discipline_mapping = {
            "original": "QA Original Estimate based on old way",
            "ai": "QA Revised Estimate based on AI",
            "actual": "QA Actual Time based on AI",
        }

    def test_calculate_percentage_diff(self):
        """Test percentage difference calculation."""
        self.assertEqual(calculate_percentage_diff(100, 120), 20)
        self.assertEqual(calculate_percentage_diff(100, 80), -20)
        self.assertEqual(calculate_percentage_diff(0, 100), 0)
        self.assertEqual(calculate_percentage_diff(100, None), 0)
        self.assertEqual(calculate_percentage_diff(None, 100), 0)

    def test_process_discipline_data(self):
        """Test discipline data processing."""
        result = process_discipline_data(self.test_data, self.discipline_mapping)

        self.assertEqual(result["Complete Data Points"], 2)
        self.assertEqual(result["Partial Data Points"], 0)
        self.assertEqual(result["Missing Data Points"], 1)
        self.assertEqual(result["Original Estimate (Total)"], 5.0)
        self.assertEqual(result["AI Estimate (Total)"], 4.0)
        self.assertEqual(result["Actual Time (Total)"], 4.6)

    def test_process_discipline_data_missing(self):
        """Test processing with missing data."""
        test_data = pd.DataFrame({
            "Key": ["TEST-1"],
            "QA Original Estimate based on old way": [2.0],
            "QA Revised Estimate based on AI": [None],
            "QA Actual Time based on AI": [1.8],
        })

        result = process_discipline_data(test_data, self.discipline_mapping)
        
        self.assertEqual(result["Complete Data Points"], 0)
        self.assertEqual(result["Partial Data Points"], 1)
        self.assertEqual(result["Missing Data Points"], 0)
        self.assertEqual(result["Original Estimate (Total)"], 2.0)
        self.assertEqual(result["AI Estimate (Total)"], 0)
        self.assertEqual(result["Actual Time (Total)"], 1.8)

if __name__ == "__main__":
    unittest.main()
