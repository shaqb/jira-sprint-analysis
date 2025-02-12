"""Tests for the analyzer module."""

import unittest
import pandas as pd
from jira_analysis.analyzer import (
    process_discipline_data,
    calculate_percentage_diff,
    calculate_improvement
)

class TestAnalyzer(unittest.TestCase):
    def setUp(self):
        """Set up test data."""
        self.test_data = pd.DataFrame({
            "Key": ["TEST-1", "TEST-2", "TEST-3", "TEST-4"],
            "BE Dev Original Estimate based on old way": [4.0, 6.0, 3.0, None],
            "BE Dev Revised Estimate based on AI": [3.5, 5.5, None, 2.0],
            "BE Dev Actual Time based on AI": [3.8, 5.8, 2.8, 1.8],
        })
        
        self.discipline_mapping = {
            "original": "BE Dev Original Estimate based on old way",
            "ai": "BE Dev Revised Estimate based on AI",
            "actual": "BE Dev Actual Time based on AI",
        }

    def test_calculate_percentage_diff(self):
        """Test percentage difference calculation."""
        self.assertEqual(calculate_percentage_diff(100, 120), 20)
        self.assertEqual(calculate_percentage_diff(100, 80), -20)
        self.assertEqual(calculate_percentage_diff(0, 100), 0)
        self.assertEqual(calculate_percentage_diff(100, None), 0)
        self.assertEqual(calculate_percentage_diff(None, 100), 0)

    def test_calculate_improvement(self):
        """Test improvement calculation."""
        # Test case where AI estimate is better than original
        result = calculate_improvement(
            original_total=10.0,  # Original overestimated by 25%
            ai_total=8.5,        # AI overestimated by 6.25%
            actual_total=8.0
        )
        self.assertAlmostEqual(result["original_vs_actual"], 25.0, places=1)
        self.assertAlmostEqual(result["ai_vs_actual"], 6.25, places=1)
        self.assertAlmostEqual(result["overall_improvement"], 75.0, places=1)

        # Test case where both estimates are under actual
        result = calculate_improvement(
            original_total=8.0,   # Original underestimated by 20%
            ai_total=9.0,        # AI underestimated by 10%
            actual_total=10.0
        )
        self.assertAlmostEqual(result["original_vs_actual"], -20.0, places=1)
        self.assertAlmostEqual(result["ai_vs_actual"], -10.0, places=1)
        self.assertAlmostEqual(result["overall_improvement"], 50.0, places=1)

        # Test case with perfect AI estimate
        result = calculate_improvement(
            original_total=12.0,  # Original overestimated by 20%
            ai_total=10.0,       # AI estimate perfect
            actual_total=10.0
        )
        self.assertAlmostEqual(result["original_vs_actual"], 20.0, places=1)
        self.assertAlmostEqual(result["ai_vs_actual"], 0.0, places=1)
        self.assertAlmostEqual(result["overall_improvement"], 100.0, places=1)

    def test_process_discipline_data(self):
        """Test discipline data processing with improvements."""
        result = process_discipline_data(self.test_data, self.discipline_mapping)

        # Should only count tickets with all three values
        self.assertEqual(result["Complete Data Points"], 2)  # Only TEST-1 and TEST-2 have all values
        self.assertEqual(result["Partial Data Points"], 2)  # TEST-3 and TEST-4 are partial
        self.assertEqual(result["Missing Data Points"], 0)

        # Totals should only include complete tickets
        self.assertEqual(result["Original Estimate (Total)"], 10.0)  # 4.0 + 6.0
        self.assertEqual(result["AI Estimate (Total)"], 9.0)        # 3.5 + 5.5
        self.assertEqual(result["Actual Time (Total)"], 9.6)        # 3.8 + 5.8

        # Verify improvements are calculated only from complete data
        self.assertIsInstance(result["Original vs Actual (%)"], float)
        self.assertIsInstance(result["AI vs Actual (%)"], float)
        self.assertIsInstance(result["Estimation Improvement (%)"], float)
        self.assertLessEqual(result["Estimation Improvement (%)"], 100.0)

    def test_process_discipline_data_no_complete(self):
        """Test processing with no complete data points."""
        test_data = pd.DataFrame({
            "Key": ["TEST-1", "TEST-2"],
            "BE Dev Original Estimate based on old way": [4.0, None],
            "BE Dev Revised Estimate based on AI": [None, 3.5],
            "BE Dev Actual Time based on AI": [3.8, 3.8],
        })

        result = process_discipline_data(test_data, self.discipline_mapping)
        
        # Should identify all tickets as partial since none have all three values
        self.assertEqual(result["Complete Data Points"], 0)
        self.assertEqual(result["Partial Data Points"], 2)
        self.assertEqual(result["Missing Data Points"], 0)

        # Totals should be zero since there are no complete tickets
        self.assertEqual(result["Original Estimate (Total)"], 0)
        self.assertEqual(result["AI Estimate (Total)"], 0)
        self.assertEqual(result["Actual Time (Total)"], 0)
        
        # No improvements should be calculated
        self.assertEqual(result["Original vs Actual (%)"], 0)
        self.assertEqual(result["AI vs Actual (%)"], 0)
        self.assertEqual(result["Estimation Improvement (%)"], 0)

if __name__ == "__main__":
    unittest.main()
