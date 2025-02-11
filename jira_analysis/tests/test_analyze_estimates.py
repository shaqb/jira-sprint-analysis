import unittest

import pandas as pd

from jira_analysis.analyze_estimates import process_discipline_data


class TestAnalyzeEstimates(unittest.TestCase):
    def setUp(self):
        # Create sample test data
        self.test_data = pd.DataFrame(
            {
                "Key": ["TEST-1", "TEST-2"],
                "QA Original Estimate based on old way": [2.0, 3.0],
                "QA Revised Estimate based on AI": [1.5, 2.5],
                "QA Actual Time based on AI": [1.8, 2.8],
            }
        )

    def test_process_discipline_data(self):
        # Test processing of QA discipline data
        discipline_mapping = {
            "original": "QA Original Estimate based on old way",
            "ai": "QA Revised Estimate based on AI",
            "actual": "QA Actual Time based on AI",
        }

        result = process_discipline_data(self.test_data, discipline_mapping)

        # Assert expected results
        self.assertEqual(result["Complete Data Points"], 2)
        self.assertEqual(result["Missing Data Points"], 0)
        self.assertEqual(result["Original Estimate (Total)"], 5.0)
        self.assertEqual(result["AI Estimate (Total)"], 4.0)
        self.assertEqual(result["Actual Time (Total)"], 4.6)


if __name__ == "__main__":
    unittest.main()
