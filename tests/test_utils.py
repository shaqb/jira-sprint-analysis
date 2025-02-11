"""Tests for the utils module."""

import unittest
import os
import tempfile
from jira_analysis.utils import ensure_directory, find_input_file

class TestUtils(unittest.TestCase):
    def setUp(self):
        """Set up test environment."""
        self.temp_dir = tempfile.mkdtemp()
        self.input_dir = os.path.join(self.temp_dir, "input")

    def tearDown(self):
        """Clean up test environment."""
        import shutil
        shutil.rmtree(self.temp_dir)

    def test_ensure_directory(self):
        """Test directory creation."""
        test_dir = os.path.join(self.temp_dir, "test_dir")
        self.assertFalse(os.path.exists(test_dir))
        
        ensure_directory(test_dir)
        self.assertTrue(os.path.exists(test_dir))
        
        # Test idempotency
        ensure_directory(test_dir)
        self.assertTrue(os.path.exists(test_dir))

    def test_find_input_file(self):
        """Test input file finding."""
        # Create input directory
        os.makedirs(self.input_dir)
        
        # Test with no files
        with self.assertRaises(FileNotFoundError):
            find_input_file()
        
        # Create test files
        test_files = [
            ("old.xlsx", 0),
            ("new.xlsx", 1),
            ("newest.xlsx", 2)
        ]
        
        for filename, delay in test_files:
            filepath = os.path.join(self.input_dir, filename)
            with open(filepath, "w") as f:
                f.write("test")
            # Set access and modify times
            os.utime(filepath, (delay, delay))
        
        # Test finding most recent file
        latest_file = find_input_file()
        self.assertEqual(os.path.basename(latest_file), "newest.xlsx")

if __name__ == "__main__":
    unittest.main()
