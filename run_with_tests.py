"""
Test runner for StockWatch application.

This script runs all tests and only launches the application if all tests pass.
"""

import unittest
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

# Import test modules
from stockwatch.tests.test_data_processor import TestDataProcessor
from stockwatch.tests.test_ui_components import TestUIComponents
from stockwatch.tests.test_main_window import TestMainWindow

def run_tests():
    """Run all tests and return True if all tests pass, False otherwise."""
    # Create a test suite with all tests
    test_suite = unittest.TestSuite()
    test_suite.addTest(unittest.makeSuite(TestDataProcessor))
    test_suite.addTest(unittest.makeSuite(TestUIComponents))
    test_suite.addTest(unittest.makeSuite(TestMainWindow))
    
    # Run the tests
    test_runner = unittest.TextTestRunner(verbosity=2)
    result = test_runner.run(test_suite)
    
    # Return True if all tests pass, False otherwise
    return result.wasSuccessful()

def main():
    """Run all tests and launch the application if tests pass."""
    print("Running StockWatch tests...")
    
    if run_tests():
        print("\nAll tests passed! Launching application...\n")
        from main import main as app_main
        app_main()
    else:
        print("\nTests failed. Please fix the issues before running the application.")
        sys.exit(1)

if __name__ == "__main__":
    main()