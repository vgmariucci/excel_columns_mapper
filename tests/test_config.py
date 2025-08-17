"""
Test Configuration for ExcelColumnMapper
"""

import os
import tempfile
from pathlib import Path

class TestConfig:
    """Test-specific configuration settings"""
    
    # Test directories
    TEST_ROOT = Path(__file__).parent
    TEST_DATA_DIR = TEST_ROOT / "test_data"
    TEST_OUTPUT_DIR = TEST_ROOT / "test_output"
    
    # Test files
    SAMPLE_SOURCE_FILE = TEST_DATA_DIR / "sample_source.xlsx"
    SAMPLE_DESTINATION_FILE = TEST_DATA_DIR / "sample_destination.xlsx"
    SAMPLE_CSV_FILE = TEST_DATA_DIR / "sample_data.csv"
    
    # Mock settings
    MOCK_WINDOW_WIDTH = 1400
    MOCK_WINDOW_HEIGHT = 800
    MOCK_SCREEN_WIDTH = 1920
    MOCK_SCREEN_HEIGHT = 1080
    
    @classmethod
    def setup_test_data(cls):
        """Create test data files"""
        import pandas as pd
        
        # Ensure test directories exist
        cls.TEST_DATA_DIR.mkdir(exist_ok=True)
        cls.TEST_OUTPUT_DIR.mkdir(exist_ok=True)
        
        # Create sample source data
        source_data = pd.DataFrame({
            'Employee_ID': [1, 2, 3, 4, 5],
            'First_Name': ['John', 'Jane', 'Bob', 'Alice', 'Charlie'],
            'Last_Name': ['Doe', 'Smith', 'Johnson', 'Williams', 'Brown'],
            'Department': ['IT', 'HR', 'Finance', 'Marketing', 'IT'],
            'Salary': [50000, 60000, 55000, 65000, 58000],
            'Start_Date': ['2020-01-15', '2019-03-20', '2021-06-10', '2018-11-05', '2022-02-28']
        })
        
        # Create sample destination template
        destination_data = pd.DataFrame({
            'ID': [0] * 5,
            'Full_Name': [''] * 5,
            'Dept': [''] * 5,
            'Annual_Salary': [0] * 5,
            'Hire_Date': [''] * 5
        })
        
        # Save test files
        source_data.to_excel(cls.SAMPLE_SOURCE_FILE, index=False)
        destination_data.to_excel(cls.SAMPLE_DESTINATION_FILE, index=False)
        source_data.to_csv(cls.SAMPLE_CSV_FILE, index=False)
        
        return cls.SAMPLE_SOURCE_FILE, cls.SAMPLE_DESTINATION_FILE
    
    @classmethod
    def cleanup_test_data(cls):
        """Clean up test data files"""
        import shutil
        
        if cls.TEST_DATA_DIR.exists():
            shutil.rmtree(cls.TEST_DATA_DIR)
        if cls.TEST_OUTPUT_DIR.exists():
            shutil.rmtree(cls.TEST_OUTPUT_DIR)