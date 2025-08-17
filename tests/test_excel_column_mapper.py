"""
Unit Tests for ExcelColumnMapper Application
Test suite for the main ExcelColumnMapper class functionality
Fixed version that handles tkinter initialization issues
"""

import unittest
from unittest.mock import Mock, MagicMock, patch, mock_open
import tkinter as tk
from tkinter import ttk
import pandas as pd
import tempfile
import os
import json
from pathlib import Path
import sys

# Fix import path issue
def setup_imports():
    """Setup import paths for testing"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    app_dir = os.path.join(project_root, 'app')
    
    if project_root not in sys.path:
        sys.path.insert(0, project_root)
    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)

setup_imports()

try:
    from main import ExcelColumnMapper, Config, ThemeManager, StatisticsManager
    IMPORT_SUCCESS = True
    print("✅ Successfully imported from main module")
except ImportError as e:
    print(f"❌ Failed to import from main: {e}")
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
        from app.main import ExcelColumnMapper, Config, ThemeManager, StatisticsManager
        IMPORT_SUCCESS = True
        print("✅ Successfully imported from app.main module")
    except ImportError as e2:
        print(f"❌ Failed to import from app.main: {e2}")
        IMPORT_SUCCESS = False


class MockStringVar:
    """Mock StringVar that doesn't require tkinter root"""
    def __init__(self, value=''):
        self._value = value
    
    def get(self):
        return self._value
    
    def set(self, value):
        self._value = value


class TestExcelColumnMapper(unittest.TestCase):
    """Test suite for ExcelColumnMapper class"""
    
    @classmethod
    def setUpClass(cls):
        """Set up class-level fixtures"""
        if not IMPORT_SUCCESS:
            raise unittest.SkipTest("Could not import required modules")
    
    def setUp(self):
        """Set up test fixtures before each test method"""
        # Create a mock root window
        self.mock_root = Mock(spec=tk.Tk)
        self.mock_root.title = Mock()
        self.mock_root.geometry = Mock()
        self.mock_root.minsize = Mock()
        self.mock_root.update_idletasks = Mock()
        self.mock_root.winfo_width.return_value = 1400
        self.mock_root.winfo_height.return_value = 800
        self.mock_root.winfo_screenwidth.return_value = 1920
        self.mock_root.winfo_screenheight.return_value = 1080
        self.mock_root.after = Mock()
        
        # Create test data
        self.sample_source_data = pd.DataFrame({
            'Name': ['John', 'Jane', 'Bob'],
            'Age': [25, 30, 35],
            'City': ['New York', 'London', 'Paris']
        })
        
        self.sample_destination_data = pd.DataFrame({
            'Full_Name': ['', '', ''],
            'Person_Age': [0, 0, 0],
            'Location': ['', '', '']
        })
        
        # Create temporary directory for testing
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up after each test method"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_import_success(self):
        """Test that all required classes can be imported"""
        self.assertTrue(IMPORT_SUCCESS, "Failed to import required classes")
        self.assertTrue(hasattr(ExcelColumnMapper, '__init__'))
        self.assertTrue(hasattr(Config, 'ensure_directories'))
        self.assertTrue(hasattr(ThemeManager, '__init__'))
        self.assertTrue(hasattr(StatisticsManager, '__init__'))
    
    @patch('tkinter.StringVar', MockStringVar)
    @patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories')
    @patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager')
    @patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager')
    def test_init(self, mock_stats_manager, mock_theme_manager, mock_ensure_dirs):
        """Test ExcelColumnMapper initialization"""
        # Mock the managers
        mock_theme_instance = Mock()
        mock_stats_instance = Mock()
        mock_theme_manager.return_value = mock_theme_instance
        mock_stats_manager.return_value = mock_stats_instance
        
        with patch.object(ExcelColumnMapper, 'setup_window'), \
             patch.object(ExcelColumnMapper, 'create_widgets'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Verify initialization
            self.assertEqual(mapper.root, self.mock_root)
            self.assertIsInstance(mapper.source_file_path, MockStringVar)
            self.assertIsInstance(mapper.destination_file_path, MockStringVar)
            self.assertEqual(mapper.source_headers, [])
            self.assertEqual(mapper.destination_headers, [])
            self.assertEqual(mapper.column_mappings, {})
            self.assertEqual(mapper.mapping_combos, {})
            
            # Verify method calls
            mock_theme_manager.assert_called_once_with(self.mock_root)
            mock_stats_manager.assert_called_once()
            mock_ensure_dirs.assert_called_once()
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_setup_window(self):
        """Test window setup configuration"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch.object(ExcelColumnMapper, 'center_window'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Verify window configuration calls
            self.mock_root.title.assert_called_with(Config.WINDOW_TITLE)
            self.mock_root.geometry.assert_called_with(Config.WINDOW_SIZE)
            self.mock_root.minsize.assert_called_with(*Config.WINDOW_MIN_SIZE)
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_center_window(self):
        """Test window centering functionality"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Test center_window method
            mapper.center_window()
            
            # Verify update_idletasks was called
            self.mock_root.update_idletasks.assert_called()
            
            # Verify geometry setting (calculating center position)
            expected_x = (1920 // 2) - (1400 // 2)
            expected_y = (1080 // 2) - (800 // 2)
            expected_geometry = f"1400x800+{expected_x}+{expected_y}"
            self.mock_root.geometry.assert_called_with(expected_geometry)
    
    @patch('tkinter.StringVar', MockStringVar)
    @patch('pandas.read_csv')
    @patch('pandas.read_excel')
    def test_read_excel_data_xlsx(self, mock_read_excel, mock_read_csv):
        """Test reading Excel (.xlsx) files"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Mock return value
            mock_read_excel.return_value = self.sample_source_data
            
            # Test Excel file reading
            result = mapper.read_excel_data('test_file.xlsx')
            
            # Verify correct method was called
            mock_read_excel.assert_called_once_with(Path('test_file.xlsx'))
            mock_read_csv.assert_not_called()
            self.assertEqual(result.equals(self.sample_source_data), True)
    
    @patch('tkinter.StringVar', MockStringVar)
    @patch('pandas.read_csv')
    @patch('pandas.read_excel')
    def test_read_excel_data_csv(self, mock_read_excel, mock_read_csv):
        """Test reading CSV files"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Mock return value
            mock_read_csv.return_value = self.sample_source_data
            
            # Test CSV file reading
            result = mapper.read_excel_data('test_file.csv')
            
            # Verify correct method was called
            mock_read_csv.assert_called_once_with(Path('test_file.csv'), encoding='utf-8')
            mock_read_excel.assert_not_called()
            self.assertEqual(result.equals(self.sample_source_data), True)
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_get_sample_data(self):
        """Test sample data extraction from columns"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Set up test data
            mapper.source_df = self.sample_source_data
            
            # Test sample data extraction
            sample = mapper.get_sample_data('Name', max_samples=2)
            expected_values = ['John', 'Jane']
            
            # Verify sample contains expected values
            for value in expected_values:
                self.assertIn(value, sample)
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_get_sample_data_nonexistent_column(self):
        """Test sample data extraction for non-existent column"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Set up test data
            mapper.source_df = self.sample_source_data
            
            # Test non-existent column
            sample = mapper.get_sample_data('NonExistentColumn')
            self.assertEqual(sample, "No data")
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_get_sample_data_empty_column(self):
        """Test sample data extraction for empty column"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Create DataFrame with empty column
            empty_data = pd.DataFrame({'Empty': [None, None, None]})
            mapper.source_df = empty_data
            
            # Test empty column
            sample = mapper.get_sample_data('Empty')
            self.assertEqual(sample, "All empty")
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_theme_functionality(self):
        """Test theme-related functionality"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Mock theme manager
            mapper.theme_manager = Mock()
            mapper.theme_manager.current_theme = "light"
            
            # Test get_theme_button_text for light mode
            text = mapper.get_theme_button_text()
            self.assertEqual(text, "🌙 Dark Mode")
            
            # Test get_theme_button_text for dark mode
            mapper.theme_manager.current_theme = "dark"
            text = mapper.get_theme_button_text()
            self.assertEqual(text, "☀ Light Mode")
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_file_browsing(self):
        """Test file browsing functionality"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'), \
             patch('tkinter.filedialog.askopenfilename') as mock_filedialog:
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Mock file dialog return
            test_file_path = '/path/to/test_file.xlsx'
            mock_filedialog.return_value = test_file_path
            
            # Mock update_status method
            mapper.update_status = Mock()
            
            # Test source file browsing
            mapper.browse_source_file()
            
            # Verify file path was set
            self.assertEqual(mapper.source_file_path.get(), test_file_path)
            mapper.update_status.assert_called_with('Source file selected: test_file.xlsx')
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_status_update(self):
        """Test status update functionality"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Mock status components
            mapper.status_var = Mock()
            mapper.progress_var = Mock()
            
            # Test status update with progress
            mapper.update_status("Test message", 50.0)
            
            # Verify status and progress were set
            mapper.status_var.set.assert_called_with("Test message")
            mapper.progress_var.set.assert_called_with(50.0)
            self.mock_root.update_idletasks.assert_called()
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_mapping_functionality(self):
        """Test column mapping functionality"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Set up test data
            mapper.source_headers = ['Name', 'Age', 'City']
            mapper.destination_headers = ['Full_Name', 'Person_Age', 'Location']
            
            # Mock combo box
            mock_combo = Mock()
            mock_combo.get.return_value = 'Name'
            mapper.mapping_combos = {'Full_Name': mock_combo}
            
            # Mock methods
            mapper.update_source_tree_mapping = Mock()
            mapper.update_status = Mock()
            mapper.update_preview = Mock()
            mapper.stats_manager = Mock()
            
            # Test mapping change
            mapper.on_mapping_changed('Full_Name')
            
            # Verify mapping was created
            self.assertEqual(mapper.column_mappings['Full_Name'], 'Name')
            mapper.update_source_tree_mapping.assert_called_with('Name', 'Full_Name', True)
            mapper.stats_manager.update_mapping_created.assert_called_once()
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_clear_mappings(self):
        """Test clearing all mappings"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Set up test data
            mapper.source_headers = ['Name', 'Age']
            mapper.column_mappings = {'Full_Name': 'Name', 'Person_Age': 'Age'}
            
            # Mock combo boxes
            mock_combo1 = Mock()
            mock_combo2 = Mock()
            mapper.mapping_combos = {'Full_Name': mock_combo1, 'Person_Age': mock_combo2}
            
            # Mock methods
            mapper.update_source_tree_mapping = Mock()
            mapper.update_status = Mock()
            mapper.update_preview = Mock()
            
            # Test clearing mappings
            mapper.clear_mappings()
            
            # Verify all mappings were cleared
            self.assertEqual(mapper.column_mappings, {})
            mock_combo1.set.assert_called_with('-- Select Source Column --')
            mock_combo2.set.assert_called_with('-- Select Source Column --')
            mapper.update_status.assert_called_with('All mappings cleared')


class TestExcelColumnMapperIntegration(unittest.TestCase):
    """Integration tests for ExcelColumnMapper with real data"""
    
    @classmethod
    def setUpClass(cls):
        """Set up class-level fixtures"""
        if not IMPORT_SUCCESS:
            raise unittest.SkipTest("Could not import required modules")
    
    def setUp(self):
        """Set up integration test fixtures"""
        # Create temporary test files
        self.temp_dir = tempfile.mkdtemp()
        
        # Create test Excel files
        self.source_file = os.path.join(self.temp_dir, 'source.xlsx')
        self.destination_file = os.path.join(self.temp_dir, 'destination.xlsx')
        
        # Create test data
        source_data = pd.DataFrame({
            'Employee_Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'Employee_Age': [25, 30, 35],
            'Department': ['IT', 'HR', 'Finance']
        })
        
        destination_data = pd.DataFrame({
            'Full_Name': ['', '', ''],
            'Age': [0, 0, 0],
            'Dept': ['', '', '']
        })
        
        # Save test files
        source_data.to_excel(self.source_file, index=False)
        destination_data.to_excel(self.destination_file, index=False)
        
        # Mock root window
        self.mock_root = Mock(spec=tk.Tk)
        self.mock_root.title = Mock()
        self.mock_root.geometry = Mock()
        self.mock_root.minsize = Mock()
        self.mock_root.update_idletasks = Mock()
        self.mock_root.winfo_width.return_value = 1400
        self.mock_root.winfo_height.return_value = 800
        self.mock_root.winfo_screenwidth.return_value = 1920
        self.mock_root.winfo_screenheight.return_value = 1080
        self.mock_root.after = Mock()
    
    def tearDown(self):
        """Clean up integration test fixtures"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @patch('tkinter.StringVar', MockStringVar)
    def test_real_file_operations(self):
        """Test operations with real Excel files"""
        with patch.object(ExcelColumnMapper, 'create_widgets'), \
             patch.object(ExcelColumnMapper, 'populate_source_tree'), \
             patch.object(ExcelColumnMapper, 'create_mapping_widgets'), \
             patch.object(ExcelColumnMapper, 'update_preview'), \
             patch('main.ThemeManager' if 'main' in sys.modules else 'app.main.ThemeManager'), \
             patch('main.StatisticsManager' if 'main' in sys.modules else 'app.main.StatisticsManager'), \
             patch('main.Config.ensure_directories' if 'main' in sys.modules else 'app.main.Config.ensure_directories'):
            
            mapper = ExcelColumnMapper(self.mock_root)
            
            # Test reading real Excel files
            source_df = mapper.read_excel_data(self.source_file)
            dest_df = mapper.read_excel_data(self.destination_file)
            
            # Verify data was loaded correctly
            self.assertEqual(len(source_df.columns), 3)
            self.assertEqual(len(dest_df.columns), 3)
            self.assertIn('Employee_Name', source_df.columns)
            self.assertIn('Full_Name', dest_df.columns)


def run_tests():
    """Run all tests and generate report"""
    print("🚀 Starting ExcelColumnMapper Test Suite")
    print("=" * 60)
    
    if not IMPORT_SUCCESS:
        print("❌ Failed to import required modules. Please check your Python path and module structure.")
        return False
    
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add test classes
    suite.addTests(loader.loadTestsFromTestCase(TestExcelColumnMapper))
    suite.addTests(loader.loadTestsFromTestCase(TestExcelColumnMapperIntegration))
    
    # Run tests with detailed output
    runner = unittest.TextTestRunner(verbosity=2, buffer=True)
    result = runner.run(suite)
    
    # Print summary
    print(f"\n{'='*60}")
    print(f"TEST SUMMARY")
    print(f"{'='*60}")
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.testsRun > 0:
        success_rate = ((result.testsRun - len(result.failures) - len(result.errors)) / result.testsRun * 100)
        print(f"Success rate: {success_rate:.1f}%")
    
    if result.failures:
        print(f"\nFAILURES:")
        for test, traceback in result.failures:
            print(f"- {test}")
            print(f"  {traceback.split('AssertionError:')[-1].strip() if 'AssertionError:' in traceback else 'See details above'}")
    
    if result.errors:
        print(f"\nERRORS:")
        for test, traceback in result.errors:
            print(f"- {test}")
            print(f"  {traceback.split('RuntimeError:')[-1].strip() if 'RuntimeError:' in traceback else 'See details above'}")
    
    return result.wasSuccessful()


if __name__ == '__main__':
    # Run the tests
    success = run_tests()
    sys.exit(0 if success else 1)