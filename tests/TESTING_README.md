# ExcelColumnMapper Testing Documentation

## Overview

This document provides comprehensive testing guidelines for the ExcelColumnMapper application. The test suite focuses specifically on testing the main `ExcelColumnMapper` class functionality through unit tests and integration tests.

## Test Suite Structure

### Test Organization

The testing framework is organized into two main categories:

1. **Unit Tests (`TestExcelColumnMapper`)** - Test individual methods and functionality in isolation
2. **Integration Tests (`TestExcelColumnMapperIntegration`)** - Test components working together with real data

### Test Coverage Areas

#### Core Functionality
- Application initialization and setup
- Window configuration and management
- File loading operations (Excel and CSV)
- Column mapping logic and validation
- Data sampling and preview generation
- Theme management (light/dark mode switching)
- Status updates and progress tracking
- Error handling and edge cases

#### User Interface Components
- File browsing and selection
- Column mapping interface
- Progress bar functionality
- Status message updates
- Theme toggle operations

#### Data Processing
- Excel file reading and parsing
- CSV file handling
- Column header extraction
- Sample data generation
- Data type preservation

## Running the Tests

### Prerequisites

Ensure you have the following dependencies installed:

```bash
pip install pandas openpyxl unittest
```

Optional testing tools:
```bash
pip install pytest pytest-cov coverage
```

### Basic Test Execution

#### Method 1: Direct Python Execution
```bash
cd tests
python test_excel_column_mapper.py
```

#### Method 2: Using unittest Module
```bash
cd tests
python -m unittest test_excel_column_mapper -v
```

#### Method 3: Using pytest (if installed)
```bash
cd tests
pytest test_excel_column_mapper.py -v
```

### Running Specific Tests

#### Run only unit tests
```bash
python -m unittest test_excel_column_mapper.TestExcelColumnMapper -v
```

#### Run only integration tests
```bash
python -m unittest test_excel_column_mapper.TestExcelColumnMapperIntegration -v
```

#### Run a specific test method
```bash
python -m unittest test_excel_column_mapper.TestExcelColumnMapper.test_init -v
```

### Advanced Test Execution

#### With coverage reporting
```bash
coverage run test_excel_column_mapper.py
coverage report
coverage html  # Generates HTML coverage report
```

#### Using pytest with coverage
```bash
pytest test_excel_column_mapper.py --cov=main --cov-report=html --cov-report=term-missing
```

## Expected Test Output

### Successful Test Run

When all tests pass, you should see output similar to:

```
Starting ExcelColumnMapper Test Suite
============================================================
Successfully imported from main module

test_clear_mappings (__main__.TestExcelColumnMapper) ... ok
test_file_browsing (__main__.TestExcelColumnMapper) ... ok
test_get_sample_data (__main__.TestExcelColumnMapper) ... ok
test_get_sample_data_empty_column (__main__.TestExcelColumnMapper) ... ok
test_get_sample_data_nonexistent_column (__main__.TestExcelColumnMapper) ... ok
test_import_success (__main__.TestExcelColumnMapper) ... ok
test_init (__main__.TestExcelColumnMapper) ... ok
test_mapping_functionality (__main__.TestExcelColumnMapper) ... ok
test_read_excel_data_csv (__main__.TestExcelColumnMapper) ... ok
test_read_excel_data_xlsx (__main__.TestExcelColumnMapper) ... ok
test_setup_window (__main__.TestExcelColumnMapper) ... ok
test_status_update (__main__.TestExcelColumnMapper) ... ok
test_theme_functionality (__main__.TestExcelColumnMapper) ... ok
test_real_file_operations (__main__.TestExcelColumnMapperIntegration) ... ok

============================================================
TEST SUMMARY
============================================================
Tests run: 14
Failures: 0
Errors: 0
Success rate: 100.0%
```

### Test Failure Output

If tests fail, the output will include:

- **Failure count** and **error count**
- **Detailed error messages** for each failed test
- **Traceback information** for debugging
- **Success rate percentage**

## Test Descriptions

### Unit Tests

#### test_import_success
Verifies that all required classes can be imported successfully and have the expected attributes.

#### test_init
Tests the initialization of the ExcelColumnMapper class, ensuring:
- Root window assignment
- Variable initialization
- Manager instantiation
- Method call verification

#### test_setup_window
Validates window configuration including:
- Title setting
- Geometry configuration
- Minimum size constraints

#### test_center_window
Tests window centering calculations and geometry updates.

#### test_read_excel_data_xlsx / test_read_excel_data_csv
Verifies file reading functionality for different formats:
- Excel file parsing
- CSV file parsing
- Correct method selection based on file extension

#### test_get_sample_data
Tests data sampling functionality:
- Valid column sampling
- Non-existent column handling
- Empty column detection

#### test_theme_functionality
Validates theme management:
- Button text updates
- Theme state tracking
- Light/dark mode switching

#### test_file_browsing
Tests file dialog interactions and path handling.

#### test_status_update
Verifies status message and progress bar updates.

#### test_mapping_functionality
Tests column mapping operations:
- Mapping creation
- Mapping updates
- Statistics tracking

#### test_clear_mappings
Validates mapping reset functionality.

### Integration Tests

#### test_real_file_operations
Tests end-to-end functionality with actual Excel files:
- Real file reading
- Data integrity verification
- Column detection

## Test Architecture

### Mocking Strategy

The test suite uses comprehensive mocking to eliminate external dependencies:

#### GUI Components
- **tkinter.StringVar**: Replaced with MockStringVar to avoid root window dependency
- **tkinter.Tk**: Mocked to prevent GUI window creation
- **UI widgets**: All tkinter components are mocked

#### External Dependencies
- **File operations**: Real files used in integration tests, mocked in unit tests
- **pandas operations**: Mocked for isolated unit tests
- **Manager classes**: ThemeManager and StatisticsManager are mocked

#### System Interactions
- **File dialogs**: Mocked to return predictable test paths
- **File system operations**: Controlled through temporary directories

### MockStringVar Class

```python
class MockStringVar:
    """Mock StringVar that doesn't require tkinter root"""
    def __init__(self, value=''):
        self._value = value
    
    def get(self):
        return self._value
    
    def set(self, value):
        self._value = value
```

This class provides the same interface as `tkinter.StringVar` without requiring tkinter initialization.

## Test Data Management

### Sample Data

The test suite uses predefined pandas DataFrames for consistent testing:

```python
source_data = pd.DataFrame({
    'Name': ['John', 'Jane', 'Bob'],
    'Age': [25, 30, 35],
    'City': ['New York', 'London', 'Paris']
})

destination_data = pd.DataFrame({
    'Full_Name': ['', '', ''],
    'Person_Age': [0, 0, 0],
    'Location': ['', '', '']
})
```

### Temporary Files

Integration tests create temporary Excel files for real file operations:
- Created in system temporary directory
- Automatically cleaned up after tests
- Contains realistic data structures

## Troubleshooting

### Common Issues and Solutions

#### Import Errors
**Problem**: `ModuleNotFoundError: No module named 'app'`

**Solution**: The test file automatically handles import path issues. Ensure you're running from the tests directory and that `../app/main.py` exists.

#### Tkinter Errors
**Problem**: `RuntimeError: Too early to create variable: no default root window`

**Solution**: This is resolved by the MockStringVar implementation. Ensure you're using the fixed test file.

#### File Permission Errors
**Problem**: Cannot create temporary files or access test data

**Solution**: 
- Ensure write permissions in the tests directory
- Check available disk space
- Verify antivirus software isn't blocking file operations

#### Missing Dependencies
**Problem**: Import errors for pandas, openpyxl, or other modules

**Solution**:
```bash
pip install pandas openpyxl pathlib unittest
```

### Debug Mode

For additional debugging information, modify the test file to enable verbose output:

```python
# Add at the top of the test file
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Test Isolation

Each test method is isolated through:
- **setUp()** method creating fresh fixtures
- **tearDown()** method cleaning up resources
- **Mock patching** preventing side effects
- **Temporary directories** for file operations

## Performance Considerations

### Test Execution Time

- **Unit tests**: Typically complete in under 5 seconds
- **Integration tests**: May take 10-15 seconds due to file I/O
- **Full suite**: Usually completes within 30 seconds

### Memory Usage

- **Peak memory**: Approximately 50-100MB during test execution
- **Temporary files**: Cleaned up automatically
- **Mock objects**: Garbage collected after each test

### Optimization Tips

1. **Run unit tests first** for faster feedback during development
2. **Use pytest markers** to selectively run test categories
3. **Parallel execution** with pytest-xdist for larger test suites
4. **Coverage reporting** only when needed to reduce overhead

## Continuous Integration

### GitHub Actions Example

```yaml
name: Test ExcelColumnMapper

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        python-version: [3.10, 3.11, 3.12]

    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v3
      with:
        python-version: ${{ matrix.python-version }}
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pandas openpyxl pytest coverage
    
    - name: Run tests
      run: |
        cd tests
        python test_excel_column_mapper.py
```

### Local Pre-commit Hooks

```bash
# .git/hooks/pre-commit
#!/bin/sh
cd tests
python test_excel_column_mapper.py
if [ $? -ne 0 ]; then
    echo "Tests failed. Commit aborted."
    exit 1
fi
```

## Test Metrics and Quality

### Target Metrics

- **Code Coverage**: Greater than 90%
- **Test Success Rate**: 100%
- **Execution Time**: Less than 30 seconds
- **Memory Usage**: Less than 100MB

### Quality Indicators

- **All tests pass consistently**
- **No skipped tests** (unless marked as expected)
- **Clear test names** describing functionality
- **Proper assertion messages** for debugging
- **Comprehensive edge case coverage**

## Best Practices

### Test Naming Convention

- Test files: `test_*.py`
- Test classes: `Test*`
- Test methods: `test_*_*` (descriptive names)

### Test Structure

```python
def test_specific_functionality(self):
    """Clear description of what is being tested"""
    # Arrange: Set up test data and mocks
    
    # Act: Execute the functionality being tested
    
    # Assert: Verify the expected behavior
```

### Assertion Guidelines

- Use **specific assertions** rather than general ones
- Include **meaningful error messages** in assertions
- Test both **positive and negative cases**
- Verify **side effects** and **state changes**

### Mock Usage

- **Mock external dependencies** completely
- **Verify mock calls** to ensure proper interaction
- **Use realistic return values** for mocks
- **Clean up mocks** between tests

## Adding New Tests

### Test Template

```python
def test_new_functionality(self, mock_root):
    """Test description of new functionality"""
    with patch.object(ExcelColumnMapper, 'create_widgets'), \
         patch('main.ThemeManager'), \
         patch('main.StatisticsManager'), \
         patch('main.Config.ensure_directories'):
        
        # Arrange
        mapper = ExcelColumnMapper(mock_root)
        # Set up test-specific data and mocks
        
        # Act
        result = mapper.method_being_tested()
        
        # Assert
        assert result == expected_value
        # Verify side effects and method calls
```

### Integration Test Template

```python
def test_integration_scenario(self, mock_root, temp_excel_files):
    """Test end-to-end scenario description"""
    source_file, dest_file = temp_excel_files
    
    with patch.object(ExcelColumnMapper, 'create_widgets'), \
         # ... other patches
        
        mapper = ExcelColumnMapper(mock_root)
        
        # Test real file operations
        result = mapper.process_files(source_file, dest_file)
        
        # Verify results with real data
        assert len(result) > 0
        assert result.columns.tolist() == expected_columns
```

## Maintenance and Updates

### Regular Maintenance Tasks

1. **Update test data** when application data structures change
2. **Review mock implementations** for accuracy with real components
3. **Add tests for new features** as they are developed
4. **Remove obsolete tests** for deprecated functionality
5. **Update documentation** to reflect test changes

### Version Compatibility

The test suite is designed to work with:
- **Python 3.10+**
- **pandas 1.5+**
- **openpyxl 3.0+**
- **unittest** (built-in)
- **pytest 6.0+** (optional)

### Backward Compatibility

When updating tests:
- **Maintain existing test interfaces**
- **Add deprecation warnings** for old test methods
- **Provide migration guides** for breaking changes
- **Support multiple pandas versions** where possible