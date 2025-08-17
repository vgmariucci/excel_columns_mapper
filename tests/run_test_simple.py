#!/usr/bin/env python3
"""
Simple Test Runner for ExcelColumnMapper
Handles import path issues automatically
"""

import os
import sys
from pathlib import Path

def setup_test_environment():
    """Setup the test environment and paths"""
    print("🔧 Setting up test environment...")
    
    # Get current script directory
    current_dir = Path(__file__).parent
    
    # Get project root (parent of tests directory)
    project_root = current_dir.parent
    
    # Get app directory
    app_dir = project_root / "app"
    
    print(f"Current directory: {current_dir}")
    print(f"Project root: {project_root}")
    print(f"App directory: {app_dir}")
    
    # Add paths to sys.path
    paths_to_add = [str(project_root), str(app_dir)]
    
    for path in paths_to_add:
        if path not in sys.path:
            sys.path.insert(0, path)
            print(f"✅ Added to Python path: {path}")
    
    # Verify app directory exists
    if not app_dir.exists():
        print(f"❌ App directory not found: {app_dir}")
        return False
    
    # Verify main.py exists
    main_py = app_dir / "main.py"
    if not main_py.exists():
        print(f"❌ main.py not found: {main_py}")
        return False
    
    print("✅ Environment setup complete")
    return True

def run_tests():
    """Run the tests"""
    print("\n🧪 Running ExcelColumnMapper tests...")
    
    try:
        # Import the test module
        from test_excel_column_mapper import run_tests as run_test_suite
        
        # Run the tests
        success = run_test_suite()
        
        if success:
            print("\n🎉 All tests passed!")
        else:
            print("\n❌ Some tests failed!")
        
        return success
        
    except ImportError as e:
        print(f"❌ Failed to import test module: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure you're running this from the tests/ directory")
        print("2. Ensure the corrected test_excel_column_mapper.py file is in the tests/ directory")
        print("3. Verify that app/main.py exists in the parent directory")
        return False
    
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def main():
    """Main entry point"""
    print("🚀 ExcelColumnMapper Test Runner")
    print("=" * 50)
    
    # Setup environment
    if not setup_test_environment():
        print("❌ Failed to setup test environment")
        sys.exit(1)
    
    # Run tests
    success = run_tests()
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()