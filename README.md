# Excel Column Mapper & Data Transfer Tool

A user-friendly Python GUI application for mapping and transferring data between Excel files. Built with Azure-inspired theme and comprehensive functionality for seamless data migration.

![Version](https://img.shields.io/badge/version-1.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Python](https://img.shields.io/badge/python-3.12+-blue.svg)


## ✨ Features

### 🎨 **Modern User Interface**
- Beautiful Azure-inspired theme with light/dark mode support
- Responsive design with modern styling
- Intuitive tabbed interface for organized workflow
- Real-time progress tracking with visual feedback

### 🔄 **Advanced Column Mapping**
- Interactive drag-and-drop style column mapping interface
- Visual preview of source data with sample values
- Real-time mapping validation and preview
- Smart column suggestion and matching

### 📊 **Data Management**
- Support for Excel (.xlsx, .xls) and CSV files
- Handles files with different row counts intelligently
- Preserves data types during transfer
- Comprehensive error handling and validation

### 📈 **Statistics & Analytics**
- Built-in usage statistics tracking
- File processing counters
- Mapping creation analytics
- Session completion tracking

### 📝 **Mapping History**
- Automatic saving of all mapping configurations
- Load and reuse previous successful mappings
- Complete history browser with timestamps
- Compatibility checking for historical mappings

### 🌓 **Theme System**
- Light and dark mode support
- Azure-inspired modern design
- Customizable UI components
- Smooth theme transitions

### **Screenshots**

##### **Light Mode**
![app_gui_light_mode](assets/app_screenshots/app_gui_light_mode.png?raw=true)

##### **Dark Mode**
![app_gui_dark_mode](assets/app_screenshots/app_gui_dark_mode.png?raw=true)

##### **Uploading Source and Destination Template Files**
![app_gui_dark_mode_in_use](assets/app_screenshots/app_gui_dark_mode_in_use.png?raw=true)

##### **Columns Mapping With Dropdown Widgets**
![app_gui_dark_mode_column_mapping](assets/app_screenshots/app_gui_dark_mode_column_mapping.png?raw=true)

##### **Mapping From History File**
![app_gui_dark_mode_mapping_from_history](assets/app_screenshots/app_gui_dark_mode_mapping_from_history.png?raw=true)

##### **Columns Mapped**
![app_gui_dark_mode_columns_mapped](assets/app_screenshots/app_gui_dark_mode_columns_mapped.png?raw=true)

##### **Data Preview in Auxiliary Tab**
![app_gui_dark_mode_data_preview](assets/app_screenshots/app_gui_dark_mode_columns_mapped.png?raw=true)

##### **Copying Data From Source File to Template Destination File After Columns Mapping**
![app_gui_dark_mode_coppy_mapped_columns](assets/app_screenshots/app_gui_dark_mode_coppy_mapped_columns.png?raw=true)

##### **Choosing a Directory to Save the Template File With Copied Data**
![app_gui_dark_mode_saving_the_file_with_mapped_columns](assets/app_screenshots/app_gui_dark_mode_saving_the_file_with_mapped_columns.png?raw=true)

##### **Data Copied and Saved With Success**
![app_gui_dark_mode_data_copied_sucess](assets/app_screenshots/app_gui_dark_mode_data_copied_sucess.png?raw=true)

##### **History File For Mapped Columns**
![history_file_for_mapped_columns](assets/app_screenshots/history_file_for_mapped_columns.png?raw=true)

##### **Used Source File**
![used_source_file](assets/app_screenshots/used_source_file.png?raw=true)

##### **Used Template Destination File**
![used_destination_file_template](assets/app_screenshots/used_destination_file_template.png?raw=true)

##### **Template Destination File After Copy and Saving the Mapped Columns**
![destination_file_template_updated](assets/app_screenshots/destination_file_template_updated.png?raw=true)


## 🚀 Installation

### Prerequisites
- Python 3.12 or higher
- pip package manager

### Quick Setup

1. **Clone the repository**
```bash
git clone <repository-url>
cd excel_columns_mapper
```

2. **Create virtual environment (recommended)**
```bash
python -m venv .venv

# On Windows:
.venv\Scripts\activate

# On macOS/Linux:
source .venv/bin/activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run the application**
```bash
cd app
python main.py
```

## 📁 Project Structure

```
excel_columns_mapper/
├── app/
│   ├── main.py                    # Main application entry point
│   └── config.json               # Application configuration
├── assets/                       # Application assets (logos, icons)
├── data/                        # Sample and test data files
├── log/                         # Application logs and history
│   └── mapping_history.csv     # Automatic mapping history
├── themes/
│   └── azure.tcl               # Azure theme definition
├── tests/
│   └── run_test_simple.py            # Simple test runner for ExcelColumnMapper
│   └── test_config.tcl               # Test configuration for ExcelColumnMapper
│   └── test_excel_column_mapper.py   # Unit tests for ExcelColumnMapper application
│   └── TESTING_README.md             # Testing documentation
├── requirements.txt            # Python dependencies
├── config.json               # Global configuration
├── LICENSE                   # MIT License
└── README.md                # This documentation
```

## 🎯 Usage Guide

### 1. **Launch Application**
Run `python main.py` from the `app` directory to start the application.

### 2. **Select Files**
- **Source File**: Click "Browse" to select your data source Excel/CSV file
- **Destination File**: Click "Browse" to select your target template file

### 3. **Load Column Headers**
- Click "Load Column Headers" to analyze both files
- View source columns with sample data in the left panel
- See destination columns available for mapping in the right panel

### 4. **Configure Mappings**
- Use dropdown menus to map source columns to destination columns
- Visual indicators show mapping status with checkmarks (✓)
- Preview mappings in the "Data Preview" tab

### 5. **Transfer Data**
- Click "Copy Mapped Data" to execute the transfer
- Choose output location and filename
- Review confirmation dialog before proceeding

### 6. **Manage History**
- **View History**: Browse all previous mapping operations
- **Load from History**: Restore compatible previous configurations
- **Clear Mappings**: Reset current session

## 🔧 Interface Overview

### Header Section
- **Application Title**: Excel Column Mapper
- **Theme Toggle**: Switch between light/dark modes (🌙/☀)
- **Statistics Panel**: Real-time usage statistics

### Control Panel
- **File Selection**: Browse and select source/destination files
- **Action Buttons**: Load headers, copy data, manage mappings
- **Progress Tracking**: Visual progress bars for operations

### Main Content Area

#### Column Mapping Tab
- **Source Columns**: Tree view with sample data preview
- **Mapping Interface**: Dropdown selectors for column mapping
- **Visual Indicators**: Mapping status and direction arrows

#### Data Preview Tab
- **Mapping Summary**: Current column mapping configuration
- **Sample Data**: Preview of data transfer results
- **Validation**: Mapping compatibility checking

### Status Bar
- **Real-time Status**: Current operation status
- **Progress Indicator**: Operation completion percentage

## ⚙️ Configuration

### Application Settings
Edit `config.json` to customize:
- Default directories
- Theme preferences
- Statistics tracking
- History retention

### Theme Customization
The Azure theme can be customized by modifying `themes/azure.tcl`:
- Color schemes
- Font settings
- Widget styling
- Animation preferences

## 📊 Features Deep Dive

### Statistics Tracking
The application automatically tracks:
- Number of files processed
- Total mappings created
- Columns successfully mapped
- Session completion rates

### Mapping History
All operations are logged with:
- Timestamp and file information
- Complete mapping configurations
- Success/failure status
- Compatibility metadata

### Error Handling
Comprehensive error management for:
- Invalid file formats
- Missing or corrupted files
- Memory limitations
- Permission issues
- Column name conflicts

## 🔍 Troubleshooting

### Common Issues

**Theme not loading properly**
- Ensure `themes/azure.tcl` exists in the correct location
- Check file permissions
- Restart the application

**Files not loading**
- Verify file format (.xlsx, .xls, .csv)
- Ensure files are not open in other applications
- Check read permissions

**Memory issues with large files**
- Close unnecessary applications
- Use 64-bit Python for large datasets
- Consider processing in smaller chunks

**Mapping history not saving**
- Verify write permissions in `log/` directory
- Check available disk space
- Ensure application has proper file access

### Performance Tips
- Keep files under 100MB for optimal performance
- Close preview tabs when working with large datasets
- Use CSV format for very large files
- Regularly clean mapping history

## 🤝 Contributing

We welcome contributions! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

### Development Setup
```bash
# Clone your fork
git clone <your-fork-url>
cd excel_columns_mapper

# Create development environment
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows

# Install development dependencies
pip install -r requirements.txt

# Run the application
cd app
python main.py
```

## 📋 Requirements

### Python Packages
- `pandas` - Data manipulation and analysis
- `tkinter` - GUI framework (usually included with Python)
- `openpyxl` - Excel file handling
- `pathlib` - Path manipulation
- `json` - Configuration management
- `datetime` - Timestamp handling

### System Requirements
- **Operating System**: Windows 10+, macOS 10.14+, or Linux
- **Python**: 3.12 or higher
- **Memory**: 4GB RAM minimum (8GB recommended for large files)
- **Storage**: 100MB free space

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ Version History

- **v1.1** - Current release with Azure theme and enhanced UI
- **v1.0** - Initial release with basic mapping functionality

## 📧 Support

For support, feature requests, or bug reports:
- Create an issue in the repository
- Check existing documentation
- Review troubleshooting section

---

**Author**: Open Source Community  
**Version**: 1.1  
**Last Updated**: August 2025  
**Compatibility**: Python 3.12+, Cross-platform