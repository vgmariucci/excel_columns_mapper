# Excel Column Mapper & Data Transfer Tool

A modern, user-friendly Python GUI application for mapping and transferring data between Excel files. Built with a beautiful Azure-inspired theme and comprehensive functionality for seamless data migration.

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
**Version**: 1.0  
**Last Updated**: August 2025  
**Compatibility**: Python 3.12+, Cross-platform