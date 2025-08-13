# Excel Column Mapper & Data Transfer

A Python GUI application that enables easy column mapping and data transfer between Excel files. This tool simplifies the process of copying data from one Excel file to another by providing an intuitive interface for mapping columns between source and destination files.

## Features

### 🔄 **Column Mapping**
- Interactive mapping between source and destination Excel file columns
- Visual column mapping with dropdown selection
- Search functionality for quick column finding
- Real-time mapping validation and preview

### 📊 **Data Preview**
- Display sample data from source columns
- Visual confirmation of column mappings
- Source column checkbox-style interface with mapping indicators

### 📝 **Mapping History**
- Automatic saving of all mapping configurations
- Load previous mappings from history
- View complete mapping history with timestamps
- Reuse successful mapping configurations

### 🔍 **Smart Search**
- Type-ahead search in column mapping dropdowns
- Case-insensitive column matching
- Filtered results as you type

### 💾 **Data Transfer**
- Copy mapped data from source to destination structure
- Handle files with different row counts
- Save results to new Excel files
- Preserve original file integrity

## Requirements

- Python 3.12+
- numpy
- pandas
- tkinter (usually included with Python)
- pytz
- six
- tzdata

## Installation

1. **Clone or download the repository**
```bash
git clone <repository-url>
cd excel-column-mapper
```

2. **Install required dependencies**
```bash
pip install -r requirements.txt
```

3. **Run the application**
```bash
python main.py
```

## Directory Structure

```
excel-column-mapper/
├── app/
    └── main.py                                 # Main application file
├── data/                                       # Recommended folder for source Excel files
        ├── source_files
            └── source_file.xlsx
        ├── destination_file
            └── destination_file_template.xlsx
├── log/                                        # for mapping history
│   └── mapping_history.csv                     # Automatic mapping history log
└── README.md                                   # This file
```

## Usage Guide

### 1. **Load Excel Files**
- **Source File**: Click "Browse" next to "Source File" to select your data source Excel file
- **Destination File**: Click "Browse" next to "Destination File" to select your target Excel file template

### 2. **Load Column Headers**
- Click "Load Column Headers" to analyze both files
- View source columns with data samples in the left panel
- See destination columns available for mapping in the right panel

### 3. **Configure Column Mappings**
- For each destination column, select the corresponding source column from the dropdown
- Use the search functionality by typing in the dropdown to filter options
- Mapped columns will show checkmarks (☑) in the source panel
- Arrow indicators (←) show the mapping direction

### 4. **Transfer Data**
- Click "Copy Mapped Data" to execute the transfer
- Choose where to save the updated file
- Review the confirmation dialog showing all mappings

### 5. **History Management**
- **View Mapping History**: See all previous mapping operations
- **Load from History**: Reuse a previous mapping configuration
- **Clear Mappings**: Reset all current mappings

## Interface Components

### Source Panel (Left)
- **Column List**: Shows all source file columns with checkboxes
- **Data Samples**: Preview actual data from each column
- **Mapping Indicators**: Visual confirmation of active mappings

### Mapping Panel (Right)
- **Destination Columns**: List of target file columns
- **Dropdown Selectors**: Choose source column for each destination
- **Search Functionality**: Type to filter available source columns

### Control Buttons
- **Load Column Headers**: Initialize file analysis
- **Copy Mapped Data**: Execute the data transfer
- **Clear Mappings**: Reset all mappings
- **Load from History**: Restore previous configuration
- **View Mapping History**: Browse past operations

## Advanced Features

### Mapping History
The application automatically maintains a detailed history of all mapping operations in `log/mapping_history.csv`. This includes:
- Timestamp of operation
- Source and destination file names
- Output file name
- Complete column mapping details

### Data Handling
- **Variable Row Counts**: Automatically handles source files with more rows than destination templates
- **Data Type Preservation**: Maintains original data types during transfer
- **Error Recovery**: Comprehensive error handling with user-friendly messages

### Search and Filter
- **Real-time Search**: Type in any mapping dropdown to filter source columns
- **Case-insensitive**: Matching works regardless of letter case
- **Partial Matching**: Find columns with partial name matches

## File Format Support

- **Input**: Excel files (.xlsx, .xls)
- **Output**: Excel files (.xlsx)
- **Compatibility**: Works with files created in Excel, Google Sheets, LibreOffice Calc, etc.

## Best Practices

1. **Organize Files**: Keep source files in the `source_files/` directory for easy access
2. **Backup Originals**: The tool preserves original files, but always keep backups
3. **Test Mappings**: Review sample data before executing large transfers
4. **Use History**: Save time by reusing successful mapping configurations
5. **Verify Results**: Always check the output file to ensure data transferred correctly

## Error Handling

The application includes comprehensive error handling for:
- Invalid file formats
- Missing files
- Corrupted Excel files
- Memory limitations with large files
- Column name mismatches
- Write permission issues

## Troubleshooting

### Common Issues

**"Failed to load files" error**
- Ensure Excel files are not open in other applications
- Check file permissions
- Verify file format (.xlsx or .xls)

**"No mapping history found"**
- The `log/` directory is created automatically after first use
- Mapping history is only saved after successful data transfers

**Dropdown not showing all columns**
- Clear the dropdown and click again to refresh the list
- Check if search filtering is active

**Memory issues with large files**
- Close other applications to free up memory
- Consider processing large files in smaller chunks

## Contributing

Feel free to contribute to this project by:
- Make a fork and be happy
- Reporting bugs
- Suggesting new features
- Improving documentation


**Version**: 1.0  
**Last Updated**: August 13rd 2025  
**Compatibility**: Python 3.12+, Windows/macOS/Linux