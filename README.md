# Excel File Merger

A professional-grade GUI application for merging Excel files with advanced column mapping, duplicate handling, and flexible merge options. Built with Python and Tkinter.

## Features

### Core Functionality
- **Multi-Sheet Support**: Select specific sheets from each Excel file
- **Flexible Header Row Selection**: Specify which row contains column headers (1-10)
- **Advanced Column Mapping**: Map columns between files with different names
- **Multiple Merge Types**: Inner, outer, left, and right joins
- **Intelligent Duplicate Handling**: Keep all, first, last, or flag duplicates
- **Superior File Selection**: Choose which file's values take precedence for conflicts

### Key Capabilities
- **Visual Column Mapping Interface**: Easy-to-use drag-and-drop style column matching
- **Dynamic Merge Options**: Labels update based on superior file selection
- **Comprehensive Merge Statistics**: View detailed information about the merge results
- **One-Click File Opening**: Launch merged Excel file directly from success dialog
- **Source Tracking**: Automatically tracks which file(s) each row came from

## Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Required Packages
pip install pandas openpyxl

### Download
git clone https://github.com/Hondaguy900/ExcelFileMergerApp.git

cd ExcelFileMergerApp

## Usage

### Quick Start

1. **Launch the application**:
   - python ExcelFileMergerApp.py

2. **Select your files**:
   - Click "Browse..." to select two Excel files
   - Click "Load Files" to load available sheets

3. **Configure sheets and headers**:
   - Select the sheet to use from each file
   - Specify which row contains the column headers (default: 1)
   - Click "Load Columns" to populate the column mapping interface

4. **Map your columns**:
   - Check the "Use for Matching" box for columns you want to merge on
   - Map columns from File 1 to corresponding columns in File 2
   - Add additional mapping rows as needed

5. **Configure merge options**:
   - **Duplicate Handling**: Choose how to handle duplicate rows
   - **Superior File**: Select which file's values take precedence
   - **Merge Type**: Choose the type of join operation

6. **Merge and save**:
   - Click "Browse..." under Output File to choose save location
   - Click "Merge Files" to execute the merge
   - Review statistics and optionally open the file directly

## Merge Options Explained

### Duplicate Handling
- **Keep all duplicates**: Retain every row, including duplicates
- **Keep first occurrence**: Remove duplicate rows, keeping only the first instance
- **Keep last occurrence**: Remove duplicate rows, keeping only the last instance
- **Flag duplicates**: Add `duplicate_count` and `is_duplicate` columns to identify duplicates

### Superior File
When columns exist in both files, the superior file's values take precedence. This setting also affects how merge type labels are displayed.

### Merge Types
The merge type determines which rows are included in the final output:

- **Keep all rows from both files** (Outer Join): Include every row from both files, even if no match exists
- **Keep all rows from [superior file]** (Left/Right Join): Include all rows from the superior file, plus matching rows from the other file
- **Keep all rows from [inferior file]** (Right/Left Join): Include all rows from the inferior file, plus matching rows from the other file
- **Only keep rows that match in both files** (Inner Join): Include only rows with matching values in both files

## Output Details

The merged Excel file includes:
- All original columns from both files (with suffixes `_file1` and `_file2` for duplicates)
- A `source` column indicating the origin of each row:
  - "File 1 Only": Row exists only in first file
  - "File 2 Only": Row exists only in second file
  - "Both Files": Row found in both files (matched)
- Optional duplicate tracking columns (when "Flag duplicates" is selected)

## Use Cases

### Example 1: Updating Customer Database
Merge an existing customer database with new customer information:
- **Superior File**: New data (File 2)
- **Merge Type**: Keep all rows from both files
- **Duplicate Handling**: Keep last occurrence
- **Result**: Updated database with new customers added and existing customers updated

### Example 2: Inventory Reconciliation
Compare warehouse inventory with ordering system:
- **Superior File**: Warehouse data (File 1)
- **Merge Type**: Only keep rows that match in both files
- **Duplicate Handling**: Flag duplicates
- **Result**: Identify discrepancies between systems

### Example 3: Combining Regional Sales Data
Merge sales data from different regions:
- **Merge Type**: Keep all rows from both files
- **Duplicate Handling**: Keep all duplicates
- **Result**: Complete sales database across all regions

## Technical Details

### Built With
- **Python 3.x**: Core programming language
- **Tkinter**: GUI framework (included with Python)
- **Pandas**: Data manipulation and Excel processing
- **openpyxl**: Excel file reading/writing engine

### File Support
- Excel 2007+ (.xlsx)
- Excel 97-2003 (.xls)

### Performance Considerations
- Handles files with thousands of rows efficiently
- Memory usage scales with file size
- Recommended maximum: ~100,000 rows per file for optimal performance

## Troubleshooting

### Common Issues

**"Failed to load Excel files"**
- Ensure files are valid Excel formats (.xlsx or .xls)
- Check that files aren't open in Excel
- Verify files aren't corrupted

**"Failed to load columns"**
- Confirm sheets are selected for both files
- Verify header row number is correct
- Check that sheets contain data

**"Please select at least one column mapping"**
- Check at least one "Use for Matching" checkbox
- Ensure both columns are selected in the mapping row
- Verify column names exist in the dropdown

## Version History

### Version 2025.04.10.001
- Initial public release
- Complete column mapping interface
- Dynamic merge type labels
- Comprehensive duplicate handling
- Source tracking functionality

## Credits

**Developer**: Hondaguy900  
**AI Assistant**: Claude (Anthropic)

Built with extensive collaboration between human expertise and AI assistance to create a robust, user-friendly data merging solution.

## License

This project is available for use under standard open-source practices. Feel free to modify and distribute with attribution.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

### Potential Enhancements
- Support for CSV files
- Batch merging (multiple file pairs)
- Custom column renaming in output
- Advanced filtering options
- Merge operation history/logging
- Command-line interface option

## Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Provide sample files (sanitized) when reporting bugs
- Include error messages and steps to reproduce

---

**Note**: This tool is designed for data professionals, analysts, and anyone who regularly works with Excel data and needs powerful, flexible merging capabilities beyond Excel's built-in features.
