# Excel File Reader 📊

A Python script that can read Excel files from your hard drive and display them on screen with various viewing options.

## Features ✨

- 📖 Read Excel files (.xlsx, .xls, .xlsm)
- 📚 Support for multiple sheets
- 📊 Display file information and data statistics
- 📋 View data in formatted tables
- 🔍 Search functionality within the data
- 📏 Customizable row display limits
- 💾 Memory usage information
- 🎯 Interactive menu system

## Installation 🚀

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the script:**
   ```bash
   python excel_reader.py
   ```

## Dependencies 📦

- `pandas` - For data manipulation and analysis
- `openpyxl` - For reading modern Excel files (.xlsx)
- `xlrd` - For reading legacy Excel files (.xls)

## Usage 💡

### Basic Usage

1. Run the script:
   ```bash
   python excel_reader.py
   ```

2. Enter the path to your Excel file when prompted

3. Choose from the menu options:
   - **Show basic file information** - File details, dimensions, columns, data types
   - **Display data** - View the actual data in table format
   - **Show statistical summary** - Descriptive statistics for numeric columns
   - **Search data** - Find specific values within the data
   - **Load different file** - Switch to another Excel file

### Using as a Module

You can also import and use the `ExcelReader` class in your own Python scripts:

```python
from excel_reader import ExcelReader

# Create reader instance
reader = ExcelReader()

# Read an Excel file
if reader.read_excel_file("path/to/your/file.xlsx"):
    # Display basic info
    reader.display_basic_info()
    
    # Show first 5 rows
    reader.display_data(rows=5)
    
    # Search for specific data
    reader.search_data("search_term", column="specific_column")
```

## File Path Examples 📁

### Windows:
```
C:\Users\YourName\Documents\data.xlsx
C:\path\to\your\file.xlsx
```

### macOS/Linux:
```
/home/username/Documents/data.xlsx
~/Documents/data.xlsx
./data.xlsx
```

## Supported File Formats 📄

- `.xlsx` - Excel 2007+ files
- `.xls` - Legacy Excel files
- `.xlsm` - Excel files with macros

## Menu Options 🎛️

1. **📊 Show basic file information**
   - File name and path
   - Number of rows and columns
   - Column names and data types
   - Memory usage

2. **📋 Display data (first 10 rows)**
   - Shows data in a formatted table
   - Customizable number of rows
   - Includes row indices

3. **📄 Display all data**
   - Shows complete dataset
   - Warning for large files

4. **📊 Show statistical summary**
   - Descriptive statistics for numeric columns
   - Unique value counts for text columns

5. **🔍 Search data**
   - Search across all columns or specific column
   - Case-insensitive search
   - Shows matching rows

6. **📁 Load different file**
   - Switch to another Excel file

7. **🚪 Exit**
   - Quit the program

## Error Handling 🛡️

The script includes comprehensive error handling for:
- File not found errors
- Invalid file formats
- Corrupted Excel files
- Missing sheets
- Memory issues with large files

## Tips 💡

- **Large Files**: For very large Excel files, use option 2 to view data in chunks rather than loading everything at once
- **Multiple Sheets**: The script will detect and let you choose between multiple sheets
- **File Paths**: You can drag and drop files in most terminals to get the file path
- **Search**: Use the search function to quickly find specific data within your Excel file

## Troubleshooting 🔧

### Common Issues:

1. **"No module named 'pandas'"**
   ```bash
   pip install pandas openpyxl xlrd
   ```

2. **"File not found"**
   - Check the file path is correct
   - Use absolute paths if relative paths don't work
   - Ensure the file exists and you have read permissions

3. **"Invalid Excel file"**
   - Verify the file has .xlsx, .xls, or .xlsm extension
   - Check if the file is corrupted

4. **Memory errors with large files**
   - Use the partial display options instead of showing all data
   - Consider processing the file in chunks

## Example Output 📝

```
🐍 Excel File Reader
==================================================

📁 Enter the path to your Excel file: data.xlsx
📖 Reading Excel file: data.xlsx
📊 Successfully loaded first sheet

==================================================
📋 DISPLAY OPTIONS
==================================================
1. 📊 Show basic file information
2. 📋 Display data (first 10 rows)
3. 📄 Display all data
4. 📊 Show statistical summary
5. 🔍 Search data
6. 📁 Load different file
7. 🚪 Exit

🎯 Enter your choice (1-7): 1

============================================================
📊 EXCEL FILE INFORMATION
============================================================
📁 File: data.xlsx
📏 Dimensions: 1000 rows × 5 columns
📋 Columns: ['Name', 'Age', 'City', 'Salary', 'Department']
💾 Memory usage: 125.50 KB

📝 Data Types:
  • Name: object
  • Age: int64
  • City: object
  • Salary: float64
  • Department: object
```
