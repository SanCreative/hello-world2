#!/usr/bin/env python3
"""
Excel File Reader
A Python script to read and display Excel files from your hard drive.
"""

import pandas as pd
import os
import sys
from pathlib import Path


class ExcelReader:
    def __init__(self):
        self.df = None
        self.file_path = None
    
    def read_excel_file(self, file_path, sheet_name=None):
        """
        Read an Excel file and load it into a pandas DataFrame.
        
        Args:
            file_path (str): Path to the Excel file
            sheet_name (str, optional): Name of the sheet to read. If None, reads the first sheet.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Convert to Path object for better handling
            self.file_path = Path(file_path)
            
            # Check if file exists
            if not self.file_path.exists():
                print(f"❌ Error: File '{file_path}' not found!")
                return False
            
            # Check if it's an Excel file
            if not self.file_path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
                print(f"❌ Error: '{file_path}' is not a valid Excel file!")
                return False
            
            print(f"📖 Reading Excel file: {self.file_path.name}")
            
            # Read the Excel file
            if sheet_name:
                self.df = pd.read_excel(file_path, sheet_name=sheet_name)
                print(f"📊 Successfully loaded sheet: '{sheet_name}'")
            else:
                self.df = pd.read_excel(file_path)
                print(f"📊 Successfully loaded first sheet")
            
            return True
            
        except Exception as e:
            print(f"❌ Error reading Excel file: {str(e)}")
            return False
    
    def get_sheet_names(self, file_path):
        """
        Get all sheet names from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file
        
        Returns:
            list: List of sheet names
        """
        try:
            xl_file = pd.ExcelFile(file_path)
            return xl_file.sheet_names
        except Exception as e:
            print(f"❌ Error getting sheet names: {str(e)}")
            return []
    
    def display_basic_info(self):
        """Display basic information about the loaded Excel data."""
        if self.df is None:
            print("❌ No data loaded. Please read an Excel file first.")
            return
        
        print("\n" + "="*60)
        print("📊 EXCEL FILE INFORMATION")
        print("="*60)
        print(f"📁 File: {self.file_path.name}")
        print(f"📏 Dimensions: {self.df.shape[0]} rows × {self.df.shape[1]} columns")
        print(f"📋 Columns: {list(self.df.columns)}")
        print(f"💾 Memory usage: {self.df.memory_usage(deep=True).sum() / 1024:.2f} KB")
        
        # Data types
        print(f"\n📝 Data Types:")
        for col, dtype in self.df.dtypes.items():
            print(f"  • {col}: {dtype}")
    
    def display_data(self, rows=10, show_all=False):
        """
        Display the Excel data in a formatted table.
        
        Args:
            rows (int): Number of rows to display (default: 10)
            show_all (bool): If True, display all data (default: False)
        """
        if self.df is None:
            print("❌ No data loaded. Please read an Excel file first.")
            return
        
        print("\n" + "="*60)
        print("📋 EXCEL DATA")
        print("="*60)
        
        if show_all or len(self.df) <= rows:
            print(self.df.to_string(index=True))
        else:
            print(f"📄 Showing first {rows} rows:")
            print(self.df.head(rows).to_string(index=True))
            
            if len(self.df) > rows:
                print(f"\n... and {len(self.df) - rows} more rows")
    
    def display_summary(self):
        """Display statistical summary of the data."""
        if self.df is None:
            print("❌ No data loaded. Please read an Excel file first.")
            return
        
        print("\n" + "="*60)
        print("📊 DATA SUMMARY")
        print("="*60)
        
        # For numeric columns
        numeric_cols = self.df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            print("🔢 Numeric columns summary:")
            print(self.df[numeric_cols].describe().to_string())
        
        # For text columns
        text_cols = self.df.select_dtypes(include=['object']).columns
        if len(text_cols) > 0:
            print(f"\n📝 Text columns summary:")
            for col in text_cols:
                unique_count = self.df[col].nunique()
                print(f"  • {col}: {unique_count} unique values")
    
    def search_data(self, search_term, column=None):
        """
        Search for data containing a specific term.
        
        Args:
            search_term (str): Term to search for
            column (str, optional): Specific column to search in
        """
        if self.df is None:
            print("❌ No data loaded. Please read an Excel file first.")
            return
        
        print(f"\n🔍 Searching for '{search_term}'...")
        
        if column:
            if column not in self.df.columns:
                print(f"❌ Column '{column}' not found!")
                return
            
            mask = self.df[column].astype(str).str.contains(search_term, case=False, na=False)
            results = self.df[mask]
        else:
            # Search in all columns
            mask = self.df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
            results = self.df[mask]
        
        if len(results) > 0:
            print(f"✅ Found {len(results)} matching rows:")
            print(results.to_string(index=True))
        else:
            print("❌ No matches found.")


def main():
    """Main function to run the Excel reader."""
    reader = ExcelReader()
    
    print("🐍 Excel File Reader")
    print("=" * 50)
    
    # Get file path from user
    while True:
        file_path = input("\n📁 Enter the path to your Excel file: ").strip()
        
        if not file_path:
            print("❌ Please enter a valid file path.")
            continue
        
        # Handle quotes around the path
        file_path = file_path.strip('"').strip("'")
        
        # Check if file exists and read it
        if reader.read_excel_file(file_path):
            break
        else:
            retry = input("❓ Would you like to try another file? (y/n): ").lower()
            if retry != 'y':
                return
    
    # Check for multiple sheets
    sheet_names = reader.get_sheet_names(file_path)
    if len(sheet_names) > 1:
        print(f"\n📚 Available sheets: {sheet_names}")
        sheet_choice = input("📋 Enter sheet name (or press Enter for first sheet): ").strip()
        if sheet_choice and sheet_choice in sheet_names:
            reader.read_excel_file(file_path, sheet_choice)
    
    # Display options menu
    while True:
        print("\n" + "="*50)
        print("📋 DISPLAY OPTIONS")
        print("="*50)
        print("1. 📊 Show basic file information")
        print("2. 📋 Display data (first 10 rows)")
        print("3. 📄 Display all data")
        print("4. 📊 Show statistical summary")
        print("5. 🔍 Search data")
        print("6. 📁 Load different file")
        print("7. 🚪 Exit")
        
        choice = input("\n🎯 Enter your choice (1-7): ").strip()
        
        if choice == '1':
            reader.display_basic_info()
        
        elif choice == '2':
            rows = input("📏 Number of rows to display (default 10): ").strip()
            try:
                rows = int(rows) if rows else 10
            except ValueError:
                rows = 10
            reader.display_data(rows=rows)
        
        elif choice == '3':
            confirm = input("⚠️  This might display a lot of data. Continue? (y/n): ").lower()
            if confirm == 'y':
                reader.display_data(show_all=True)
        
        elif choice == '4':
            reader.display_summary()
        
        elif choice == '5':
            search_term = input("🔍 Enter search term: ").strip()
            if search_term:
                column = input("📋 Search in specific column (or press Enter for all): ").strip()
                column = column if column else None
                reader.search_data(search_term, column)
        
        elif choice == '6':
            main()
            return
        
        elif choice == '7':
            print("👋 Goodbye!")
            return
        
        else:
            print("❌ Invalid choice. Please enter 1-7.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Program interrupted. Goodbye!")
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {str(e)}")