# Excel Table Extractor

## Overview

The Excel Table Extractor is a Python script that automatically detects and extracts tables from an Excel file. It identifies header rows, aligns data correctly, and processes the tables into pandas DataFrames.

## Features

- Automatically detects header rows in Excel sheets.
- Aligns data rows correctly under detected headers.
- Processes and outputs tables as pandas DataFrames.
- Configurable keywords for identifying headers.

## Installation

1. Clone the repository or download the script file.
2. Install the required dependencies:
    ```bash
    pip install pandas
    ```

## Usage

1. Ensure you have an Excel file (`.xlsx`) that you want to process.
2. Update the `file_path` variable in the script to point to your Excel file.

### Example Usage

```python
from excel_table_extractor import ExcelTableExtractor

file_path = "your_excel_file.xlsx"
extractor = ExcelTableExtractor(file_path)
tables = extractor.extract_tables()

for i, table in enumerate(tables):
    print(f"Table {i + 1}:")
    print(table)
    print("\n")
