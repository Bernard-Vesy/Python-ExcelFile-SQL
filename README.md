# Python-ExcelFile-SQL

A comprehensive Python application for manipulating Excel files and integrating them with SQL databases.

## Overview

This project demonstrates how to manipulate Excel files using Python, including reading, writing, formatting, and SQL database integration. It provides a clean, object-oriented interface for common Excel operations.

## Features

- **Create and Write**: Create new Excel workbooks and write data to cells, rows, and columns
- **Read Operations**: Read data from existing Excel files, individual cells or entire worksheets
- **Formatting**: Apply formatting to cells (fonts, colors, alignment, etc.)
- **Formulas**: Add Excel formulas to cells for calculations
- **SQL Integration**: Import Excel data to SQL databases and export SQL query results to Excel
- **Pandas Integration**: Use pandas DataFrames for advanced data manipulation

## Requirements

- Python 3.7 or higher
- openpyxl
- pandas
- sqlalchemy

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Bernard-Vesy/Python-ExcelFile-SQL.git
cd Python-ExcelFile-SQL
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Excel Manipulation

```python
from excel_manipulator import ExcelManipulator

# Create a new Excel file
excel = ExcelManipulator("my_file.xlsx")
excel.create_workbook()

# Write data to cells
excel.write_data(1, 1, "Name")
excel.write_data(1, 2, "Age")

# Write an entire row
excel.write_row(2, ["John Doe", 30])

# Format cells
excel.format_cell(1, 1, font_bold=True, bg_color="CCCCCC")

# Add formulas
excel.add_formula(3, 2, "=B2*2")

# Save the workbook
excel.save_workbook()

# Read data from an existing file
excel2 = ExcelManipulator("my_file.xlsx")
excel2.load_workbook()
value = excel2.read_data(1, 1)
all_data = excel2.read_all_data()
```

### Excel and SQL Integration

```python
from excel_manipulator import ExcelSQLIntegration

# Convert Excel to pandas DataFrame
df = ExcelSQLIntegration.excel_to_dataframe("data.xlsx")

# Save DataFrame to Excel
ExcelSQLIntegration.dataframe_to_excel(df, "output.xlsx")

# Import Excel data to SQL database
ExcelSQLIntegration.excel_to_sql(
    excel_file="data.xlsx",
    table_name="employees",
    db_path="database.db"
)

# Export SQL query results to Excel
ExcelSQLIntegration.sql_to_excel(
    query="SELECT * FROM employees WHERE salary > 50000",
    excel_file="high_earners.xlsx",
    db_path="database.db"
)
```

### Running the Examples

Run the example script to see all features in action:
```bash
python excel_manipulator.py
```

This will create example Excel files demonstrating various operations.

## API Reference

### ExcelManipulator Class

Main class for Excel file manipulation.

#### Methods:

- `__init__(filename)`: Initialize with a filename
- `create_workbook()`: Create a new Excel workbook
- `load_workbook()`: Load an existing Excel workbook
- `save_workbook()`: Save the current workbook
- `write_data(row, col, value)`: Write data to a specific cell
- `read_data(row, col)`: Read data from a specific cell
- `write_row(row_num, data_list)`: Write a list of values to a row
- `write_column(col_num, data_list)`: Write a list of values to a column
- `format_cell(row, col, font_bold, font_size, bg_color)`: Apply formatting to a cell
- `read_all_data()`: Read all data from the worksheet
- `add_formula(row, col, formula)`: Add a formula to a cell

### ExcelSQLIntegration Class

Class for Excel and SQL database integration.

#### Static Methods:

- `excel_to_dataframe(filename, sheet_name)`: Convert Excel file to pandas DataFrame
- `dataframe_to_excel(df, filename, sheet_name, index)`: Save DataFrame to Excel
- `excel_to_sql(excel_file, table_name, db_path, sheet_name)`: Import Excel data to SQL
- `sql_to_excel(query, excel_file, db_path, sheet_name)`: Export SQL results to Excel

## Use Cases

1. **Data Entry and Reports**: Create formatted Excel reports programmatically
2. **Data Migration**: Move data between Excel and SQL databases
3. **Automated Processing**: Process Excel files in batch operations
4. **Data Analysis**: Combine Excel data with Python's data analysis capabilities
5. **Business Intelligence**: Generate dynamic reports from database queries

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License.

## Author

Bernard Vesy

## Acknowledgments

- openpyxl library for Excel file manipulation
- pandas for data manipulation and analysis
- SQLAlchemy for database connectivity