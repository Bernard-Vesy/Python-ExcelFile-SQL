"""
Simple usage examples for Excel file manipulation.
Run this script to see basic operations.
"""

from excel_manipulator import ExcelManipulator

def simple_example():
    """A simple example to get started quickly."""
    
    print("Creating a simple Excel file...")
    
    # Create a new Excel file
    excel = ExcelManipulator("simple_example.xlsx")
    excel.create_workbook()
    
    # Write headers
    excel.write_row(1, ["Product", "Price", "Quantity", "Total"])
    
    # Format headers
    for col in range(1, 5):
        excel.format_cell(1, col, font_bold=True, bg_color="4472C4")
    
    # Write product data
    products = [
        ["Laptop", 1200, 5],
        ["Mouse", 25, 20],
        ["Keyboard", 75, 15],
        ["Monitor", 300, 8]
    ]
    
    for idx, product in enumerate(products, start=2):
        excel.write_row(idx, product)
        # Add formula for total (Price * Quantity)
        excel.add_formula(idx, 4, f"=B{idx}*C{idx}")
    
    # Add total row
    excel.write_data(6, 3, "Grand Total:")
    excel.format_cell(6, 3, font_bold=True)
    excel.add_formula(6, 4, "=SUM(D2:D5)")
    excel.format_cell(6, 4, font_bold=True, bg_color="FFC000")
    
    # Save the file
    excel.save_workbook()
    
    print("\n✓ File 'simple_example.xlsx' created successfully!")
    print("\nTo read the file:")
    print("  excel = ExcelManipulator('simple_example.xlsx')")
    print("  excel.load_workbook()")
    print("  data = excel.read_all_data()")


if __name__ == "__main__":
    simple_example()
