import xlwings as xw
import os
import pandas as pd

def excel_sheets_to_csvs(excel_path, output_dir=None):
    # Open the workbook without displaying Excel
    app = xw.App(visible=False)
    wb = app.books.open(excel_path)

    # Create output directory if not provided
    if output_dir is None:
        output_dir = os.path.splitext(excel_path)[0] + "_csvs"
    os.makedirs(output_dir, exist_ok=True)

    try:
        for sheet in wb.sheets:
            # Read all values as a DataFrame
            data = sheet.used_range.options(pd.DataFrame, header=1, index=False).value
            # Construct the CSV file path
            csv_filename = f"{sheet.name}.csv"
            csv_path = os.path.join(output_dir, csv_filename)
            # Save to CSV
            data.to_csv(csv_path, index=False)
            print(f"Saved {csv_filename}")
    finally:
        wb.close()
        app.quit()

# Example usage
excel_file_path = "data_powerbi_banking.xlsx"  # Replace with your Excel file path
excel_sheets_to_csvs(excel_file_path)
