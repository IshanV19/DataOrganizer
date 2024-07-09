# SubTables.py

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def process_data_file(file_path, output_directory):
    """
    Process a data file (either .xlsx or .csv), organize data into separate tables for each assay,
    and save the processed Excel file in the output directory. Recovery values outside 80-120 are highlighted.

    Args:
        file_path (str): Path to the input data file (either .xlsx or .csv).
        output_directory (str): Directory where output Excel files will be saved.

    Returns:
        tuple: Tuple containing (excel_file, success), where:
            - excel_file (str): Path to the output Excel file.
            - success (bool): True if processing was successful, False otherwise.
    """
    filename = os.path.basename(file_path)

    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets into a dictionary of DataFrames
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path, skiprows=[0])  # Skip the first row because of Plate Name
            # Convert % Recovery column to numeric
            df['% Recovery'] = pd.to_numeric(df['% Recovery'], errors='coerce')  # coerce will handle non-numeric values
        else:
            raise ValueError("Unsupported file format. Supported formats are .xlsx and .csv.")

        # Initialize an Excel writer object for the organized data file
        excel_file = os.path.join(output_directory, os.path.splitext(filename)[0] + '_organized_data.xlsx')
        if os.path.exists(excel_file):
            print(f"Excel file '{excel_file}' already exists. Skipping processing for '{filename}'.")
            return excel_file, True

        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:  # Use openpyxl engine explicitly
            for assay in df['Assay'].unique():
                assay_df = df[df['Assay'] == assay]

                # Aggregate data for each assay
                table_df = assay_df.groupby('Sample').agg(
                    Assay=('Assay', 'first'),
                    Calc_Concentration_1=('Calc. Concentration', lambda x: x.iloc[0] if len(x) >= 1 else None),
                    Calc_Concentration_2=('Calc. Concentration', lambda x: x.iloc[1] if len(x) >= 2 else None),
                    Avg_Calc_Conc=('Calc. Concentration', 'mean'),
                    Std_Dev_Calc_Conc=('Calc. Concentration', lambda x: x.std() / x.mean() * 100 if len(x) > 1 else None),
                    Recovery_1=('% Recovery', lambda x: x.iloc[0] if len(x) >= 1 else None),
                    Recovery_2=('% Recovery', lambda x: x.iloc[1] if len(x) >= 2 else None),
                    Avg_Recovery=('% Recovery', 'mean'),
                    Std_Dev_Recovery=('% Recovery', lambda x: x.std() / x.mean() * 100 if len(x) > 1 else None),
                ).reset_index()

                # Write to Excel sheet for current assay
                table_df.to_excel(writer, sheet_name=assay, index=False)

                # Apply conditional formatting to % Recovery columns
                wb = writer.book
                ws = wb[assay]  # Get worksheet by name

                red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')  # Red fill color

                # Function to apply conditional formatting
                def apply_conditional_formatting(cell):
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        value = float(cell.value)
                        if value < 80 or value > 120:
                            cell.fill = red_fill

                # Loop through relevant cells and apply formatting
                for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column, max_row=ws.max_row):
                    for cell in row:
                        if cell.column_letter in ['G', 'H']:  # Adjust to match your Recovery columns
                            apply_conditional_formatting(cell)

        print(f"Excel file '{excel_file}' created successfully with assay results organized and formatted.")

        return excel_file, True

    except Exception as e:
        print(f"An error occurred while processing '{filename}': {e}")
        return None, False