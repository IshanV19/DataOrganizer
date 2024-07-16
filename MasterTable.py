#MasterTable.py

import os
import pandas as pd
from datetime import datetime

def combine_and_organize(input_directory, master_table_file, processed_files=None):
    """
    Combine multiple organized data files into a single master table, avoiding duplication.

    Args:
        input_directory (str): Directory where organized data files are located.
        master_table_file (str): Path to the output Excel file for the master table.
        processed_files (list, optional): List of files already processed. Default is None.
    """
    if processed_files is None:
        processed_files = []

    try:
        # Get all organized data files
        organized_files = [f for f in os.listdir(input_directory) if f.endswith('MSDProinflam_organized_data.xlsx') and f not in processed_files]
        
        if not organized_files:
            print(f"No new organized data files found in '{input_directory}'. Exiting.")
            return
        
        print(f"Found organized data files: {organized_files}")
        
        # Initialize an empty DataFrame for the master table
        master_df = pd.DataFrame()

        # Iterate through each organized data file and append to master_df
        for file in organized_files:
            file_path = os.path.join(input_directory, file)
            print(f"Processing file: {file_path}")
            df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets into a dictionary of DataFrames

            # Process each sheet (assay) and append to master_df
            for sheet_name, sheet_df in df.items():
                if not master_df.empty:
                    already_added = (master_df['Data Source'] == file) & (master_df['Assay'] == sheet_name)
                    if already_added.any():
                        print(f"Skipping '{file}' for sheet '{sheet_name}' as it's already in the master table.")
                        continue

                # Add columns for data source and time added
                sheet_df = sheet_df.copy()  # Create a copy to avoid SettingWithCopyWarning
                sheet_df['Data Source'] = file  # Assuming 'file' is the source identifier
                sheet_df['Time Added'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Timestamp of when added

                # Append to master_df
                master_df = pd.concat([master_df, sheet_df], ignore_index=True)

            # Mark this file as processed
            processed_files.append(file)

        # Remove rows starting with C0 or S0 in 'Sample' column
        master_df = master_df[~master_df['Sample'].astype(str).str.startswith(('C0', 'S0'))]

        # Save master_df to Excel
        master_df.to_excel(master_table_file, index=False)
        print(f"Master table updated successfully at '{master_table_file}'.")

    except Exception as e:
        print(f"An error occurred while combining data into master table: {e}")
