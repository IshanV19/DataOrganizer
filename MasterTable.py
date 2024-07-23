#MasterTable.py

import os
import pandas as pd
from datetime import datetime
import re

def extract_positive_time_unit(sample):
    """Extracts positive time units from the sample string."""
    match = re.search(r'([dwm])(\d+)', sample) 
    if match:
        unit = match.group(1)
        value = int(match.group(2))

        if unit == 'd':
            return f"{value}"
        elif unit == 'w':
            return f"{value * 7}" 
        elif unit == 'm':
            return f"{value * 30}"
    return None

def extract_negative_time_unit(sample):
    """Extracts negative time units from the sample string."""
    match = re.search(r'([dwm])-?(\d+)', sample) 
    if match:
        unit = match.group(1)
        value = int(match.group(2))

        if unit == 'd':
            return f"{-value}"
        elif unit == 'w':
            return f"{-value * 7}"
        elif unit == 'm':
            return f"{-value * 30}"
    return None

def extract_time_unit(sample):
    """Extracts time units from the sample string, handling both positive and negative cases."""
    if re.search(r'[dwm]-\d+', sample):
        return extract_negative_time_unit(sample)
    elif re.search(r'[dwm]\d+', sample):
        return extract_positive_time_unit(sample)
    return None

def trim_sample_name(sample):
    """Trim parts of the sample name that start with 'm', 'w', or 'd'."""
    if isinstance(sample, str):
        parts = sample.split()
        trimmed_parts = [part for part in parts if not (part.startswith('m') or part.startswith('w') or part.startswith('d'))]
        return ' '.join(trimmed_parts)
    return sample

def combine_and_organize(input_directory, output_directory, processed_files=None):
    
    if processed_files is None:
        processed_files = []

    try:
        # Ensure output directory exists
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        
        # Get all organized data files
        organized_files = [f for f in os.listdir(input_directory) if f.endswith('MSDProinflam_organized_data.xlsx') and f not in processed_files]
        
        print(f"Found organized files: {organized_files}")  # Debugging line
        
        if not organized_files:
            print(f"No new organized data files found in '{input_directory}'. Exiting.")
            return
        
        master_dfs = {
            '101': pd.DataFrame(),
            '201': pd.DataFrame(),
            '301': pd.DataFrame()
        }

        for file in organized_files:
            file_path = os.path.join(input_directory, file)
            df = pd.read_excel(file_path, sheet_name=None)  

            for sheet_name, sheet_df in df.items():
                if 'Sample' in sheet_df.columns:
                    sheet_df['Days'] = sheet_df['Sample'].apply(extract_time_unit)
                    sheet_df['Sample'] = sheet_df['Sample'].apply(trim_sample_name)

                    sample_group = None
                    if sheet_df['Sample'].astype(str).str.startswith('101').any():
                        sample_group = '101'
                    elif sheet_df['Sample'].astype(str).str.startswith('201').any():
                        sample_group = '201'
                    elif sheet_df['Sample'].astype(str).str.startswith('301').any():
                        sample_group = '301'

                    if sample_group:
                        sheet_df = sheet_df.copy()  
                        sheet_df['Data Source'] = file  
                        sheet_df['Time Added'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Timestamp of when added

                        master_dfs[sample_group] = pd.concat([master_dfs[sample_group], sheet_df], ignore_index=True)

            # Mark this file as processed
            processed_files.append(file)

        # Remove specified columns and save each master_df to its own Excel file
        columns_to_remove = ['Recovery_1', 'Recovery_2', 'Avg_Recovery', 'Std_Dev_Recovery']
        for sample_group, master_df in master_dfs.items():
            if not master_df.empty:
                # Remove rows starting with C0 or S0 in 'Sample' column
                master_df = master_df[~master_df['Sample'].astype(str).str.startswith(('C0', 'S0'))].copy()

                # Remove specified columns
                master_df.drop(columns=columns_to_remove, errors='ignore', inplace=True)

                # Save to Excel with specified sheet name
                master_table_file = os.path.join(output_directory, f'{sample_group}_master_table.xlsx')
                with pd.ExcelWriter(master_table_file) as writer:
                    master_df.to_excel(writer, sheet_name=f'{sample_group}_master', index=False)
                print(f"Master table for sample group '{sample_group}' updated successfully at '{master_table_file}'.")

    except Exception as e:
        print(f"An error occurred while combining data into master tables: {e}")
